"""
generate_certificates.py
Reads the Excel roster + config.yaml, produces one PDF certificate per student.
Works in two modes (set in config.yaml -> template_mode):
  - "docx"  : fills a Word template with {{ placeholders }} via docxtpl, converts to PDF
  - "image" : draws text on top of a flat image/design at fixed coordinates, saves as PDF

Usage:
    python generate_certificates.py
Output:
    output/pdf/<Student_Name>__<Cluster_ID>.pdf   (one per student)
    output/manifest.csv                            (student -> pdf path -> poc email, used by send_emails.py)
"""

import os
import re
import shutil
import subprocess
import sys

import pandas as pd
import yaml
from docxtpl import DocxTemplate
from PIL import Image, ImageDraw, ImageFont


def load_config(path="config.yaml"):
    with open(path, "r") as f:
        return yaml.safe_load(f)


def find_soffice():
    """Return the path to the LibreOffice soffice.exe (or soffice on Linux/Mac).

    Search order:
      1. Common Windows install locations (so this works without touching PATH).
      2. shutil.which() – covers PATH-based installs and Linux/Mac.

    Raises FileNotFoundError if LibreOffice cannot be located anywhere.
    """
    # Common Windows installation paths – checked in priority order.
    windows_candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in windows_candidates:
        if os.path.isfile(path):
            return path

    # Fall back to whatever is on PATH (works on Linux / macOS / custom Windows installs).
    found = shutil.which("soffice")
    if found:
        return found

    raise FileNotFoundError(
        "LibreOffice (soffice.exe) was not found.\n"
        "Checked:\n"
        + "\n".join(f"  {p}" for p in windows_candidates)
        + "\n  (PATH)\n"
        "Please install LibreOffice from https://www.libreoffice.org/download/"
    )


def safe_filename(s):
    """Strip characters that are unsafe in filenames."""
    return re.sub(r"[^A-Za-z0-9_\-]+", "_", str(s)).strip("_")


def read_excel_any(file_path, sheet_name=0):
    """
    Robustly read Excel files (.xlsx, .xls, XML Spreadsheet 2003, .csv)
    supporting standard openpyxl, legacy XML formats from web exports, and xlrd.
    """
    import xml.etree.ElementTree as ET

    # 1. Try standard pandas read_excel
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception:
        pass

    # 2. Check if file is an XML Spreadsheet (SpreadsheetML) often saved with .xls extension
    try:
        with open(file_path, "rb") as f:
            header = f.read(500)
        if b"<?xml" in header or b"mso-application" in header or b"Workbook" in header:
            tree = ET.parse(file_path)
            root = tree.getroot()
            ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
            rows = []
            for row in root.findall(".//ss:Table/ss:Row", ns):
                cells = []
                for cell in row.findall("ss:Cell", ns):
                    idx = cell.attrib.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                    if idx:
                        idx_num = int(idx) - 1
                        while len(cells) < idx_num:
                            cells.append("")
                    data = cell.find("ss:Data", ns)
                    cells.append(data.text.strip() if data is not None and data.text else "")
                rows.append(cells)
            if rows:
                max_len = max(len(r) for r in rows)
                padded_rows = [r + [""] * (max_len - len(r)) for r in rows]
                headers = [str(c) if c else f"Column_{i+1}" for i, c in enumerate(padded_rows[0])]
                df = pd.DataFrame(padded_rows[1:], columns=headers)
                return df.dropna(how="all")
    except Exception:
        pass

    # 3. Try xlrd engine for legacy binary .xls files
    try:
        return pd.read_excel(file_path, engine="xlrd", sheet_name=sheet_name)
    except Exception:
        pass

    # 4. Fallback to read_csv
    return pd.read_csv(file_path)



def get_active_cfg(cfg):
    at = cfg.get("active_type", "participation")
    return cfg.get("types", {}).get(at, cfg)


def load_roster(cfg):
    active_cfg = get_active_cfg(cfg)
    excel_path = active_cfg.get("excel_path", "")
    if not excel_path or not os.path.exists(excel_path):
        raise FileNotFoundError("No Excel file uploaded. Please upload your Excel sheet first.")
    df = read_excel_any(excel_path, sheet_name=active_cfg.get("excel_sheet_name", 0))
    col_map = active_cfg.get("columns", {})

    # Only validate columns that have a non-empty, non-NONE value in the mapping
    active_map = {key: real for key, real in col_map.items()
                  if real and str(real).strip() and str(real).strip().upper() != "NONE"}

    missing = [real for key, real in active_map.items() if real not in df.columns]
    if missing:
        raise ValueError(
            f"These columns are in config.yaml but NOT found in the Excel file: {missing}\n"
            f"Excel columns found: {list(df.columns)}"
        )

    # Rename to our internal standard names
    rename_map = {real: key for key, real in active_map.items()}
    df = df.rename(columns=rename_map)

    # Required fields based on active_type
    at = cfg.get("active_type", "participation")
    if at == "school":
        req = ["school", "poc_email"]
    elif at == "volunteer":
        req = ["volunteer_name", "poc_email"]
    else:
        req = ["student_name", "poc_email"]

    for r in req:
        if r not in df.columns:
            raise ValueError(f"Required field '{r}' missing — check the `columns:` mapping in config.yaml")

    # If volunteer type, ensure student_name and volunteer_email mirror volunteer_name and poc_email
    if at == "volunteer":
        df["student_name"] = df["volunteer_name"]
        df["volunteer_email"] = df["poc_email"]
        if "poc_name" not in df.columns or df["poc_name"].isna().all():
            df["poc_name"] = df["volunteer_name"]

    # school: optional certificate field — blank if not mapped
    if "school" not in df.columns and "school" not in req:
        df["school"] = ""

    # event_name: optional — blank if not mapped
    if "event_name" not in df.columns and "event_name" not in req:
        df["event_name"] = ""

    # cluster_id: always derived from poc_email prefix — not user-mapped
    # Different poc_email = different cluster/POC group
    df["cluster_id"] = df["poc_email"].apply(lambda e: str(e).split("@")[0])

    # poc_name: use custom mapping if present, otherwise fallback to cluster_id (email prefix)
    if "poc_name" in df.columns:
        df["poc_name"] = df["poc_name"].fillna("").astype(str).str.strip()
        df.loc[df["poc_name"] == "", "poc_name"] = df["cluster_id"]
    else:
        df["poc_name"] = df["cluster_id"]

    blank_email = df[df["poc_email"].isna() | (df["poc_email"].astype(str).str.strip() == "")]
    if len(blank_email):
        print(f"WARNING: {len(blank_email)} row(s) have a blank POC email and will be skipped:")
        print(blank_email[[ "school" if "school" in blank_email.columns else "volunteer_name" if "volunteer_name" in blank_email.columns else "student_name" ]].to_string(index=False))
        df = df.drop(blank_email.index)

    return df


# ---------------- DOCX MODE ----------------

def generate_docx_mode(df, cfg):
    active_cfg = get_active_cfg(cfg)
    out_dir = cfg.get("output_dir", "output")
    docx_dir = os.path.join(out_dir, "docx")
    pdf_dir = os.path.join(out_dir, "pdf")
    os.makedirs(docx_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)

    template_path = active_cfg.get("docx_template_path", "")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"docx_template_path not found: {template_path}")

    manifest_rows = []

    for _, row in df.iterrows():
        context = {k: ("" if pd.isna(v) else v) for k, v in row.to_dict().items()}
        fname_base = safe_filename(row.get('student_name', row.get('school', 'cert')))

        doc = DocxTemplate(template_path)
        doc.render(context)
        docx_out = os.path.join(docx_dir, fname_base + ".docx")
        doc.save(docx_out)

        manifest_rows.append({
            "student_name": row.get("student_name", ""),
            "school": row.get("school", ""),
            "event_name": row.get("event_name", ""),
            "cluster_id": row["cluster_id"],
            "poc_name": row["poc_name"],
            "poc_email": row["poc_email"],
            "position": row.get("position", ""),
            "docx_path": docx_out,
            "pdf_path": os.path.join(pdf_dir, fname_base + ".pdf"),
        })

    # convert each generated docx to pdf via LibreOffice headless
    soffice = find_soffice()
    print(f"Converting {len(manifest_rows)} certificates to PDF via LibreOffice...")
    print(f"  Using LibreOffice at: {soffice}")
    for fname in os.listdir(docx_dir):
        if fname.endswith(".docx"):
            src = os.path.join(docx_dir, fname)
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, src],
                check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )

    return pd.DataFrame(manifest_rows)


from pathlib import Path

def draw_text_field(draw, img, field_cfg, text):
    raw_font_path = field_cfg.get("font_path", "sample_data/DejaVuSans.ttf")
    font_file = Path(raw_font_path)
    if not font_file.is_absolute():
        font_file = Path(__file__).parent / raw_font_path
    if not font_file.exists():
        font_file = Path(__file__).parent / "sample_data" / "DejaVuSans.ttf"

    font_size = int(field_cfg.get("font_size", 24))
    font = ImageFont.truetype(str(font_file), font_size)
    text_str = str(text)

    x = field_cfg["x"]
    y = field_cfg["y"]
    width = field_cfg.get("width", 0) or 0
    align = field_cfg.get("align", "center")
    color = field_cfg.get("color", "#000000")

    if width > 0:
        # Bounding Box (Line) mode:
        # Text is centered between X and X+Width, sitting on the baseline at Y.
        center_x = x + (width / 2)
        draw.text((center_x, y), text_str, font=font, fill=color, anchor='ms')
    else:
        # Point mode: align relative to exact (x, y) using Pillow native anchors
        # Matches CSS transform: translate(-50%/-100%/0%, -50%) exactly across Linux & Windows
        anchor_map = {
            "center": "mm",
            "right": "rm",
            "left": "lm"
        }
        anchor = anchor_map.get(align, "mm")
        draw.text((x, y), text_str, font=font, fill=color, anchor=anchor)



def load_base_image(template_path, dpi=300):
    """Load the design as a PIL Image. If it's a PDF, rasterize page 1 at high DPI first."""
    if template_path.lower().endswith(".pdf"):
        import fitz
        import io
        doc = fitz.open(template_path)
        if len(doc) == 0:
            raise ValueError(f"Could not read any pages from {template_path}")
        if len(doc) > 1:
            print(f"WARNING: {template_path} has {len(doc)} pages — using page 1 only.")
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=dpi)
        img_data = pix.tobytes("png")
        return Image.open(io.BytesIO(img_data)).convert("RGB")
    return Image.open(template_path).convert("RGB")


def generate_image_mode(df, cfg):
    active_cfg = get_active_cfg(cfg)
    at = cfg.get("active_type", "participation")
    out_dir = cfg.get("output_dir", "output")
    pdf_dir = os.path.join(out_dir, "pdf")
    os.makedirs(pdf_dir, exist_ok=True)

    template_path = active_cfg.get("image_template_path", "")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"image_template_path not found: {template_path}")

    render_dpi = active_cfg.get("image_render_dpi", 300)
    
    # Pre-cache helper to avoid reloading templates from disk repeatedly
    img_cache = {}
    def get_base_img(path):
        if not path or not os.path.exists(path):
            path = template_path
        if path not in img_cache:
            img_cache[path] = load_base_image(path, dpi=render_dpi)
        return img_cache[path]

    fields_cfg = active_cfg.get("image_text_fields", {})
    manifest_rows = []

    for _, row in df.iterrows():
        # Select background and text fields based on position for winner type
        curr_template = template_path
        curr_fields = fields_cfg

        if at == "winner" and "position" in row and not pd.isna(row["position"]):
            pos_val = str(row["position"]).strip().lower()
            pos_key = None
            if "1" in pos_val or "first" in pos_val:
                pos_key = "1st"
            elif "2" in pos_val or "second" in pos_val:
                pos_key = "2nd"
            elif "3" in pos_val or "third" in pos_val:
                pos_key = "3rd"

            if pos_key:
                curr_template = active_cfg.get(f"image_template_path_{pos_key}") or template_path
                curr_fields = active_cfg.get(f"image_text_fields_{pos_key}") or fields_cfg

        base_image = get_base_img(curr_template)
        img = base_image.copy()
        draw = ImageDraw.Draw(img)

        for field_name, field_cfg in curr_fields.items():
            if field_name in row and not pd.isna(row[field_name]):
                draw_text_field(draw, img, field_cfg, row[field_name])

        fname_base = safe_filename(row.get('student_name', row.get('volunteer_name', row.get('school', 'cert'))))
        pdf_out = os.path.join(pdf_dir, fname_base + ".pdf")
        img.save(pdf_out, "PDF", resolution=float(render_dpi))

        manifest_rows.append({
            "student_name": row.get("student_name", ""),
            "school": row.get("school", ""),
            "event_name": row.get("event_name", ""),
            "cluster_id": row["cluster_id"],
            "poc_name": row["poc_name"],
            "poc_email": row["poc_email"],
            "position": row.get("position", ""),
            "pdf_path": pdf_out,
        })

    return pd.DataFrame(manifest_rows)


# ---------------- MAIN ----------------

def main():
    cfg = load_config()
    df = load_roster(cfg)
    print(f"Loaded {len(df)} student rows across {df['poc_email'].nunique()} POC(s).")

    active_cfg = get_active_cfg(cfg)
    mode = active_cfg.get("template_mode", "docx")
    if mode == "docx":
        manifest = generate_docx_mode(df, cfg)
    elif mode == "image":
        manifest = generate_image_mode(df, cfg)
    else:
        print(f"Unknown template_mode '{mode}' — must be 'docx' or 'image'.", file=sys.stderr)
        sys.exit(1)

    manifest_path = os.path.join(cfg.get("output_dir", "output"), "manifest.csv")
    manifest.to_csv(manifest_path, index=False)

    missing_pdfs = manifest[~manifest["pdf_path"].apply(os.path.exists)]
    if len(missing_pdfs):
        print(f"WARNING: {len(missing_pdfs)} PDF(s) failed to generate — check errors above:")
        print(missing_pdfs[["student_name", "cluster_id"]].to_string(index=False))
    else:
        print(f"All {len(manifest)} certificates generated successfully.")

    print(f"Manifest written to: {manifest_path}")
    print("Next step: python send_emails.py")


if __name__ == "__main__":
    main()
