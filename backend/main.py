"""
backend/main.py  —  Certificate Automation FastAPI backend

Run from the PROJECT ROOT (cert_automation/):
    uvicorn backend.main:app --reload

All file paths in config.yaml are relative to the project root, so the
working directory MUST be the project root when uvicorn starts.
"""

import asyncio
import base64
import io
import json
import logging
import os
import smtplib
import ssl
import sys
import threading
from pathlib import Path
from typing import Dict, Optional

logger = logging.getLogger("cert_automation")


import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

# ── Paths ──────────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "backend"))

# Force working directory to project root so relative paths in config.yaml
# (excel_path, output_dir, etc.) resolve correctly.
os.chdir(ROOT)

# ── Import project helpers ─────────────────────────────────────────────────
from config_utils import patch_config, read_config, write_config  # noqa: E402
from session_store import (  # noqa: E402
    clear_credentials,
    get_app_password,
    get_sender_email,
    has_app_password,
    load_from_disk,
    set_app_password,
    set_credentials,
)

# Try to restore saved credentials on startup
load_from_disk()

# ── Import core functions — wrap, don't rewrite ────────────────────────────
from generate_certificates import (  # noqa: E402
    generate_docx_mode,
    generate_image_mode,
    load_base_image,
    load_roster,
    read_excel_any,
)
from send_emails import build_email, generate_excel_bytes, generate_event_zips  # noqa: E402

# ── App setup ──────────────────────────────────────────────────────────────
app = FastAPI(title="Certificate Automation API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── SSE helpers ────────────────────────────────────────────────────────────

class _QueueWriter(io.TextIOBase):
    """Redirect print() calls from background threads into an asyncio.Queue."""

    def __init__(self, loop: asyncio.AbstractEventLoop, q: asyncio.Queue):
        self._loop = loop
        self._q = q

    def write(self, text: str) -> int:
        if text.strip():
            for line in text.splitlines():
                if line.strip():
                    self._loop.call_soon_threadsafe(self._q.put_nowait, line.strip())
        return len(text)

    def flush(self) -> None:
        pass


def _sse(msg: str) -> str:
    return f"data: {msg}\n\n"


async def _run_in_thread_and_stream(fn, *args, **kwargs):
    """
    Execute fn(*args, **kwargs) in a daemon thread, capturing all print()
    output and yielding it as SSE events. Yields "__DONE__" when complete.
    """
    loop = asyncio.get_event_loop()
    q: asyncio.Queue = asyncio.Queue()

    def _target():
        old = sys.stdout
        sys.stdout = _QueueWriter(loop, q)
        try:
            fn(*args, **kwargs)
            loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
        except Exception as exc:
            loop.call_soon_threadsafe(q.put_nowait, f"ERROR: {exc}")
            loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
        finally:
            sys.stdout = old

    threading.Thread(target=_target, daemon=True).start()

    while True:
        try:
            msg = await asyncio.wait_for(q.get(), timeout=300.0)
        except asyncio.TimeoutError:
            yield _sse("[timeout — no output for 5 min, aborting]")
            break
        yield _sse(msg)
        if msg == "__DONE__":
            break


# ── Routes ─────────────────────────────────────────────────────────────────

def get_active_cfg(cfg=None):
    if cfg is None:
        cfg = read_config()
    at = cfg.get("active_type", "participation")
    return cfg.get("types", {}).get(at, {})

@app.get("/config/active-type")
def get_active_type():
    cfg = read_config()
    return {"active_type": cfg.get("active_type", "participation")}

class ActiveType(BaseModel):
    active_type: str

@app.put("/config/active-type")
def set_active_type(body: ActiveType):
    cfg = read_config()
    if body.active_type not in cfg.get("types", {}):
        raise HTTPException(400, "Invalid active_type")
    cfg["active_type"] = body.active_type
    write_config(cfg)
    return {"active_type": body.active_type}


@app.get("/")
def root():
    return {"status": "ok", "message": "Certificate Automation API is running"}


@app.get("/health")
def health():
    return {"status": "ok"}


# ---------- Upload Excel ----------

@app.post("/upload/excel")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx / .xls files are accepted")

    import time
    filename_only, ext = os.path.splitext(file.filename)
    unique_filename = f"{filename_only}_{int(time.time())}{ext}"
    dest = ROOT / "sample_data" / unique_filename
    dest.parent.mkdir(exist_ok=True)

    file_bytes = await file.read()

    try:
        with open(dest, "wb") as f:
            f.write(file_bytes)
    except Exception as e:
        raise HTTPException(500, f"Could not save file: {e}")

    rel_path = f"sample_data/{unique_filename}"
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = cfg.get("types", {}).get(at, {})

    # Read raw Excel to return column names + grouped preview FIRST
    # (Before writing config, because write_config triggers uvicorn reload)
    try:
        df = read_excel_any(dest, sheet_name=active_cfg.get("excel_sheet_name", 0))
        excel_columns = list(df.columns)
    except Exception as exc:
        raise HTTPException(500, f"Could not parse Excel file: {exc}")

    # Reset column mapping when a new file is uploaded — old mappings may not apply
    active_cfg["excel_path"] = rel_path
    # Setup default mappings based on type
    if at == "school":
        active_cfg["columns"] = {"school": "", "poc_email": "", "poc_name": ""}
    elif at == "volunteer":
        active_cfg["columns"] = {"volunteer_name": "", "volunteer_email": ""}
    else:
        active_cfg["columns"] = {"student_name": "", "event_name": "", "school": "", "poc_email": "", "poc_name": ""}
    cfg["types"][at] = active_cfg
    write_config(cfg)

    return {
        "excel_columns": excel_columns,
        "groups": [],           # no preview until user maps columns
        "excel_path": rel_path,
    }


@app.delete("/upload/excel")
def clear_excel():
    """Clear the uploaded Excel file from config so user can start fresh."""
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = cfg.get("types", {}).get(at, {})
    active_cfg["excel_path"] = ""
    if at == "school":
        active_cfg["columns"] = {"school": "", "poc_email": "", "poc_name": ""}
    elif at == "volunteer":
        active_cfg["columns"] = {"volunteer_name": "", "volunteer_email": ""}
    else:
        active_cfg["columns"] = {"student_name": "", "event_name": "", "school": "", "poc_email": "", "poc_name": ""}
    cfg["types"][at] = active_cfg
    write_config(cfg)
    return {"status": "cleared"}


# ---------- Upload Template ----------


@app.post("/upload/template")
async def upload_template(file: UploadFile = File(...), winner_position: Optional[str] = None):
    fname_lower = (file.filename or "").lower()
    if fname_lower.endswith(".docx"):
        mode, key = "docx", "docx_template_path"
    elif fname_lower.endswith((".pdf", ".png", ".jpg", ".jpeg")):
        mode, key = "image", "image_template_path"
    else:
        raise HTTPException(400, "Accepted formats: .docx, .pdf, .png, .jpg")

    dest = ROOT / "sample_data" / file.filename
    dest.parent.mkdir(exist_ok=True)
    with open(dest, "wb") as f:
        f.write(await file.read())

    rel_path = f"sample_data/{file.filename}"
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = cfg.get("types", {}).get(at, {})
    active_cfg["template_mode"] = mode
    
    if at == "winner" and winner_position in ("1st", "2nd", "3rd"):
        active_cfg[f"image_template_path_{winner_position}"] = rel_path
        # Use 1st place template as base image_template_path if that was uploaded,
        # or if no default path was set yet.
        if winner_position == "1st" or not active_cfg.get("image_template_path"):
            active_cfg["image_template_path"] = rel_path
    else:
        active_cfg[key] = rel_path
        
    cfg["types"][at] = active_cfg
    write_config(cfg)

    return {"template_mode": mode, "path": rel_path}


# ---------- Column Mapping ----------

@app.get("/config/columns")
def get_columns():
    cfg = read_config()
    active_cfg = get_active_cfg(cfg)
    columns = active_cfg.get("columns", {})
    excel_path_str = active_cfg.get("excel_path", "")
    excel_path = ROOT / excel_path_str if excel_path_str else None
    
    excel_columns = []
    groups = []
    
    if excel_path and excel_path.is_file():
        try:
            df = read_excel_any(excel_path, sheet_name=active_cfg.get("excel_sheet_name", 0))
            excel_columns = list(df.columns)
            
            # Re-create grouped preview if columns are mapped
            active_map = {k: v for k, v in columns.items() if v and str(v).strip()}
            if active_map and all(v in df.columns for v in active_map.values()):
                rename_map = {v: k for k, v in active_map.items()}
                dfrn = df.rename(columns=rename_map)
                if "poc_email" in dfrn.columns:
                    for poc_email, grp in dfrn.groupby("poc_email"):
                        student_details = []
                        for _, r in grp.iterrows():
                            student_details.append({
                                "name": str(r.get("student_name", "")),
                                "school": str(r.get("school", "")),
                                "event_name": str(r.get("event_name", "")),
                            })
                        poc_name = grp["poc_name"].iloc[0] if "poc_name" in grp.columns else str(poc_email).split("@")[0]
                        groups.append({
                            "poc_email": str(poc_email),
                            "poc_name": str(poc_name),
                            "student_count": len(grp),
                            "students": grp["student_name"].tolist() if "student_name" in grp else [],
                            "student_details": student_details,
                        })
        except Exception:
            pass
            
    return {
        "columns": columns,
        "excel_path": excel_path_str,
        "excel_columns": excel_columns,
        "groups": groups,
        "excel_sheet_name": active_cfg.get("excel_sheet_name", 0),
    }


class ColumnMapping(BaseModel):
    columns: Dict[str, str]


@app.post("/config/columns")
def set_columns(body: ColumnMapping):
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = cfg["types"][at]
    active_cfg["columns"] = body.columns
    cfg["types"][at] = active_cfg
    write_config(cfg)

    # Re-read the Excel to refresh the grouped preview with new mapping
    excel_path_str = cfg.get("excel_path", "")
    excel_path = ROOT / excel_path_str if excel_path_str else None
    groups = []
    excel_columns = []
    if excel_path and excel_path.is_file():
        try:
            df = read_excel_any(excel_path, sheet_name=active_cfg.get("excel_sheet_name", 0))
            excel_columns = list(df.columns)
            col_map = body.columns
            active_map = {k: v for k, v in col_map.items() if v and str(v).strip()}
            if active_map and all(v in df.columns for v in active_map.values()):
                rename_map = {v: k for k, v in active_map.items()}
                dfrn = df.rename(columns=rename_map)
                if "poc_email" in dfrn.columns:
                    for poc_email, grp in dfrn.groupby("poc_email"):
                        student_details = []
                        for _, r in grp.iterrows():
                            student_details.append({
                                "name": str(r.get("student_name", "")),
                                "school": str(r.get("school", "")),
                                "event_name": str(r.get("event_name", "")),
                            })
                        poc_name = grp["poc_name"].iloc[0] if "poc_name" in grp.columns else str(poc_email).split("@")[0]
                        groups.append({
                            "poc_email": str(poc_email),
                            "poc_name": str(poc_name),
                            "student_count": len(grp),
                            "students": grp["student_name"].tolist() if "student_name" in grp else [],
                            "student_details": student_details,
                        })
        except Exception as exc:
            import traceback; traceback.print_exc()
            pass

    return {"status": "ok", "columns": body.columns, "groups": groups, "excel_columns": excel_columns}


# ---------- Debug endpoints (non-destructive) ----------

@app.get("/debug/raw-config")
def debug_raw_config():
    """Return the RAW image_text_fields stored in config.yaml, no transforms or defaults applied."""
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = get_active_cfg(cfg)
    raw_fields = active_cfg.get("image_text_fields", {})
    template_path = active_cfg.get("image_template_path", "")
    template_exists = os.path.exists(template_path) if template_path else False
    dpi = active_cfg.get("image_render_dpi", 300)
    img_size = None
    if template_exists:
        try:
            img = load_base_image(template_path, dpi=dpi)
            img_size = {"width": img.size[0], "height": img.size[1]}
        except Exception as e:
            img_size = {"error": str(e)}
    return {
        "active_type": at,
        "template_path": template_path,
        "template_exists": template_exists,
        "image_size_px": img_size,
        "render_dpi": dpi,
        "image_text_fields": raw_fields,
    }


@app.get("/debug/cert-render")
def debug_cert_render():
    try:
        cfg = read_config()
        at = cfg.get("active_type", "participation")
        active_cfg = get_active_cfg(cfg)

        template_path = active_cfg.get("image_template_path", "")
        if not template_path or not os.path.exists(template_path):
            raise HTTPException(404, "Template not found — upload a template first.")

        raw_fields = active_cfg.get("image_text_fields", {})
        if not raw_fields:
            raise HTTPException(400, "No image_text_fields saved in config — set positions first.")

        dpi = active_cfg.get("image_render_dpi", 300)
        img = load_base_image(template_path, dpi=dpi)
        draw = __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(img)
        w, h = img.size

        DEBUG_SAMPLE = {
            "student_name": "Test Student",
            "event_name": "Test Event 2026",
            "school": "Test School",
            "volunteer_name": "Test Volunteer",
        }

        for field_name, field_cfg in raw_fields.items():
            sample_text = DEBUG_SAMPLE.get(field_name, field_name.replace("_", " ").title())
            draw_text_field(draw, img, field_cfg, sample_text)

            fx, fy = int(field_cfg.get("x", 0)), int(field_cfg.get("y", 0))
            r = max(12, int(w * 0.005))
            draw.ellipse([fx-r, fy-r, fx+r, fy+r], fill=(255, 0, 0), outline=(255, 255, 0))
            draw.line([(fx-r*2, fy), (fx+r*2, fy)], fill=(255, 255, 0), width=max(2, r//3))
            draw.line([(fx, fy-r*2), (fx, fy+r*2)], fill=(255, 255, 0), width=max(2, r//3))

        max_w = 1400
        if w > max_w:
            ratio = max_w / w
            img = img.resize((max_w, int(h * ratio)), resample=__import__("PIL.Image", fromlist=["Image"]).LANCZOS)

        buf = __import__("io").BytesIO()
        img.save(buf, format="PNG")
        b64 = __import__("base64").b64encode(buf.getvalue()).decode()
        return {"image": f"data:image/png;base64,{b64}", "width": w, "height": h, "field_names": list(raw_fields.keys())}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(500, detail=f"Debug render error: {exc}")


# ---------- Coordinates Preview (image mode) ----------

@app.post("/coordinates/preview")
def coordinates_preview(winner_position: Optional[str] = None):
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = get_active_cfg(cfg)
    
    tpl = ""
    if at == "winner" and winner_position in ("1st", "2nd", "3rd"):
        tpl = active_cfg.get(f"image_template_path_{winner_position}", "")
    if not tpl:
        tpl = active_cfg.get("image_template_path", "")

    template_path = ROOT / tpl if tpl else None

    if not template_path or not template_path.exists():
        raise HTTPException(404, "image_template_path not found — upload an image/PDF template first")

    try:
        from PIL import ImageDraw
        dpi = active_cfg.get("image_render_dpi", 300)
        img = load_base_image(str(template_path), dpi=dpi)
        draw = ImageDraw.Draw(img)
        w, h = img.size
        step = 100

        for x in range(0, w, step):
            draw.line([(x, 0), (x, h)], fill=(220, 38, 38), width=1)
            draw.text((x + 2, 2), str(x), fill=(220, 38, 38))
        for y in range(0, h, step):
            draw.line([(0, y), (w, y)], fill=(37, 99, 235), width=1)
            draw.text((2, y + 2), str(y), fill=(37, 99, 235))

        # Downscale if very large before base64-encoding
        max_w = 1400
        if w > max_w:
            ratio = max_w / w
            img = img.resize((max_w, int(h * ratio)))

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode()
        return {"image": f"data:image/png;base64,{b64}", "width": w, "height": h}
    except Exception as exc:
        raise HTTPException(500, f"Could not generate preview: {exc}")


# ---------- Template config (lightweight read for frontend hydration) ----------

@app.get("/config/template")
def get_template_config():
    cfg = read_config()
    active_cfg = get_active_cfg(cfg)
    return {
        "template_mode": active_cfg.get("template_mode", "docx"),
        "image_template_path": active_cfg.get("image_template_path", ""),
        "docx_template_path": active_cfg.get("docx_template_path", ""),
        "image_template_path_1st": active_cfg.get("image_template_path_1st", ""),
        "image_template_path_2nd": active_cfg.get("image_template_path_2nd", ""),
        "image_template_path_3rd": active_cfg.get("image_template_path_3rd", ""),
    }


# ---------- Image Text Fields ----------

@app.get("/config/text-fields")
def get_text_fields(winner_position: Optional[str] = None):
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = get_active_cfg(cfg)

    fields = {}
    if at == "winner" and winner_position in ("1st", "2nd", "3rd"):
        fields = active_cfg.get(f"image_text_fields_{winner_position}", {})
    if not fields:
        fields = active_cfg.get("image_text_fields", {})

    # Color extracted from template images (dominant non-background tone)
    CERT_COLOR = "#13314B"  # dark navy matching the certificate design

    # Per-type font defaults
    TYPE_FONT = {
        "participation": {
            "font_path": "sample_data/DancingScript-Regular.ttf",
            "font_family": "'Dancing Script', cursive",
            "font_size": 90,
            "is_bold": False,
        },
        "winner": {
            "font_path": "sample_data/Montserrat-SemiBold.ttf",
            "font_family": "'Montserrat', sans-serif",
            "font_size": 75,
            "is_bold": True,
        },
        "school": {
            "font_path": "sample_data/PlayfairDisplay-SemiBold.ttf",
            "font_family": "'Playfair Display', serif",
            "font_size": 75,
            "is_bold": True,
        },
        "volunteer": {
            "font_path": "sample_data/PlayfairDisplay-SemiBold.ttf",
            "font_family": "'Playfair Display', serif",
            "font_size": 75,
            "is_bold": True,
        },
    }

    font_defaults = TYPE_FONT.get(at, TYPE_FONT["winner"])
    allowed = ["school"] if at == "school" else ["volunteer_name"] if at == "volunteer" else ["student_name", "event_name", "school"]
    filtered_fields = {}

    for i, f in enumerate(allowed):
        if f in fields:
            filtered_fields[f] = fields[f]
            if "width" not in filtered_fields[f]:
                filtered_fields[f]["width"] = 1000
        else:
            filtered_fields[f] = {
                "x": 600, "y": 400 + (i * 100),
                "width": 1000,
                "color": CERT_COLOR,
                "align": "center",
                **font_defaults,
            }

    return {"image_text_fields": filtered_fields}



@app.post("/config/text-fields")
def set_text_fields(fields: dict, winner_position: Optional[str] = None):
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = cfg.get("types", {}).get(at, {})
    
    if at == "winner" and winner_position in ("1st", "2nd", "3rd"):
        active_cfg[f"image_text_fields_{winner_position}"] = fields

    # Always sync default image_text_fields if it's 1st or default isn't set yet
    if at != "winner" or winner_position == "1st" or not active_cfg.get("image_text_fields"):
        active_cfg["image_text_fields"] = fields

    cfg["types"][at] = active_cfg
    write_config(cfg)
    return {"status": "ok"}


@app.post("/config/reset-text-fields")
def reset_text_fields(winner_position: Optional[str] = None):
    cfg = read_config()
    at = cfg.get("active_type", "participation")
    active_cfg = cfg.get("types", {}).get(at, {})

    if at == "winner" and winner_position in ("1st", "2nd", "3rd"):
        active_cfg[f"image_text_fields_{winner_position}"] = {}
    if at != "winner" or winner_position == "1st":
        active_cfg["image_text_fields"] = {}

    cfg["types"][at] = active_cfg
    write_config(cfg)
    return {"status": "ok"}


# ---------- Generate (SSE) ----------

@app.post("/generate")
async def generate_endpoint():
    loop = asyncio.get_event_loop()
    q: asyncio.Queue = asyncio.Queue()

    def _run():
        old = sys.stdout
        sys.stdout = _QueueWriter(loop, q)
        try:
            cfg = read_config()
            df = load_roster(cfg)
            loop.call_soon_threadsafe(
                q.put_nowait,
                f"Loaded {len(df)} student rows across {df['poc_email'].nunique()} POC group(s)."
            )

            active_cfg = get_active_cfg(cfg)
            mode = active_cfg.get("template_mode", "docx")
            if mode == "docx":
                manifest = generate_docx_mode(df, cfg)
            elif mode == "image":
                manifest = generate_image_mode(df, cfg)
            else:
                raise ValueError(f"Unknown template_mode '{mode}' — must be 'docx' or 'image'")

            out_dir = cfg.get("output_dir", "output")
            manifest_path = os.path.join(out_dir, "manifest.csv")
            manifest.to_csv(manifest_path, index=False)

            missing = manifest[~manifest["pdf_path"].apply(os.path.exists)]
            if len(missing):
                loop.call_soon_threadsafe(
                    q.put_nowait,
                    f"WARNING: {len(missing)} PDF(s) did not generate — check template/data."
                )
            else:
                loop.call_soon_threadsafe(
                    q.put_nowait,
                    f"All {len(manifest)} certificates generated successfully."
                )

            loop.call_soon_threadsafe(q.put_nowait, f"Manifest: {manifest_path}")
            loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
        except Exception as exc:
            loop.call_soon_threadsafe(q.put_nowait, f"ERROR: {exc}")
            loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
        finally:
            sys.stdout = old

    threading.Thread(target=_run, daemon=True).start()

    async def _stream():
        while True:
            try:
                msg = await asyncio.wait_for(q.get(), timeout=300.0)
            except asyncio.TimeoutError:
                yield _sse("[timeout — generation took too long]")
                break
            yield _sse(msg)
            if msg == "__DONE__":
                break

    return StreamingResponse(
        _stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ---------- Manifest summary ----------

@app.get("/manifest")
def get_manifest():
    cfg = read_config()
    manifest_path = ROOT / cfg.get("output_dir", "output") / "manifest.csv"
    if not manifest_path.exists():
        return {"groups": [], "total": 0, "first_pdf": None}

    df = pd.read_csv(str(manifest_path))
    # Keep only rows whose PDFs exist on disk (absolute OR relative to ROOT)
    def _pdf_exists(p: str) -> bool:
        return os.path.exists(p) or os.path.exists(ROOT / p)

    df = df[df["pdf_path"].apply(_pdf_exists)]

    groups = []
    for poc_email, grp in df.groupby("poc_email"):
        student_details = []
        for _, r in grp.iterrows():
            student_details.append({
                "name": str(r.get("student_name", "")),
                "school": str(r.get("school", "")),
                "event_name": str(r.get("event_name", "")),
                "position": str(r.get("position", "")) if "position" in r else "",
            })
        poc_name = grp["poc_name"].iloc[0] if "poc_name" in grp.columns else str(poc_email).split("@")[0]
        groups.append({
            "poc_email": str(poc_email),
            "poc_name": str(poc_name),
            "student_count": len(grp),
            "students": grp["student_name"].tolist(),
            "student_details": student_details,
            "pdf_files": [os.path.basename(p) for p in grp["pdf_path"].tolist()],
        })

    first_pdf = os.path.basename(df["pdf_path"].iloc[0]) if len(df) else None
    return {"groups": groups, "total": len(df), "first_pdf": first_pdf}


# ---------- Serve PDF ----------

@app.get("/pdf/{filename:path}")
def serve_pdf(filename: str, download: bool = False):
    safe_name = os.path.basename(filename)  # prevent directory traversal
    pdf_path = ROOT / "output" / "pdf" / safe_name
    if not pdf_path.exists():
        raise HTTPException(404, f"PDF not found: {safe_name}")

    # Always set Content-Disposition explicitly so the browser uses the correct filename
    disposition = "attachment" if download else "inline"
    from urllib.parse import quote
    encoded_name = quote(safe_name, safe="")
    headers = {
        "Content-Disposition": f'{disposition}; filename="{safe_name}"; filename*=UTF-8\'\'{encoded_name}',
        "Cache-Control": "no-cache",
    }
    return FileResponse(str(pdf_path), media_type="application/pdf", headers=headers)


# ---------- PDF Preview (renders first page as PNG) ----------

@app.get("/pdf-preview/{filename:path}")
def serve_pdf_preview(filename: str):
    safe_name = os.path.basename(filename)  # prevent directory traversal
    pdf_path = ROOT / "output" / "pdf" / safe_name
    if not pdf_path.exists():
        raise HTTPException(404, f"PDF not found: {safe_name}")
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        if len(doc) == 0:
            raise HTTPException(500, "PDF has no pages")
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        return StreamingResponse(io.BytesIO(img_data), media_type="image/png")
    except Exception as exc:
        raise HTTPException(500, f"Could not generate PDF preview: {exc}")


# ---------- Email auth ----------

class EmailAuth(BaseModel):
    sender_email: str
    app_password: str  # NEVER written to disk; stored in session_store only


@app.post("/config/email-auth")
def set_email_auth(body: EmailAuth):
    cfg = read_config()
    cfg["email"]["sender_email"] = body.sender_email
    write_config(cfg)                                 # sender_email written to config.yaml
    set_credentials(body.sender_email, body.app_password, persist=True)  # persisted locally
    return {"status": "ok", "sender_email": body.sender_email, "password_set": True}


@app.post("/config/email-auth/signout")
def sign_out_email_auth():
    clear_credentials()  # wipes memory + .auth_credentials file
    return {"status": "signed_out"}


@app.get("/config/email-auth")
def get_email_auth_status():
    cfg = read_config()
    sender = get_sender_email() or cfg.get("email", {}).get("sender_email", "")
    return {
        "sender_email": sender,
        "password_set": has_app_password(),
    }


@app.get("/config/drive-auth")
def get_drive_auth_status():
    from drive_service import (
        is_drive_available, find_credentials_file, get_service_account_email,
        is_oauth_drive_available, is_oauth_client_available, is_oauth_token_available,
        get_oauth_user_email,
    )
    cred_file = find_credentials_file()
    cfg = read_config()
    return {
        "drive_available": is_drive_available(),
        # Service Account
        "has_credentials": cred_file is not None,
        "credentials_path": str(cred_file.name) if cred_file else None,
        "service_account_email": get_service_account_email(),
        # OAuth2 user credentials
        "oauth_client_available": is_oauth_client_available(),
        "oauth_token_available": is_oauth_token_available(),
        "oauth_drive_available": is_oauth_drive_available(),
        "oauth_user_email": get_oauth_user_email(),
        # Config
        "root_folder_id": cfg.get("email", {}).get("drive_root_folder_id", ""),
    }


class DriveFolderRequest(BaseModel):
    folder_url_or_id: str


@app.post("/config/drive-auth/root-folder")
def set_drive_root_folder(body: DriveFolderRequest):
    from drive_service import extract_folder_id
    clean_id = extract_folder_id(body.folder_url_or_id)
    cfg = read_config()
    if "email" not in cfg:
        cfg["email"] = {}
    cfg["email"]["drive_root_folder_id"] = clean_id or ""
    write_config(cfg)
    return {"status": "saved", "drive_root_folder_id": clean_id}


@app.post("/config/drive-auth/upload")
async def upload_drive_credentials(file: UploadFile = File(...)):
    if not file.filename.endswith(".json"):
        raise HTTPException(status_code=400, detail="Credentials file must be a .json file.")

    contents = await file.read()
    try:
        data = json.loads(contents)
        if "type" not in data or "project_id" not in data:
            raise ValueError("File does not appear to be a valid Google Service Account JSON key.")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON file: {exc}")

    target_path = ROOT / "service_account.json"
    target_path.write_bytes(contents)

    from drive_service import is_drive_available
    return {
        "status": "uploaded",
        "drive_available": is_drive_available(),
        "filename": file.filename
    }


@app.post("/config/drive-auth/delete")
def delete_drive_credentials():
    from drive_service import find_credentials_file
    cred_file = find_credentials_file()
    if cred_file and cred_file.exists():
        cred_file.unlink()
    return {"status": "deleted"}


# ─── OAuth2 Drive endpoints ───────────────────────────────────────────────────

@app.post("/config/drive-auth/oauth-upload")
async def upload_oauth_client_secrets(file: UploadFile = File(...)):
    """Upload the OAuth2 client_secrets.json (Web Application type) from Google Cloud Console."""
    if not file.filename.endswith(".json"):
        raise HTTPException(status_code=400, detail="Must be a .json file.")
    contents = await file.read()
    try:
        data = json.loads(contents)
        if "installed" not in data and "web" not in data:
            raise ValueError("Does not appear to be an OAuth2 client_secrets.json file.")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid file: {exc}")
    from drive_service import OAUTH_CLIENT_SECRETS_FILE
    OAUTH_CLIENT_SECRETS_FILE.write_bytes(contents)
    return {"status": "uploaded"}


@app.get("/config/drive-auth/oauth-url")
def get_oauth_url(redirect_uri: str):
    """
    Generate the Google OAuth2 authorization URL for the web redirect flow.
    The frontend opens this URL in a new tab; Google redirects back to redirect_uri.
    """
    from drive_service import is_oauth_client_available, get_oauth_authorization_url
    if not is_oauth_client_available():
        raise HTTPException(status_code=400, detail="oauth_client_secrets.json not uploaded yet.")
    try:
        auth_url, state = get_oauth_authorization_url(redirect_uri)
        return {"url": auth_url, "state": state}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/config/drive-auth/oauth-callback")
def oauth_callback(code: str, state: str):
    """
    Google redirects here after the user grants (or denies) Drive access.
    Exchanges the code for tokens, saves them, and returns a self-closing HTML page.
    """
    from drive_service import complete_oauth_from_callback
    from fastapi.responses import HTMLResponse

    success_html = """<!DOCTYPE html>
<html>
<head><title>Google Drive Connected</title></head>
<body style="font-family:sans-serif;text-align:center;padding:60px;background:#f0fdf4;">
  <div style="max-width:400px;margin:auto;background:#fff;border-radius:12px;padding:40px;box-shadow:0 4px 24px #0001;">
    <div style="font-size:48px;margin-bottom:16px;">✅</div>
    <h2 style="color:#15803d;margin:0 0 12px">Google Drive Connected!</h2>
    <p style="color:#166534;margin:0 0 20px">Authorization successful. This tab will close in 3 seconds.</p>
    <p style="color:#94a3b8;font-size:12px">Return to the Certificate Automation app.</p>
  </div>
  <script>setTimeout(() => window.close(), 3000);</script>
</body>
</html>"""

    error_html = """<!DOCTYPE html>
<html>
<head><title>Authorization Failed</title></head>
<body style="font-family:sans-serif;text-align:center;padding:60px;background:#fff1f2;">
  <div style="max-width:400px;margin:auto;background:#fff;border-radius:12px;padding:40px;box-shadow:0 4px 24px #0001;">
    <div style="font-size:48px;margin-bottom:16px;">❌</div>
    <h2 style="color:#be123c;margin:0 0 12px">Authorization Failed</h2>
    <p style="color:#9f1239;margin:0 0 8px">{error}</p>
    <p style="color:#94a3b8;font-size:12px">Close this tab and try again in the app.</p>
  </div>
</body>
</html>"""

    try:
        complete_oauth_from_callback(code=code, state=state)
        return HTMLResponse(success_html)
    except Exception as exc:
        return HTMLResponse(error_html.format(error=str(exc)), status_code=400)


@app.get("/config/drive-auth/oauth-status")
def oauth_status():
    """Poll the current state of the OAuth2 authorization flow."""
    from drive_service import get_oauth_flow_status, get_oauth_user_email, is_oauth_drive_available
    state = get_oauth_flow_status()
    return {
        **state,
        "oauth_drive_available": is_oauth_drive_available(),
        "oauth_user_email": get_oauth_user_email(),
    }


@app.post("/config/drive-auth/oauth-revoke")
def oauth_revoke():
    """Revoke and delete the saved OAuth2 token (sign out)."""
    from drive_service import revoke_oauth_token
    revoke_oauth_token()
    return {"status": "revoked"}


# ---------- Send (SSE) ----------

class SendRequest(BaseModel):
    poc_email: Optional[str] = None


@app.post("/send")
async def send_endpoint(body: SendRequest):
    loop = asyncio.get_event_loop()
    q: asyncio.Queue = asyncio.Queue()

    def _run():
        old = sys.stdout
        sys.stdout = _QueueWriter(loop, q)
        try:
            cfg = read_config()
            manifest_path = ROOT / cfg.get("output_dir", "output") / "manifest.csv"

            if not manifest_path.exists():
                loop.call_soon_threadsafe(
                    q.put_nowait, "ERROR: No manifest found — run Generate first."
                )
                loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
                return

            df = pd.read_csv(str(manifest_path))
            df = df[df["pdf_path"].apply(os.path.exists)]

            if body.poc_email:
                df = df[df["poc_email"] == body.poc_email]

            if len(df) == 0:
                loop.call_soon_threadsafe(
                    q.put_nowait, "ERROR: No valid PDFs found in manifest to send."
                )
                loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
                return

            password = get_app_password()
            if not password:
                loop.call_soon_threadsafe(
                    q.put_nowait, "ERROR: App password not set. Fill in the auth panel first."
                )
                loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
                return

            at = cfg.get("active_type", "participation")
            active_cfg = cfg.get("types", {}).get(at, {})
            global_ecfg = cfg.get("email", {})
            type_ecfg = active_cfg.get("email", {})
            ecfg = {**global_ecfg, **type_ecfg}
            # Always send to real POC emails — no dry_run mode
            ecfg["dry_run"] = False
            cfg_patched = {**cfg, "email": ecfg}

            total_pocs = df['poc_email'].nunique()
            loop.call_soon_threadsafe(q.put_nowait, f"Sending to {total_pocs} POC(s)...")

            ssl_context = ssl.create_default_context()
            smtp_timeout = 60  # seconds per socket operation
            sent, failed = 0, 0
            results = []

            def _create_smtp_connection(host: str, port: int):
                # Force IPv4 to prevent [Errno 101] Network is unreachable on cloud environments (like Render)
                target_ip = host
                try:
                    import socket
                    addrs = socket.getaddrinfo(host, port, socket.AF_INET, socket.SOCK_STREAM)
                    if addrs:
                        target_ip = addrs[0][4][0]
                except Exception as exc:
                    logger.warning(f"IPv4 resolution for {host} failed: {exc}")

                if port == 465:
                    return smtplib.SMTP_SSL(target_ip, port, timeout=smtp_timeout, context=ssl_context, server_hostname=host)
                else:
                    srv = smtplib.SMTP(target_ip, port, timeout=smtp_timeout)
                    srv.ehlo()
                    srv.starttls(context=ssl_context)
                    srv.ehlo()
                    return srv

            # Check if Gmail API is available (OAuth token has gmail.send scope)
            def _gmail_api_available() -> bool:
                try:
                    from drive_service import is_oauth_drive_available, OAUTH_TOKEN_FILE
                    import json
                    if not is_oauth_drive_available():
                        return False
                    data = json.loads(OAUTH_TOKEN_FILE.read_text(encoding="utf-8"))
                    scopes = data.get("scopes", [])
                    return any("gmail" in s for s in scopes)
                except Exception:
                    return False

            _use_gmail_api = _gmail_api_available()

            def _send_one(msg_obj):
                """Send one email. Tries Gmail REST API first (HTTPS, always works on cloud);
                falls back to direct SMTP if Gmail API not authorized."""
                if _use_gmail_api:
                    from drive_service import send_via_gmail_api
                    send_via_gmail_api(msg_obj)
                    return

                # SMTP fallback
                port = int(ecfg.get("smtp_port", 465))
                host = ecfg["smtp_host"]
                srv = None
                try:
                    srv = _create_smtp_connection(host, port)
                except Exception as first_err:
                    alt_port = 587 if port == 465 else 465
                    logger.warning(f"SMTP {host}:{port} failed ({first_err}). Retrying port {alt_port}...")
                    srv = _create_smtp_connection(host, alt_port)

                try:
                    srv.login(ecfg["sender_email"], password)
                    srv.send_message(msg_obj)
                finally:
                    try:
                        srv.quit()
                    except Exception:
                        pass

            import time

            for poc_email, group in df.groupby("poc_email"):
                # Safely read optional columns from manifest
                poc_name = (
                    group["poc_name"].iloc[0]
                    if "poc_name" in group.columns
                    else str(poc_email).split("@")[0]
                )
                event_name = (
                    group["event_name"].iloc[0]
                    if "event_name" in group.columns
                    else ""
                )
                student_names = group["student_name"].tolist()
                pdf_paths = group["pdf_path"].tolist()

                # Google Drive Automated Upload Integration (bulk mode only)
                drive_data = None
                if at not in ["school", "volunteer"]:
                    try:
                        from drive_service import is_drive_available, upload_poc_certificates_to_drive
                        if is_drive_available():
                            loop.call_soon_threadsafe(
                                q.put_nowait, f"  → Uploading certificates to Google Drive for {poc_email}..."
                            )
                            drive_data = upload_poc_certificates_to_drive(
                                poc_name, poc_email, group, pdf_paths,
                                parent_folder_id=ecfg.get("drive_root_folder_id"),
                                active_type=at
                            )
                            if drive_data and drive_data.get("poc_folder_url"):
                                loop.call_soon_threadsafe(
                                    q.put_nowait, f"  ✓ Drive folder created & certificates uploaded."
                                )
                    except Exception as drive_exc:
                        err_txt = str(drive_exc)
                        if "storageQuotaExceeded" in err_txt or "storage quota" in err_txt.lower():
                            from drive_service import get_service_account_email
                            sa_email = get_service_account_email() or "Service Account"
                            err_txt = f"Quota Exceeded: Share a folder on your personal Google Drive with '{sa_email}' as Editor and paste its URL in Data tab"
                        loop.call_soon_threadsafe(
                            q.put_nowait, f"  → Note: Drive upload skipped ({err_txt}). Using fallback attachments."
                        )

                try:
                    msg = build_email(
                        poc_name, poc_email, poc_email, event_name,
                        student_names, pdf_paths, cfg_patched, group, drive_data=drive_data
                    )
                    _send_one(msg)
                    actual_to = str(msg["To"])
                    loop.call_soon_threadsafe(
                        q.put_nowait,
                        f"Sent → {actual_to}  ({len(pdf_paths)} certificate(s))  [{sent + 1}/{total_pocs}]"
                    )
                    sent += 1
                    results.append({
                        "poc_name": poc_name,
                        "sent_to": actual_to,
                        "attachments": len(pdf_paths),
                        "status": "sent",
                    })
                except Exception as exc:
                    err_msg = str(exc)
                    if "Server not connected" in err_msg or isinstance(exc, smtplib.SMTPServerDisconnected):
                        err_msg = f"{exc} (SMTP size limit or socket timeout)"
                    loop.call_soon_threadsafe(
                        q.put_nowait, f"FAILED → {poc_email}: {err_msg}  [{sent + failed + 1}/{total_pocs}]"
                    )
                    failed += 1

                    # ── Fallback: save PDFs + message.html + christ_header.png + student_list.xlsx so user can send manually ──
                    try:
                        import shutil
                        # @ is valid in Windows/Linux folder names — use address as-is
                        fallback_dir = ROOT / "fallback" / str(poc_email)
                        fallback_dir.mkdir(parents=True, exist_ok=True)

                        # Save event-grouped ZIP files into fallback folder
                        try:
                            event_zips = generate_event_zips(group, pdf_paths, active_type=at)
                            for zip_name, zip_bytes in event_zips.items():
                                (fallback_dir / zip_name).write_bytes(zip_bytes)
                        except Exception:
                            pass

                        # Copy every PDF into the fallback folder
                        for pdf_path in pdf_paths:
                            if os.path.exists(pdf_path):
                                shutil.copy2(pdf_path, fallback_dir / os.path.basename(pdf_path))

                        # Copy the logo header banner for offline rendering
                        header_src = ROOT / "sample_data" / "christ_header.png"
                        if header_src.exists():
                            shutil.copy2(header_src, fallback_dir / "christ_header.png")

                        # Save student_list.xlsx in fallback folder
                        try:
                            excel_bytes = generate_excel_bytes(student_names, pdf_paths, group, active_type=at)
                            if excel_bytes:

                                (fallback_dir / "student_list.xlsx").write_bytes(excel_bytes)
                        except Exception:
                            pass

                        subject = str(msg["Subject"]) if msg["Subject"] else "(no subject)"
                        
                        # 1. Plain text fallback message
                        body_part = msg.get_body(preferencelist=("plain",))
                        body_text = body_part.get_content() if body_part else "(no body)"
                        message_txt = (
                            f"TO: {poc_email}\n"
                            f"SUBJECT: {subject}\n"
                            f"{'=' * 60}\n\n"
                            f"{body_text.strip()}\n\n"
                            f"{'=' * 60}\n"
                            f"ATTACHMENTS ({len(pdf_paths) + 1}):\n"
                            f"  - student_list.xlsx\n"
                            + "\n".join(f"  - {os.path.basename(p)}" for p in pdf_paths)
                        )
                        (fallback_dir / "message.txt").write_text(message_txt, encoding="utf-8")

                        # 2. HTML fallback message (perfect for copying/pasting directly into Gmail)
                        html_part = msg.get_body(preferencelist=("html",))
                        if html_part:
                            html_content = html_part.get_content()
                            # Convert header image to Base64 data URL so it renders properly when copy-pasted to email clients
                            header_src = ROOT / "sample_data" / "christ_header.png"
                            if header_src.exists():
                                import base64
                                try:
                                    b64_data = base64.b64encode(header_src.read_bytes()).decode("utf-8")
                                    html_for_copy = html_content.replace("cid:christ_header", f"data:image/png;base64,{b64_data}")
                                except Exception:
                                    html_for_copy = html_content.replace("cid:christ_header", "christ_header.png")
                            else:
                                html_for_copy = html_content.replace("cid:christ_header", "christ_header.png")
                            (fallback_dir / "message.html").write_text(html_for_copy, encoding="utf-8")

                        loop.call_soon_threadsafe(
                            q.put_nowait,
                            f"  → Fallback saved: fallback/{poc_email}/ ({len(pdf_paths)} PDF(s) + student_list.xlsx + message.html)"
                        )
                    except Exception as fb_exc:
                        loop.call_soon_threadsafe(
                            q.put_nowait, f"  → WARNING: Could not write fallback: {fb_exc}"
                        )

                    results.append({
                        "poc_name": poc_name,
                        "sent_to": str(poc_email),
                        "attachments": len(pdf_paths),
                        "status": "failed",
                    })

                # Brief pause between sends to stay well within Gmail rate limits
                if total_pocs > 1:
                    time.sleep(1)

            loop.call_soon_threadsafe(
                q.put_nowait, f"Done. {sent} email(s) sent, {failed} failed."
            )
            loop.call_soon_threadsafe(
                q.put_nowait, f"__RESULTS__:{json.dumps(results)}"
            )
            loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
        except Exception as exc:
            import traceback
            err_detail = traceback.format_exc()
            loop.call_soon_threadsafe(q.put_nowait, f"ERROR: {exc}")
            loop.call_soon_threadsafe(q.put_nowait, f"DETAIL: {err_detail.splitlines()[-1]}")
            loop.call_soon_threadsafe(q.put_nowait, "__DONE__")
        finally:
            sys.stdout = old

    threading.Thread(target=_run, daemon=True).start()

    async def _stream():
        while True:
            try:
                msg = await asyncio.wait_for(q.get(), timeout=600.0)
            except asyncio.TimeoutError:
                yield _sse("[timeout]")
                break
            yield _sse(msg)
            if msg == "__DONE__":
                break

    return StreamingResponse(
        _stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
