# Certificate Automation

Generates a personalized certificate PDF for every student in your Excel roster, then
sends **one bulk email per POC** with all of that POC's students' certificates attached.

Tested and working in **two modes**, so it doesn't matter yet whether your final design
turns out to be a Word file or a Canva/image export — just flip one setting once you know.

```
cert_automation/
├── config.yaml                 <- EDIT THIS. Everything else stays as-is.
├── generate_certificates.py    <- Step 1: builds the PDFs
├── send_emails.py              <- Step 2: sends the grouped bulk emails
├── find_coordinates.py         <- Only needed for image-mode templates
├── sample_data/                <- Demo files so you can test right now
└── output/                     <- Generated PDFs + manifest.csv land here
```

## Quick start (with the included sample data — try this first)

```bash
pip install -r requirements.txt
python generate_certificates.py     # makes 6 sample certificates in output/pdf/
python send_emails.py               # dry-run: sends everything to YOUR OWN email
```

Open `output/pdf/` to see the generated certificates, and check your inbox (dry-run
redirects all sends to you) before ever touching a real POC's address.


## Web Dashboard Setup

This project includes a Web Dashboard to upload datasets/templates, configure column mappings, visually position certificate text fields, preview outputs, and send bulk emails.

To run the Web Dashboard, you must start **both** the FastAPI backend and the React/Vite frontend:

### 1. Start the Backend API (Python)
From the **project root** directory (`cert_automation/`):
```bash
# Install required Python packages for CLI and backend
pip install -r requirements.txt
pip install -r backend/requirements.txt

# Start the backend server (runs on http://localhost:8000)
uvicorn backend.main:app --reload
```

### 2. Start the Frontend (Vite + React)
Open a **new/separate** terminal tab/window, navigate to the `frontend` folder, and run:
```bash
# Change directory to frontend
cd frontend

# Install frontend dependencies
npm install

# Start the frontend dev server (runs on http://localhost:5173)
npm run dev
```

### 3. Access the Dashboard
Once both servers are running, open your web browser and navigate to:
**[http://localhost:5173](http://localhost:5173)**

---

## When you get your real files


1. **Excel roster** → put it anywhere, then set `excel_path` in `config.yaml`.
   Update the `columns:` section to match your actual column headers — you don't need
   to rename your Excel columns, just point the config at whatever they're called.

2. **Template** → two cases:

   **Case A — you get a Word (.docx) design.**
   Open it and manually type `{{ student_name }}`, `{{ event_name }}`, `{{ date }}`
   etc. wherever those values should appear (ask the designer to leave those spots as
   plain editable text, not inside a text box/image). Set:
   ```yaml
   template_mode: "docx"
   docx_template_path: "path/to/real_template.docx"
   ```

   **Case B — you get a Canva/Photoshop/PDF design with no editable text.**
   PDF works directly — no need to convert it yourself:
   ```yaml
   template_mode: "image"
   image_template_path: "path/to/real_template.pdf"   # .png / .jpg also work
   image_render_dpi: 300     # keep this the SAME every time you touch coordinates
   ```
   Run `python find_coordinates.py`, open `output/grid_preview.png`, read off the
   x,y pixel coordinates where the name/date/etc. should sit, and fill those into
   `image_text_fields:` in config.yaml. Takes about 5 minutes of trial and error —
   re-run `generate_certificates.py` after each tweak and check one output PDF.

   **Important:** coordinates are tied to `image_render_dpi`. If you change that
   number later, every x/y you already picked will land in the wrong spot — pick a
   DPI once (300 is print-quality and a safe default) and leave it alone.

3. **Email** → fill in `email.sender_email` and `email.sender_app_password`
   (Gmail: create one at https://myaccount.google.com/apppasswords — you need
   2-Step Verification turned on first). Leave `dry_run: true` until you've checked
   the sample output.

## Running for real

```bash
python generate_certificates.py   # regenerate for your real roster + template
python send_emails.py             # still dry-run — check your inbox
# once it all looks right:
```
Open `config.yaml`, set `email.dry_run: false`, then run `python send_emails.py` again.
This sends the real emails — one per POC, grouped automatically by the `poc_email`
column, each with all their students' certificates attached.

## Notes / gotchas

- **Re-running is safe.** Certificates are named `StudentName__ClusterID.pdf`, so
  fixing one row and re-running regenerates only what changed (existing files are
  overwritten, nothing duplicates).
- **Blank POC emails** are automatically skipped with a warning printed — fix them
  in the Excel and re-run.
- **Gmail sending limit** is ~500/day on a free account — fine for a single event,
  but if your list is huge, use your university's SMTP relay instead (`smtp_host` /
  `smtp_port` in config.yaml).
- **manifest.csv** in `output/` is your audit trail — one row per student showing
  exactly which PDF was generated and which POC it'll go to. Open it in Excel before
  sending if you want a final human check.
