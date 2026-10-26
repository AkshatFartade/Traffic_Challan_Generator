# HANDOVER.md — Traffic Notice Generator

> **Print this page** and keep it with the machine. Operators need this to start, stop, and
> recover the system. Developers need it to run the test suite and investigate layout problems.

---

## 1. What This System Does

Converts an official e-challan Excel register (`.xlsx`) or CSV into stamped, two-sided A4 PDF
traffic notices. Every output is audited by three verification gates (G1 preflight, G2 page count,
G3 word-position) before the Download button is enabled. The template and signature image are never
regenerated — only placeholder values are substituted.

---

## 2. Clean-PC Setup (new machine or rebuild)

### 2a. Prerequisites

| Software | Minimum | Notes |
|---|---|---|
| Python | 3.10+ | <https://python.org> |
| LibreOffice | 7.6+ | <https://libreoffice.org> — **must be installed system-wide** |
| Git | any | for version control |

### 2b. Install

```bat
rem 1. Clone or copy the project folder
cd A:\TWE_Challans_Generator\notice_generator

rem 2. Create virtual environment
python -m venv .venv

rem 3. Activate
.venv\Scripts\activate

rem 4. Install Python dependencies
pip install flask waitress openpyxl pdfplumber pypdfium2 pypdf python-docx psutil

rem 5. Verify LibreOffice is detectable
python -c "import config; print('soffice:', config.SOFFICE_PATH)"
```

### 2c. Fonts (required for Hindi text in PDFs)

The Noto Sans Devanagari fonts are bundled in `static/fonts/`. LibreOffice must be able to access
them. On a fresh machine:

1. Open `static/fonts/` in Explorer.
2. Right-click `NotoSansDevanagari-Regular.ttf` → **Install for all users**.
3. Repeat for `NotoSansDevanagari-Bold.ttf`.
4. Restart LibreOffice if it was open.

---

## 3. Supervisor Password

The web interface requires a password (set by the senior police supervisor, never committed to code).

**Set the password before starting the server:**

```bat
set CHALLAN_SUPERVISOR_PASSWORD=ChooseAStrongPassphrase
```

Or put it in `.env` and source it (never commit `.env` — it is in `.gitignore`).

> If you forget the passphrase, **delete `.session_secret`** and set a new one.  
> All active browser sessions will be logged out.

### Password rules
- Minimum 12 characters recommended.
- No special characters needed; a 4-word pass-phrase works well.
- Change it when any supervisor leaves.

---

## 4. Start / Stop

### Start (daily use)

```bat
rem Option A — double-click run_site.bat (reads CHALLAN_SUPERVISOR_PASSWORD from environment)
run_site.bat

rem Option B — command line with explicit password
set CHALLAN_SUPERVISOR_PASSWORD=YourPassphrase
.venv\Scripts\python -m waitress --listen=127.0.0.1:5000 --call "app:create_app"
```

Browser: <http://127.0.0.1:5000>

### Stop

Press **Ctrl+C** in the terminal window. The server stops gracefully.  
LibreOffice child processes (used during generation) are killed automatically.

### LAN access (optional)

```bat
set CHALLAN_HOST=0.0.0.0
```

Then open port 5000 on Windows Firewall. All other machines on the same network can connect.

---

## 5. Job Retention / Cleanup

- Generated PDFs and uploaded files live in `jobs/<job_id>/`.
- Default retention: **60 minutes** from last modification time.
- Cleanup runs **at startup** and can be changed:

```bat
set CHALLAN_JOB_RETENTION_MINUTES=120
```

To manually purge all jobs immediately:

```bat
rd /s /q jobs
mkdir jobs
```

---

## 6. Running the Golden Regression Test

The golden test compares a freshly generated 2-record PDF against `tests/golden/snapshot.pdf`
(committed to git). Run it after any change to `notice_lib.py` or the template:

```bat
.venv\Scripts\python -m pytest tests\test_golden.py -v
```

**Expected output:** `1 passed`

If it fails:
1. Open `tmp/test_golden/out_golden/` and inspect the generated PDF visually.
2. If the new PDF is *correct* and the template was intentionally updated, regenerate the snapshot:
   ```bat
   .venv\Scripts\python scripts\generate_snapshot.py
   ```
3. Commit the new snapshot with a note explaining why it changed.

---

## 7. Running All Tests

```bat
.venv\Scripts\python -m pytest tests\ -v --tb=short
```

Full suite (13 base + comprehensive) takes ~5–10 minutes (LibreOffice renders ~20 PDFs).

### Skipping slow scale tests (quick check)

```bat
.venv\Scripts\python -m pytest tests\ -v --ignore=tests\test_comprehensive.py -k "not scale"
```

---

## 8. What to Do If a Notice Looks Wrong

Work through this checklist in order:

### 8a. Check the preflight report first
Upload the file via the website. Read the **Check Report** page carefully — all blocking errors
must be zero before generation runs.

### 8b. Check the PDF visually
After download, open the PDF and compare one notice against the Excel row:
- Serial Number, Challan Number, Vehicle Number, Date, Fine in the **front notice** ✓
- Name, Address, Mobile, Section description in the **back slip** ✓

### 8c. Run the G3 word-position audit manually

```python
import verifier, notice_lib as nl
records = nl.read_records("original_files/Bike_19-10-26.xlsx")
# Generate first
run = nl.generate_all("original_files/Bike_19-10-26.xlsx",
                      "templates/Formate_with_Sign_SectionSplit.docx",
                      "output/debug_run", batch_size=50, workers=1)
result = verifier.verify_all_gates([], run.batches, records)
print(result.passed, result.errors[:10])
```

### 8d. Check for data corruption in the Excel file
Common causes:
| Symptom | Likely Cause |
|---|---|
| Number printed as `5e+09` | Mobile stored as float/scientific in Excel |
| Date printed as `43831` | Date cell not formatted, value is an Excel serial |
| Garbled text | Non-UTF-8 characters copied from WhatsApp/PDF |
| Red placeholder text | Generation ran against the template without data substitution |

### 8e. Restore the template if it was accidentally modified

```bat
git checkout baseline-working -- templates/Formate_with_Sign_SectionSplit.docx
```

The pinned SHA-256 hash in `config.py` (`PINNED_TEMPLATE_HASH`) will prevent the server from
starting if the template hash does not match. The correct hash is:

```
2b3c563af36b29526bcf467f5a11a9fdff6361740f9990002e6ea7e9e4e8ff8b
```

### 8f. Re-run the golden test
```bat
.venv\Scripts\python -m pytest tests\test_golden.py -v
```

If this fails after restoring the template, contact the developer. Do not issue notices until
the golden test passes.

---

## 9. Security Notes

- The server binds to `127.0.0.1` by default — **only accessible from the same machine**.
- Session cookies are `HttpOnly`, `SameSite=Strict`, never sent over plain HTTP unless
  `SESSION_COOKIE_SECURE=False` (default for LAN use).
- The download endpoint is gated: **verified=True** must be in `status.json` before any file
  is served. This cannot be bypassed.
- Job IDs are 32-character hex tokens — not guessable.
- Uploaded files and generated PDFs are auto-deleted after retention period.
- Log files contain batch indices and timings only — **no names, addresses, or mobile numbers**.

---

## 10. Quick Reference Card

| Action | Command |
|---|---|
| Start server | `run_site.bat` |
| Run all tests | `.venv\Scripts\python -m pytest tests\ -v` |
| Run golden test only | `.venv\Scripts\python -m pytest tests\test_golden.py -v` |
| Restore template | `git checkout baseline-working -- templates\Formate_with_Sign_SectionSplit.docx` |
| Check template hash | `Get-FileHash templates\Formate_with_Sign_SectionSplit.docx -Algorithm SHA256` |
| Purge jobs manually | `rd /s /q jobs && mkdir jobs` |
| Change password | Set `CHALLAN_SUPERVISOR_PASSWORD` in environment |
| Inspect a PDF page count | `.venv\Scripts\python -c "import notice_lib as nl; print(nl.count_pdf_pages('output/file.pdf'))"` |

---

*Last updated: 2026-10-26. Generator version: commit `89849e6` + `27dd452`.*
