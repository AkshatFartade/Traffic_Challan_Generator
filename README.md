# Traffic Notice Generator (Web & CLI)

A high-performance, deterministic legal notice generation system for traffic e-challans. It replaces legacy Microsoft Word COM automation with in-memory DOCX templating and headless LibreOffice conversion, delivering orders-of-magnitude faster generation without any Microsoft Office license requirements.

---

## Key Features

- **Web & CLI Interfaces**: Run locally as a production web service (`Waitress` WSGI server with modern UI) or run batches directly via the CLI (`generate_notices.py`).
- **Template Fidelity**: Uses the official, byte-pinned `.docx` template (`templates/Formate_with_Sign_SectionSplit.docx`). Every generated notice matches the court-approved layout, typography, and Hindi Noto Sans Devanagari font rendering.
- **Three Strict Verification Gates**:
  - **Gate 1 (Preflight)**: Inspects uploaded spreadsheets/CSVs, validates date formats (`DD/MM/YYYY`, `DD-MM-YYYY`, ISO datetimes), enforces column mapping, and blocks duplicate challan numbers.
  - **Gate 2 (Engine Check)**: Verifies exact page counts per batch (2 notices per page front and back = 1 page pair per 2 challans).
  - **Gate 3 (G3 Forensic Audit)**: Uses `pdfplumber` to audit text positions (serial numbers, challans, vehicle registrations, mobile numbers) ensuring top and bottom notices never leak across the tear-line divider.
- **Odd-Count Handling**: Deterministically blanks out the bottom half on the final page when the total record count is odd, preserving page layout, borders, and tear-off dividers.
- **Multi-Worker Scaling**: Generates and converts independent batches in parallel (default: 4 workers) with zero race conditions.
- **Complete Packaging**: Outputs individual batch PDFs as well as a consolidated ZIP bundle containing a cryptographic `manifest.txt` (with SHA-256 hashes, timestamps, and row counts).

---

## System Requirements

1. **Operating System**: Windows 10/11, Windows Server, or Linux.
2. **Python**: Python 3.10+ (tested up to Python 3.13).
3. **LibreOffice**: Free, open-source office suite used headlessly for high-speed PDF export.
   - Download: [https://www.libreoffice.org/download/download/](https://www.libreoffice.org/download/download/)
   - Default Windows path: `C:\Program Files\LibreOffice\program\soffice.exe` (auto-detected).
   - If installed elsewhere, set the `SOFFICE_PATH` environment variable.

---

## Installation & Setup

1. **Clone or extract** the project into your desired folder:
   ```cmd
   cd a:\TWE_Challans_Generator\notice_generator
   ```

2. **Create and activate a virtual environment**:
   ```cmd
   python -m venv .venv
   .\.venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```cmd
   pip install -r requirements.txt
   ```

---

## Running the Web Interface

### Option A: Using the Windows Batch Script
1. Set your supervisor password:
   ```cmd
   set CHALLAN_SUPERVISOR_PASSWORD=your_secure_password
   ```
2. Double-click or run `run_site.bat`:
   ```cmd
   run_site.bat
   ```
3. Open your browser and navigate to:
   ```
   http://127.0.0.1:5000
   ```

### Option B: Using PowerShell / CLI
```powershell
$env:CHALLAN_SUPERVISOR_PASSWORD = "your_secure_password"
.\.venv\Scripts\python.exe -m waitress --listen=127.0.0.1:5000 --call "app:create_app"
```

### Web Workflow
1. **Login**: Enter the supervisor password configured in `CHALLAN_SUPERVISOR_PASSWORD`.
2. **Upload**: Drag-and-drop or select an Excel (`.xlsx`) or CSV file.
   - The system automatically detects sheet layouts (e.g., `Sheet1` row 8 or row 2, `Bike` row 8).
   - If multiple sheets or custom layouts exist, select the desired sheet and header row.
3. **Preflight**: Review Gate 1 preflight checks (record counts, detected columns, non-blocking warnings).
4. **Configure & Start**: Set batch size (default: 1000 notices = 500 pages per PDF) and worker count (recommended: 4). Click **Start Generation**.
5. **Real-Time Monitoring**: Watch progress updates per batch as documents are generated and converted.
6. **Download Results**: Download individual batch PDFs or download the complete ZIP archive with verification manifest.

---

## Running via Command Line (CLI)

Use `generate_notices.py` to batch-process files without starting the web server.

### Example 1: Standard Excel Input
```cmd
python generate_notices.py --input "Bike_19-10-26.xlsx" --sheet Bike --start-row 8 --batch-size 1000 --workers 4 --zip
```

### Example 2: CSV Input
```cmd
python generate_notices.py --input "challans.csv" --batch-size 1000 --workers 4 --output "output/" --zip
```

### CLI Arguments Reference

| Flag | Default | Description |
|---|---|---|
| `--input`, `-i` | *(Required)* | Path to input `.xlsx` or `.csv` file |
| `--template`, `-t` | `templates/Formate_with_Sign_SectionSplit.docx` | Path to approved `.docx` template |
| `--output`, `-o` | `output/` | Destination directory for generated PDFs |
| `--batch-size`, `-b` | `1000` | Number of notices per output PDF (500 pages) |
| `--workers`, `-w` | `4` | Number of parallel LibreOffice conversion workers |
| `--sheet`, `-s` | `Bike` | Sheet name when using an Excel file |
| `--start-row`, `-r` | `8` | 1-based row number where records begin in Excel |
| `--zip` | `False` | Bundles all batch PDFs into a single `.zip` archive |

---

## Input Data Format

The system expects the following 9 standard columns (in any column order):

| Column Name | Description | Example |
|---|---|---|
| `SERIAL_NUMBER` | Sequential notice/register number | `8787` |
| `E_CHALLAN_NUMBER` | Unique e-challan reference ID | `CG5010992607180001` |
| `VEHICLE_NUMBER` | Vehicle registration plate number | `CG07AB1234` |
| `VIOLATION_DATE` | Date/time of violation | `18/07/2026`, `18-07-2026`, or `2026-07-18T14:30:00` |
| `NAME` | Violator / vehicle owner name | `RAMESH KUMAR` |
| `ADDRESS` | Postal delivery address | `HOUSE NO 12, SECTOR 6, BHILAI, DURG, CG 490006` |
| `MOBILE_NUMBER` | Contact mobile number | `9826100000` |
| `FINE_AMOUNT` | Fine penalty amount in INR | `500` or `1000` |
| `SECTION` | Motor Vehicles Act section code & description | `129/194(D)Driving without helmet` |

---

## Environment Variables

| Variable | Default | Purpose |
|---|---|---|
| `CHALLAN_SUPERVISOR_PASSWORD` | `""` | Password required to log in to the web interface |
| `CHALLAN_HOST` | `127.0.0.1` | Host address for web server |
| `CHALLAN_PORT` | `5000` | Port for web server |
| `CHALLAN_WORKERS` | `4` | Default worker threads for background generation |
| `CHALLAN_JOB_RETENTION_MINUTES` | `60` | Auto-cleanup interval for completed jobs |
| `SOFFICE_PATH` | Auto-detected | Path to LibreOffice `soffice.exe` executable |

---

## Running the Automated Test Suite

The test suite covers unit tests, preflight validation, odd-count math, engine page conversions, adversarial web inputs, and forensic G3 verifications:

```powershell
.\.venv\Scripts\pytest.exe tests/
```

Key test modules:
- `tests/test_preflight.py`: Header parsing, date validation, and duplicate blocking.
- `tests/test_odd_count.py`: Validates odd record handling (1, 3, 999 records) and blank half assertions.
- `tests/test_verifier.py`: Verification gates (G1, G2, G3), including negative tests with deliberately swapped records.
- `tests/test_comprehensive.py`: Scale runs (1000, 1001 records), boundary conditions, Hindi Unicode, and special characters.
- `tests/test_web_adversarial.py`: Web API edge cases, invalid tokens, oversized payloads, and session security.

---

## Architecture & Code Structure

```
notice_generator/
│
├── app.py               # Flask app factory, template SHA-256 integrity verification, cleanups
├── auth.py              # Session authentication and login protection
├── config.py            # Central configuration & environment variable resolver
├── generate_notices.py  # Standalone CLI batch runner
├── input_inspector.py   # Gate 1 preflight checker and automatic layout detection
├── job_runner.py        # Background job threadpool orchestrator & status reporter
├── notice_lib.py        # Core DOCX XML injection, stitching, and LibreOffice PDF conversion
├── requirements.txt     # Python package requirements
├── routes.py            # Web API endpoints (upload, preflight, start, poll, download)
├── run_site.bat         # Single-click Windows launcher
├── verifier.py          # Gates G1, G2, and G3 forensic pdfplumber position auditor
│
├── static/              # CSS styles, design tokens, JavaScript, and Hindi fonts
├── templates/           # Pinned DOCX templates (Formate_with_Sign_SectionSplit.docx)
├── templates_html/      # Jinja2 HTML templates for the web interface
└── tests/               # Pytest suite and golden reference files
```

---

## License & Operational Notes

Designed for official traffic enforcement notice distribution. All template layouts and output formatting are cryptographically pinned to guarantee compliance with legal dispatch requirements.
