# Traffic Notice Generator — Python edition

Replaces `TrafficNoticeGenerator_FIXED_V3_2_4.ps1` (PowerShell + Word COM
automation). Same approved template, same output, much faster, and no
dependency on a licensed copy of MS Word — which also means this same code
can run on a server later for the "upload a CSV, get PDFs back" website.

## Why it's faster

The old script opened Microsoft Word through COM **once per pair of
challans** (Open → set page geometry → export PDF → close), tens of
thousands of times for a big run. That per-call overhead was the actual
bottleneck, not the printing itself.

This version:
1. Fills in the exact same approved `.docx` template (byte-for-byte,
   nothing redrawn or re-laid-out) for every pair of challans, in memory.
2. Stitches every pair in a batch (default: 1000 challans = 500 pairs)
   into **one** combined `.docx`.
3. Converts that one file to PDF with **one** headless LibreOffice call.
4. Verifies the page count came out exactly right before accepting the
   batch (same safety check the old script had — if a record's text would
   spill a notice onto a third page, that batch is flagged, not silently
   shipped).

Optionally, independent batches can be converted **in parallel** across
your CPU cores (`--workers N`) since they don't depend on each other.

## Setup

1. Install Python 3.10+ (already fine on your machine).
2. Install **LibreOffice** (free): https://www.libreoffice.org/download/download/
   — this replaces MS Word entirely for this tool. No license needed.
3. Install the two Python packages this uses:
   ```
   pip install openpyxl pypdf
   ```
   (`psutil` is optional, only used by `benchmark.py` to report memory use:
   `pip install psutil`)

## Usage

```
python generate_notices.py --input daily.xlsx --sheet Bike --start-row 8 \
       --template templates/Formate_with_Sign_SectionSplit.docx \
       --output out/ --batch-size 1000 --workers 4 --zip
```

Or with a CSV (this is the format the future website will accept — a
header row with these exact column names, any order):

```
SERIAL_NUMBER,E_CHALLAN_NUMBER,VEHICLE_NUMBER,VIOLATION_DATE,NAME,ADDRESS,MOBILE_NUMBER,FINE_AMOUNT,SECTION
```

```
python generate_notices.py --input challans.csv \
       --template templates/Formate_with_Sign_SectionSplit.docx \
       --output out/ --batch-size 1000 --workers 4 --zip
```

Key options:
- `--batch-size` — challans per output PDF (1000 by default, matches what
  you asked for: "1000 challans in one pdf").
- `--workers` — how many batches to convert at the same time. Set this to
  your CPU's core count (check with Task Manager → Performance → Cores).
  Your i5-10th-gen-U is typically 4 cores / 8 threads, so try `--workers 4`.
- `--zip` — also bundles every batch PDF from this run into one
  `Traffic_Notices_all_batches.zip`, so you can choose zip-or-not per run
  as you asked.

Nothing crashes the whole run: each batch is generated and verified
independently, so if (say) batch 47 out of 70 fails, the other 69 still
complete and you get a clear error for just that one to look into.

## What each batch PDF is

Each PDF has the front (2 challans per page, top half / bottom half — the
"tear line" layout you described) followed immediately by the matching
back page for the same pair, repeated pair after pair — identical layout
and page order to what the PowerShell tool produced, just generated much
faster.

## Verifying fidelity before a real run

Before trusting this on real data:
1. Run it once on the small `tests/Traffic_Notice_Test_4_Records.xlsx`
   file.
2. Open the resulting PDF and compare it side-by-side against a PDF made
   the old way from the same 4 records. They should be visually
   indistinguishable (I already did this during development — see
   `TESTING.md`).
3. Only then run it against a full daily/weekly file.

## Next step you mentioned: the upload website

Once you're happy with this, the same `notice_lib.py` (`generate_all()`)
drops straight into a small Flask/FastAPI app: a page that accepts a CSV
upload, calls `generate_all()`, and returns the zip. No Word/Windows
dependency, so it can run on any small Linux server too. Happy to build
that next once you've load-tested this on your laptop.
