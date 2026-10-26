"""
notice_lib.py
=============
Core engine for the Traffic Notice / Challan PDF generator.

This ports the logic that used to live in TrafficNoticeGenerator_FIXED_V3_2_4.ps1
(Excel -> Word COM -> qpdf) into pure Python, but restructures the slow part:

  OLD: for every pair of challans -> open Word (COM) -> export 1 PDF -> repeat
       thousands of times -> merge all the small PDFs with qpdf.
       (thousands of COM Open/Export/Close calls = the actual bottleneck)

  NEW: for every BATCH of challans (e.g. 1000) -> build ONE docx containing
       every pair's front+back sections back to back -> ONE headless
       LibreOffice conversion call for the whole batch -> done.
       Word/COM is not used at all, so this also works on a server (for the
       future "upload a CSV, get PDFs back" website), not just a desktop PC
       with MS Office licensed on it.

The actual approved Word template (Formate_with_Sign_SectionSplit.docx) is
used untouched byte-for-byte except for the {{PLACEHOLDER}} substitutions --
so the printed layout, fonts, borders, and signature image are exactly what
was already approved. Nothing is redrawn or re-laid-out by hand.
"""

from __future__ import annotations

import csv
import datetime
import io
import json
import logging
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

logger = logging.getLogger("notice_lib")

FIELD_KEYS = [
    "SERIAL_NUMBER",
    "E_CHALLAN_NUMBER",
    "VEHICLE_NUMBER",
    "VIOLATION_DATE",
    "NAME",
    "ADDRESS",
    "MOBILE_NUMBER",
    "FINE_AMOUNT",
    "SECTION",
]

EMPTY_RECORD = {k: "" for k in FIELD_KEYS}

_PLACEHOLDER_RE = re.compile(r"\{\{(" + "|".join(FIELD_KEYS) + r")\}\}")
_SECTPR_RE = re.compile(r"<w:sectPr\b[^>]*>.*?</w:sectPr>", re.DOTALL)
_RUN_WITH_DATA_RE = re.compile(
    r"<w:r\b[^>]*>.*?\[\[DATA\]\].*?\[\[/DATA\]\].*?</w:r>", re.DOTALL
)
_SZ_RE = re.compile(r'<w:sz\s+w:val="\d+"\s*/>')
_SZCS_RE = re.compile(r'<w:szCs\s+w:val="\d+"\s*/>')

# Section-number / description split, ported from the PS1 Split-SectionText regex.
# Examples handled the same way as before:
#   129/194(D)Driving without helmet -> "129/194(D)" / "Driving without helmet"
#   190(2). For 2 wheeler ...        -> "190(2)"     / "For 2 wheeler ..."
_SECTION_SPLIT_RE = re.compile(
    r"^\s*(\(?\d+(?:\(\w+\))?(?:/\d+(?:\(\w+\))?)?\)?)(?:\s*\.\s*|\s+|(?=[A-Za-z]))(.*)$",
    re.IGNORECASE,
)

_XML_ESCAPES = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&apos;",
}


def escape_xml(value) -> str:
    if value is None:
        return ""
    s = str(value)
    return re.sub(r"[&<>\"']", lambda m: _XML_ESCAPES[m.group(0)], s)


def format_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


_DATE_FORMATS = [
    "%d-%m-%Y %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%Y-%m-%d",
]


def format_date_value(value) -> str:
    """Mirrors Format-DateValue from the PS1 script: accepts Excel serials,
    datetime objects, or a handful of common string formats, and always
    renders dd/MM/yyyy (or returns the original text if nothing matches)."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, (int, float)):
        try:
            # Excel serial date epoch (matches .NET DateTime.FromOADate)
            epoch = datetime.datetime(1899, 12, 30)
            return (epoch + datetime.timedelta(days=float(value))).strftime("%d/%m/%Y")
        except Exception:
            pass
    s = str(value).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s


def split_section_text(value) -> dict:
    s = "" if value is None else str(value).strip()
    if not s:
        return {"number": "", "description": ""}
    front_parts, back_parts = [], []
    for raw_part in s.split("|"):
        part = raw_part.strip()
        if not part:
            continue
        m = _SECTION_SPLIT_RE.match(part)
        if m:
            front_parts.append(m.group(1).strip())
            desc = m.group(2).strip()
            if desc:
                back_parts.append(desc)
        else:
            # Unrecognized format: keep the whole value on the front rather
            # than silently dropping official data.
            front_parts.append(part)
    return {"number": " | ".join(front_parts), "description": " | ".join(back_parts)}


# --------------------------------------------------------------------------
# Reading challan records
# --------------------------------------------------------------------------

def read_records_csv(path: Path) -> list[dict]:
    """CSV must have a header row with these exact column names (any order):
    SERIAL_NUMBER, E_CHALLAN_NUMBER, VEHICLE_NUMBER, VIOLATION_DATE, NAME,
    ADDRESS, MOBILE_NUMBER, FINE_AMOUNT, SECTION
    This is the format the future web upload tool will use."""
    records = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        missing = [k for k in FIELD_KEYS if k not in (reader.fieldnames or [])]
        if missing:
            raise ValueError(f"CSV is missing required column(s): {', '.join(missing)}")
        for row in reader:
            serial = format_text(row.get("SERIAL_NUMBER"))
            challan = format_text(row.get("E_CHALLAN_NUMBER"))
            if not serial and not challan:
                continue
            records.append(
                {
                    "SERIAL_NUMBER": serial,
                    "E_CHALLAN_NUMBER": challan,
                    "VEHICLE_NUMBER": format_text(row.get("VEHICLE_NUMBER")),
                    "VIOLATION_DATE": format_date_value(row.get("VIOLATION_DATE")),
                    "NAME": format_text(row.get("NAME")),
                    "ADDRESS": format_text(row.get("ADDRESS")),
                    "MOBILE_NUMBER": format_text(row.get("MOBILE_NUMBER")),
                    "FINE_AMOUNT": format_text(row.get("FINE_AMOUNT")),
                    "SECTION": format_text(row.get("SECTION")),
                }
            )
    return records


def read_records_excel(path: Path, sheet_name: str = "Bike", start_row: int = 8) -> list[dict]:
    """Backward-compatible with the existing daily Excel export: columns
    B..J starting at start_row (default row 8, same as the PS1 tool)."""
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' was not found in {path.name}.")
    ws = wb[sheet_name]
    records = []
    for row in ws.iter_rows(min_row=start_row, min_col=2, max_col=10):
        vals = [c.value for c in row]
        serial = format_text(vals[0])
        challan = format_text(vals[1])
        if not serial and not challan:
            continue
        records.append(
            {
                "SERIAL_NUMBER": serial,
                "E_CHALLAN_NUMBER": challan,
                "VEHICLE_NUMBER": format_text(vals[2]),
                "VIOLATION_DATE": format_date_value(vals[3]),
                "NAME": format_text(vals[4]),
                "ADDRESS": format_text(vals[5]),
                "MOBILE_NUMBER": format_text(vals[6]),
                "FINE_AMOUNT": format_text(vals[7]),
                "SECTION": format_text(vals[8]),
            }
        )
    wb.close()
    return records


def read_records(path, sheet_name: str = "Bike", start_row: int = 8) -> list[dict]:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return read_records_csv(path)
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return read_records_excel(path, sheet_name=sheet_name, start_row=start_row)
    raise ValueError(f"Unsupported input file type: {path.suffix}")


# --------------------------------------------------------------------------
# Template handling (ported from Get-TemplateParts / Replace-Pair / etc.)
# --------------------------------------------------------------------------

@dataclass
class TemplateParts:
    original_xml: str
    body_open_end: int
    body_close_start: int
    content: str          # front+back content for ONE pair, sectPr for section 1 embedded, no trailing sectPr
    final_section: str    # the raw <w:sectPr>...</w:sectPr> for the pair's 2nd (back) section


def load_template_parts(template_path: Path) -> TemplateParts:
    with zipfile.ZipFile(template_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")

    body_open = re.search(r"<w:body[^>]*>", xml)
    body_close = re.search(r"</w:body>", xml)
    if not body_open or not body_close:
        raise ValueError("Invalid Word template: no <w:body> found.")

    bo_end = body_open.end()
    bc_start = body_close.start()
    inner = xml[bo_end:bc_start]

    sect_matches = list(_SECTPR_RE.finditer(inner))
    if len(sect_matches) < 2:
        raise ValueError("The approved template must contain the front and back sections.")

    final_sect = sect_matches[-1].group(0)
    content = inner[: sect_matches[-1].start()]
    # Placeholders in the approved template are marked in red; once real data
    # is inserted it must print as normal black official text.
    content = content.replace('<w:color w:val="C00000"/>', '<w:color w:val="000000"/>')

    return TemplateParts(
        original_xml=xml,
        body_open_end=bo_end,
        body_close_start=bc_start,
        content=content,
        final_section=final_sect,
    )


def replace_pair(template_content: str, rec1: dict, rec2: dict) -> str:
    counts = {k: 0 for k in FIELD_KEYS}

    def repl(m: re.Match) -> str:
        key = m.group(1)
        counts[key] += 1
        n = counts[key]

        if key == "SECTION":
            # SECTION appears 4 times in the signed template:
            # 1 = record1 front, 2 = record2 front, 3 = record1 back, 4 = record2 back.
            rec = rec1 if n in (1, 3) else rec2
            split = split_section_text(rec.get("SECTION"))
            val = split["number"] if n in (1, 2) else split["description"]
            return "[[DATA]]" + escape_xml(val) + "[[/DATA]]"

        if key in ("NAME", "ADDRESS", "MOBILE_NUMBER"):
            # Appear twice on front, twice on back: 1/3 = record1, 2/4 = record2.
            rec = rec1 if n in (1, 3) else rec2
        else:
            # Other fields occur once per front notice.
            rec = rec1 if n == 1 else rec2

        return "[[DATA]]" + escape_xml(rec.get(key)) + "[[/DATA]]"

    return _PLACEHOLDER_RE.sub(repl, template_content)


def normalize_inserted_data_font(content: str, half_points: int = 20) -> str:
    """Apply the typography used by the approved 2_Challans reference.

    - Body/legal text: Noto Sans Devanagari, 11 pt.
    - Inserted challan values: 10 pt, preserving the template's boldness.
    - English/URL runs: Calibri (boldness preserved).
    - Official heading and //नोटिस//: Noto Sans Devanagari, 10 pt.
    - Remove accidental italic formatting.
    """
    def set_rpr(run: str, font: str, size: int) -> str:
        m = re.search(r'<w:rPr\b[^>]*>(.*?)</w:rPr>', run, re.DOTALL)
        rpr = m.group(1) if m else ''
        rpr = re.sub(r'<w:rFonts\b[^>]*/>', '', rpr)
        rpr = re.sub(r'<w:rFonts\b[^>]*>.*?</w:rFonts>', '', rpr, flags=re.DOTALL)
        rpr = re.sub(r'<w:sz\b[^>]*/>', '', rpr)
        rpr = re.sub(r'<w:szCs\b[^>]*/>', '', rpr)
        rpr = re.sub(r'<w:i\b[^>]*/>', '', rpr)
        rpr = re.sub(r'<w:iCs\b[^>]*/>', '', rpr)
        fonts = (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" '
                 f'w:eastAsia="{font}" w:cs="{font}"/>')
        props = fonts + f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        if m:
            return run[:m.start(1)] + props + rpr + run[m.end(1):]
        return run.replace('<w:r>', f'<w:r><w:rPr>{props}</w:rPr>', 1)

    def fix_run(m: re.Match) -> str:
        run = m.group(0)
        texts = re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', run, re.DOTALL)
        text = ''.join(texts)
        if not text:
            return run
        visible = text.replace('[[DATA]]', '').replace('[[/DATA]]', '')
        is_data = '[[DATA]]' in text or '[[/DATA]]' in text
        is_heading = 'कार्यालय अतिरिक्त पुलिस अधीक्षक' in visible
        is_notice_title = 'नोटिस' in visible and len(visible.strip()) <= 16
        is_dev = any('\u0900' <= ch <= '\u097f' for ch in visible)
        is_latin = any(('A' <= ch <= 'Z') or ('a' <= ch <= 'z') for ch in visible)
        font = 'Noto Sans Devanagari' if is_dev and not is_latin else 'Calibri'
        if is_data:
            size = 20
        elif is_heading or is_notice_title:
            size = 20
        else:
            size = 22
        run = set_rpr(run, font, size)
        return run.replace('[[DATA]]', '').replace('[[/DATA]]', '')

    result = re.sub(r'<w:r\b[^>]*>.*?</w:r>', fix_run, content, flags=re.DOTALL)

    def reduce_official_heading(pm: re.Match) -> str:
        para = pm.group(0)
        ptext = ''.join(re.findall(r'<w:t\b[^>]*>(.*?)</w:t>', para, re.DOTALL))
        if ('कार्यालय' in ptext and 'अतिरिक्त' in ptext and 'पुलिस' in ptext and 'अधीक्षक' in ptext) or (ptext.strip() == '//नोटिस//'):
            para = para.replace('<w:sz w:val="22"/>', '<w:sz w:val="20"/>')
            para = para.replace('<w:szCs w:val="22"/>', '<w:szCs w:val="20"/>')
        return para

    return re.sub(r'<w:p\b[^>]*>.*?</w:p>', reduce_official_heading, result, flags=re.DOTALL)



def build_pair_inner(parts: TemplateParts, rec1: dict, rec2: dict, wrap_final_section: bool) -> str:
    content = replace_pair(parts.content, rec1, rec2)
    content = normalize_inserted_data_font(content)
    if wrap_final_section:
        # Not the true last section of the merged document: the section
        # break must be embedded in a paragraph, not a body-level sectPr.
        tail = f"<w:p><w:pPr>{parts.final_section}</w:pPr></w:p>"
    else:
        tail = parts.final_section
    return content + tail


def make_pairs(records: list[dict]) -> list[tuple[dict, dict]]:
    pairs = []
    for i in range(0, len(records), 2):
        r1 = records[i]
        r2 = records[i + 1] if i + 1 < len(records) else EMPTY_RECORD
        pairs.append((r1, r2))
    return pairs


def build_batch_docx(template_path: Path, parts: TemplateParts, pairs: list, out_docx: Path) -> None:
    """Writes ONE docx containing every pair's front+back sections, so the
    whole batch converts to PDF in a single LibreOffice call."""
    n = len(pairs)
    inner_chunks = []
    for i, (r1, r2) in enumerate(pairs):
        is_last = i == n - 1
        inner_chunks.append(build_pair_inner(parts, r1, r2, wrap_final_section=not is_last))
    merged_inner = "".join(inner_chunks)

    xml = parts.original_xml
    new_xml = xml[: parts.body_open_end] + merged_inner + xml[parts.body_close_start :]

    if out_docx.exists():
        out_docx.unlink()
    with zipfile.ZipFile(template_path) as src, zipfile.ZipFile(
        out_docx, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
            dst.writestr(item, data)


# --------------------------------------------------------------------------
# Headless LibreOffice conversion (replaces Word COM entirely)
# --------------------------------------------------------------------------

class ConversionError(RuntimeError):
    pass


def convert_docx_to_pdf(
    docx_path: Path,
    out_dir: Path,
    soffice_path: str = "soffice",
    timeout: int = 600,
    retries: int = 2,
) -> Path:
    """Runs one headless LibreOffice process to turn a docx into a pdf.
    Uses a private, per-call user-profile dir so concurrent/parallel batches
    never fight over a LibreOffice lock file."""
    out_dir.mkdir(parents=True, exist_ok=True)
    expected_pdf = out_dir / (docx_path.stem + ".pdf")

    last_err = None
    for attempt in range(1, retries + 1):
        profile_dir = Path(tempfile.mkdtemp(prefix="lo_profile_"))
        try:
            cmd = [
                soffice_path,
                "--headless",
                "--norestore",
                "--nologo",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(out_dir),
                str(docx_path),
            ]
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=timeout
            )
            if proc.returncode == 0 and expected_pdf.exists():
                return expected_pdf
            last_err = f"soffice exit={proc.returncode} stdout={proc.stdout[-500:]} stderr={proc.stderr[-500:]}"
            logger.warning("Conversion attempt %d failed for %s: %s", attempt, docx_path.name, last_err)
        except subprocess.TimeoutExpired as e:
            last_err = f"timed out after {timeout}s: {e}"
            logger.warning("Conversion attempt %d timed out for %s", attempt, docx_path.name)
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)
        time.sleep(1.5 * attempt)

    raise ConversionError(f"Failed to convert {docx_path} to PDF after {retries} attempts: {last_err}")


def count_pdf_pages(pdf_path: Path) -> int:
    from pypdf import PdfReader

    return len(PdfReader(str(pdf_path)).pages)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

@dataclass
class BatchResult:
    batch_index: int
    challans: int
    pairs: int
    expected_pages: int
    actual_pages: Optional[int] = None
    pdf_path: Optional[str] = None
    ok: bool = False
    seconds: float = 0.0
    error: Optional[str] = None


@dataclass
class RunResult:
    total_challans: int
    batch_size: int
    batches: list = field(default_factory=list)
    total_seconds: float = 0.0
    zip_path: Optional[str] = None

    @property
    def ok_batches(self):
        return [b for b in self.batches if b.ok]

    @property
    def failed_batches(self):
        return [b for b in self.batches if not b.ok]


def _batch_job(start, batch_index, batch_records, template_path, tmp_dir, output_dir, soffice_path):
    pairs = make_pairs(batch_records)
    first_serial = batch_records[0]["SERIAL_NUMBER"] or str(start + 1)
    last_serial = batch_records[-1]["SERIAL_NUMBER"] or str(start + len(batch_records))
    safe_first = re.sub(r"[^A-Za-z0-9_-]", "", str(first_serial)) or str(start + 1)
    safe_last = re.sub(r"[^A-Za-z0-9_-]", "", str(last_serial)) or str(start + len(batch_records))
    out_name = f"Traffic_Notices_batch_{batch_index:04d}_{safe_first}-{safe_last}"
    return {
        "batch_index": batch_index,
        "batch_records": batch_records,
        "pairs": pairs,
        "docx_path": str(tmp_dir / f"{out_name}.docx"),
        "final_pdf": str(output_dir / f"{out_name}.pdf"),
        "template_path": str(template_path),
        "tmp_dir": str(tmp_dir),
        "soffice_path": soffice_path,
    }


def _run_one_job(job: dict) -> BatchResult:
    """Standalone, picklable worker used for both sequential and parallel
    (ProcessPoolExecutor) execution -- reloads the template itself so no
    large shared state needs to cross process boundaries."""
    template_path = Path(job["template_path"])
    tmp_dir = Path(job["tmp_dir"])
    docx_path = Path(job["docx_path"])
    final_pdf = Path(job["final_pdf"])
    pairs = job["pairs"]

    result = BatchResult(
        batch_index=job["batch_index"],
        challans=len(job["batch_records"]),
        pairs=len(pairs),
        expected_pages=len(pairs) * 2,
    )
    bt0 = time.perf_counter()
    try:
        parts = load_template_parts(template_path)
        build_batch_docx(template_path, parts, pairs, docx_path)
        produced_pdf = convert_docx_to_pdf(docx_path, tmp_dir, soffice_path=job["soffice_path"])
        actual_pages = count_pdf_pages(produced_pdf)
        if actual_pages != result.expected_pages:
            raise ConversionError(
                f"page count mismatch: expected {result.expected_pages}, got {actual_pages} "
                "(a record's text is probably overflowing the fixed 2-page layout)"
            )
        shutil.move(str(produced_pdf), str(final_pdf))
        result.actual_pages = actual_pages
        result.pdf_path = str(final_pdf)
        result.ok = True
    except Exception as e:
        result.error = str(e)
    finally:
        docx_path.unlink(missing_ok=True)
        result.seconds = time.perf_counter() - bt0
    return result


def generate_all(
    input_path,
    template_path,
    output_dir,
    batch_size: int = 1000,
    sheet_name: str = "Bike",
    start_row: int = 8,
    make_zip: bool = False,
    soffice_path: str = "soffice",
    tmp_dir=None,
    progress_cb=None,
    workers: int = 1,
) -> RunResult:
    input_path = Path(input_path)
    template_path = Path(template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    tmp_dir = Path(tmp_dir) if tmp_dir else Path(tempfile.mkdtemp(prefix="notice_run_"))
    tmp_dir.mkdir(parents=True, exist_ok=True)

    t0 = time.perf_counter()
    records = read_records(input_path, sheet_name=sheet_name, start_row=start_row)
    if not records:
        raise ValueError("No eligible records were found in the input file.")

    # Fail fast on an obviously broken template before doing any batch work.
    load_template_parts(template_path)
    total = len(records)
    run = RunResult(total_challans=total, batch_size=batch_size)

    jobs = []
    batch_index = 0
    for start in range(0, total, batch_size):
        batch_index += 1
        batch_records = records[start : start + batch_size]
        jobs.append(
            _batch_job(start, batch_index, batch_records, template_path, tmp_dir, output_dir, soffice_path)
        )

    if workers <= 1:
        for job in jobs:
            result = _run_one_job(job)
            run.batches.append(result)
            if progress_cb:
                progress_cb(result, run)
    else:
        import concurrent.futures as cf

        with cf.ProcessPoolExecutor(max_workers=workers) as pool:
            futures = {pool.submit(_run_one_job, job): job["batch_index"] for job in jobs}
            pending = {}
            for fut in cf.as_completed(futures):
                result = fut.result()
                pending[result.batch_index] = result
                if progress_cb:
                    progress_cb(result, run)
        run.batches = [pending[i] for i in sorted(pending)]

    run.total_seconds = time.perf_counter() - t0

    if make_zip and run.ok_batches:
        zip_path = output_dir / "Traffic_Notices_all_batches.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for b in run.ok_batches:
                zf.write(b.pdf_path, arcname=Path(b.pdf_path).name)
        run.zip_path = str(zip_path)

    shutil.rmtree(tmp_dir, ignore_errors=True)
    return run


def run_result_to_dict(run: RunResult) -> dict:
    return {
        "total_challans": run.total_challans,
        "batch_size": run.batch_size,
        "total_seconds": round(run.total_seconds, 2),
        "batches_ok": len(run.ok_batches),
        "batches_failed": len(run.failed_batches),
        "zip_path": run.zip_path,
        "batches": [
            {
                "batch_index": b.batch_index,
                "challans": b.challans,
                "expected_pages": b.expected_pages,
                "actual_pages": b.actual_pages,
                "ok": b.ok,
                "seconds": round(b.seconds, 2),
                "error": b.error,
                "pdf_path": b.pdf_path,
            }
            for b in run.batches
        ],
    }
