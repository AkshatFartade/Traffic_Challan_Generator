"""
input_inspector.py — Inspects input files, auto-detects data blocks,
performs preflight validation gates, and exports flagged row reports.
"""
from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import openpyxl
import notice_lib as nl

EXPECTED_COLUMNS = [
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

# Common challan number regex (e.g. CG..., SYN..., or alphanumeric 10-25 chars)
_CHALLAN_RE = re.compile(r"^[A-Za-z0-9_-]{8,30}$")
_CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


@dataclass
class InspectionResult:
    valid: bool
    file_type: str
    sheet_name: Optional[str] = None
    available_sheets: list[str] = field(default_factory=list)
    start_row: int = 1
    total_records: int = 0
    records: list[dict] = field(default_factory=list)
    blocking_errors: list[dict] = field(default_factory=list)
    warnings: list[dict] = field(default_factory=list)
    first_5: list[dict] = field(default_factory=list)
    last_5: list[dict] = field(default_factory=list)
    error_message: Optional[str] = None
    needs_sheet_selection: bool = False


def check_magic_bytes(path: Path) -> str:
    """Verifies file type using magic bytes and rejects unsupported files."""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise ValueError(
            "Legacy Excel 97-2003 (.xls) format is not supported. "
            "Please open the file and Save As modern Excel (.xlsx) or CSV (.csv)."
        )

    with open(path, "rb") as f:
        header = f.read(8)

    if suffix == ".xlsx":
        # Standard ZIP header
        if not header.startswith(b"PK\x03\x04"):
            raise ValueError("Invalid .xlsx file: corrupt header or not a valid Excel file.")
        return "xlsx"
    elif suffix == ".csv":
        return "csv"
    else:
        raise ValueError(f"Unsupported file type '{suffix}'. Please upload a .xlsx or .csv file.")


def detect_excel_data_block(path: Path, preferred_sheet: Optional[str] = None) -> tuple[str, int, list[str]]:
    """Scans all sheets in the Excel workbook to identify qualifying data blocks.
    Returns (selected_sheet, start_row, qualifying_sheets)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    qualifying = []
    sheet_start_rows = {}

    for name in wb.sheetnames:
        ws = wb[name]
        found_start = None
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=25, min_col=2, max_col=10, values_only=True), start=1):
            if not row or len(row) < 2:
                continue
            val_b, val_c = row[0], row[1]
            if val_b is None or val_c is None:
                continue

            str_b = str(val_b).strip()
            str_c = str(val_c).strip()

            # Check if this row is the column header text
            if str_b.upper() in ("SERIAL_NUMBER", "SERIAL NO", "SL NO", "S.NO.", "SR NO") or "CHALLAN" in str_c.upper():
                found_start = r_idx + 1
                break

            # Check if this row is the first row of real data:
            # Column B is integer-like serial and Column C looks like a challan code
            is_serial = False
            if isinstance(val_b, int) or (isinstance(val_b, float) and val_b.is_integer()):
                is_serial = True
            elif str_b.isdigit():
                is_serial = True

            if is_serial and _CHALLAN_RE.match(str_c):
                found_start = r_idx
                break

        if found_start is not None:
            qualifying.append(name)
            sheet_start_rows[name] = found_start

    wb.close()

    if not qualifying:
        raise ValueError("Could not find a valid traffic notice data block in any sheet. Please ensure columns B..J contain challan data.")

    if preferred_sheet and preferred_sheet in sheet_start_rows:
        return preferred_sheet, sheet_start_rows[preferred_sheet], qualifying

    if len(qualifying) == 1:
        chosen = qualifying[0]
        return chosen, sheet_start_rows[chosen], qualifying

    # Multiple qualifying sheets
    return qualifying[0], sheet_start_rows[qualifying[0]], qualifying


def run_preflight(raw_records: list[dict], start_excel_row: int = 1) -> tuple[list[dict], list[dict]]:
    """Runs all BLOCKING and WARNING preflight checks on records.
    Never auto-corrects values. Returns (blocking_errors, warnings)."""
    blocking = []
    warnings = []
    seen_challans = {}
    seen_serials = {}

    for idx, r in enumerate(raw_records):
        excel_row = start_excel_row + idx
        challan = r.get("E_CHALLAN_NUMBER", "")
        serial = r.get("SERIAL_NUMBER", "")
        veh = r.get("VEHICLE_NUMBER", "")
        vdate = r.get("VIOLATION_DATE", "")
        name = r.get("NAME", "")
        addr = r.get("ADDRESS", "")
        mobile = r.get("MOBILE_NUMBER", "")
        fine = r.get("FINE_AMOUNT", "")
        sec = r.get("SECTION", "")

        # --- BLOCKING CHECKS ---
        # 1. Missing challan number
        if not challan:
            blocking.append({"row": excel_row, "field": "E_CHALLAN_NUMBER", "reason": "Missing challan number", "value": ""})
        else:
            # Duplicate challan number
            if challan in seen_challans:
                blocking.append({
                    "row": excel_row,
                    "field": "E_CHALLAN_NUMBER",
                    "reason": f"Duplicate challan number (first seen on row {seen_challans[challan]})",
                    "value": challan
                })
            else:
                seen_challans[challan] = excel_row

        # 2. Header row read as data
        if str(serial).upper() in ("SERIAL_NUMBER", "SR NO") or "CHALLAN" in str(challan).upper():
            blocking.append({"row": excel_row, "field": "ROW", "reason": "Header row read as data", "value": f"{serial}, {challan}"})

        # 3. Blank essential fields
        if not veh:
            blocking.append({"row": excel_row, "field": "VEHICLE_NUMBER", "reason": "Blank vehicle number", "value": ""})
        if not name:
            blocking.append({"row": excel_row, "field": "NAME", "reason": "Blank violator name", "value": ""})
        if not sec:
            blocking.append({"row": excel_row, "field": "SECTION", "reason": "Blank violation section", "value": ""})

        # 4. Unparseable date
        if not vdate or not re.match(r"^\d{2}[/-]\d{2}[/-]\d{4}$", str(vdate).strip()):
            blocking.append({"row": excel_row, "field": "VIOLATION_DATE", "reason": f"Unparseable violation date: '{vdate}'", "value": vdate})

        # 5. Numeric cells that would print wrongly (e.g. "500.0", "9.1e+09", scientific notation)
        for field_name in ("FINE_AMOUNT", "SERIAL_NUMBER", "MOBILE_NUMBER"):
            val_str = str(r.get(field_name, ""))
            if "e+" in val_str.lower() or "e-" in val_str.lower():
                blocking.append({"row": excel_row, "field": field_name, "reason": f"Scientific notation in numeric cell: '{val_str}'", "value": val_str})
            if field_name == "FINE_AMOUNT" and val_str.endswith(".0"):
                blocking.append({"row": excel_row, "field": field_name, "reason": f"Float decimal in fine amount: '{val_str}'", "value": val_str})

        # 6. Control characters that corrupt Word XML
        for field_name, val in r.items():
            val_str = str(val or "")
            if _CONTROL_CHAR_RE.search(val_str):
                blocking.append({"row": excel_row, "field": field_name, "reason": f"Illegal XML control character in {field_name}", "value": repr(val_str)})

        # --- WARNING CHECKS ---
        # 1. Mobile blank or not 10 digits
        clean_mobile = re.sub(r"[^0-9]", "", str(mobile or ""))
        if not clean_mobile:
            warnings.append({"row": excel_row, "field": "MOBILE_NUMBER", "reason": "Mobile number is blank", "value": ""})
        elif len(clean_mobile) != 10:
            warnings.append({"row": excel_row, "field": "MOBILE_NUMBER", "reason": f"Mobile number is not 10 digits ({len(clean_mobile)} digits)", "value": str(mobile)})

        # 2. Address longer than known fit (> 135 chars)
        if len(str(addr or "")) > 135:
            warnings.append({"row": excel_row, "field": "ADDRESS", "reason": f"Address exceeds standard length ({len(str(addr))} characters)", "value": str(addr)})

        # 3. Duplicate serial
        if serial:
            if serial in seen_serials:
                warnings.append({
                    "row": excel_row,
                    "field": "SERIAL_NUMBER",
                    "reason": f"Duplicate serial number '{serial}' (first seen on row {seen_serials[serial]})",
                    "value": serial
                })
            else:
                seen_serials[serial] = excel_row

        # 4. Unusual fine amount (outside common 100 - 25000 range or not a multiple of 100/500)
        try:
            fine_val = int(float(fine))
            if fine_val < 100 or fine_val > 25000:
                warnings.append({"row": excel_row, "field": "FINE_AMOUNT", "reason": f"Unusual fine amount: Rs. {fine_val}", "value": str(fine)})
        except (ValueError, TypeError):
            pass

    return blocking, warnings


def inspect_file(path: Path, preferred_sheet: Optional[str] = None) -> InspectionResult:
    """Performs full inspection of an uploaded file."""
    path = Path(path)
    file_type = check_magic_bytes(path)

    if file_type == "csv":
        sheet_name = None
        available_sheets = []
        start_row = 2
        records = nl.read_records(path)
        needs_sheet_selection = False
    else:
        sheet_name, start_row, available_sheets = detect_excel_data_block(path, preferred_sheet=preferred_sheet)
        needs_sheet_selection = len(available_sheets) > 1 and not preferred_sheet
        records = nl.read_records(path, sheet_name=sheet_name, start_row=start_row)

    blocking, warnings = run_preflight(records, start_excel_row=start_row)

    return InspectionResult(
        valid=len(blocking) == 0,
        file_type=file_type,
        sheet_name=sheet_name,
        available_sheets=available_sheets,
        start_row=start_row,
        total_records=len(records),
        records=records,
        blocking_errors=blocking,
        warnings=warnings,
        first_5=records[:5],
        last_5=records[-5:] if len(records) > 5 else records,
        needs_sheet_selection=needs_sheet_selection,
    )


def export_problems_xlsx(blocking: list[dict], warnings: list[dict], out_path: Path) -> Path:
    """Exports flagged preflight rows into an Excel workbook for supervisor review."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Flagged Rows"

    headers = ["Excel Row", "Severity", "Field Name", "Reason", "Current Value"]
    ws.append(headers)

    for item in blocking:
        ws.append([item["row"], "BLOCKING", item["field"], item["reason"], str(item.get("value", ""))])

    for item in warnings:
        ws.append([item["row"], "WARNING", item["field"], item["reason"], str(item.get("value", ""))])

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(out_path)
    return out_path
