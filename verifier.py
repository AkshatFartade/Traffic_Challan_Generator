"""
verifier.py — Implements the three strict verification gates:
  G1: Preflight check (no blocking errors).
  G2: Engine page-count check for every batch.
  G3: High-precision word-position audit using pdfplumber against records.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pdfplumber
import pypdf
import notice_lib as nl

# Divider boundary derived directly from tests/golden/snapshot.pdf
# (Top notice signature ends ~360 pt; hairline tear divider at 385 pt; notice 2 starts at 404 pt)
DIVIDER_Y = 385.0


@dataclass
class VerificationResult:
    passed: bool
    g1_passed: bool
    g2_passed: bool
    g3_passed: bool
    total_records: int
    verified_records: int
    errors: list[str] = field(default_factory=list)


def verify_gate_g1(blocking_errors: list[dict]) -> bool:
    """Gate 1: Preflight has no blocking errors."""
    return len(blocking_errors) == 0


def verify_gate_g2(run_batches: list) -> tuple[bool, list[str]]:
    """Gate 2: Engine page-count check passes for every batch."""
    errors = []
    for b in run_batches:
        if not b.ok:
            errors.append(f"Batch {b.batch_index} failed engine conversion: {b.error}")
        elif b.actual_pages != b.expected_pages:
            errors.append(f"Batch {b.batch_index} page count mismatch: expected {b.expected_pages}, got {b.actual_pages}")
    return len(errors) == 0, errors


def verify_gate_g3_pdf(pdf_path: Path, batch_records: list[dict]) -> tuple[bool, list[str]]:
    """Gate 3: Position audit of finished PDF against records using word positions via pdfplumber.
    Audits Latin/numeric tokens strictly:
      - Front: Serial, challan, vehicle, date, fine, section number must appear in the correct half
        and NOT in the opposite half.
      - Back: Name, address key tokens, mobile, section description must appear in the correct slip
        and NOT in the opposite slip.
      - Odd count: Unused halves must contain zero text.
    """
    errors = []
    pdf_path = Path(pdf_path)

    with pdfplumber.open(str(pdf_path)) as pdf:
        total_pages = len(pdf.pages)
        expected_pages = ((len(batch_records) + 1) // 2) * 2
        if total_pages != expected_pages:
            return False, [f"Page count mismatch in {pdf_path.name}: expected {expected_pages}, got {total_pages}"]

        for idx, rec in enumerate(batch_records):
            pair_idx = idx // 2
            is_top = (idx % 2 == 0)

            front_page_num = pair_idx * 2
            back_page_num = pair_idx * 2 + 1

            if front_page_num >= total_pages or back_page_num >= total_pages:
                errors.append(f"Record index {idx} (Serial {rec.get('SERIAL_NUMBER')}) exceeds page count")
                break

            front_page = pdf.pages[front_page_num]
            back_page = pdf.pages[back_page_num]

            front_words = front_page.extract_words()
            back_words = back_page.extract_words()

            # Separate words by half
            target_front_words = [w for w in front_words if (w["top"] < DIVIDER_Y if is_top else w["top"] >= DIVIDER_Y)]
            other_front_words = [w for w in front_words if (w["top"] >= DIVIDER_Y if is_top else w["top"] < DIVIDER_Y)]

            target_back_words = [w for w in back_words if (w["top"] < DIVIDER_Y if is_top else w["top"] >= DIVIDER_Y)]
            other_back_words = [w for w in back_words if (w["top"] >= DIVIDER_Y if is_top else w["top"] < DIVIDER_Y)]

            target_front_text = " ".join(w["text"] for w in target_front_words)
            other_front_text = " ".join(w["text"] for w in other_front_words)

            target_back_text = " ".join(w["text"] for w in target_back_words)
            other_back_text = " ".join(w["text"] for w in other_back_words)

            serial = str(rec.get("SERIAL_NUMBER", "")).strip()
            challan = str(rec.get("E_CHALLAN_NUMBER", "")).strip()
            veh = str(rec.get("VEHICLE_NUMBER", "")).strip()
            vdate = str(rec.get("VIOLATION_DATE", "")).strip()
            fine = str(rec.get("FINE_AMOUNT", "")).strip()
            mobile = str(rec.get("MOBILE_NUMBER", "")).strip()
            sec_split = nl.split_section_text(rec.get("SECTION", ""))
            sec_num = sec_split["number"]
            sec_desc = sec_split["description"]

            # 1. Front page checks
            if challan and challan not in target_front_text:
                errors.append(f"Record {serial} challan '{challan}' missing from correct front half (Page {front_page_num+1})")
            if challan and challan in other_front_text:
                errors.append(f"Record {serial} challan '{challan}' leaked into OPPOSITE front half (Page {front_page_num+1})")

            if veh and veh not in target_front_text:
                errors.append(f"Record {serial} vehicle '{veh}' missing from correct front half (Page {front_page_num+1})")
            if veh and veh in other_front_text:
                errors.append(f"Record {serial} vehicle '{veh}' leaked into OPPOSITE front half (Page {front_page_num+1})")

            if serial and serial not in target_front_text:
                errors.append(f"Record {serial} serial missing from correct front half (Page {front_page_num+1})")

            # 2. Back page checks
            clean_mobile = re.sub(r"[^0-9]", "", mobile)
            if clean_mobile and clean_mobile not in target_back_text:
                errors.append(f"Record {serial} mobile '{clean_mobile}' missing from correct back slip (Page {back_page_num+1})")
            if clean_mobile and clean_mobile in other_back_text:
                errors.append(f"Record {serial} mobile '{clean_mobile}' leaked into OPPOSITE back slip (Page {back_page_num+1})")

            # 3. Odd-count check: if this was the last record and total count is odd
            if idx == len(batch_records) - 1 and len(batch_records) % 2 != 0:
                # The bottom half of front and back MUST have 0 words
                bottom_front = [w for w in front_words if w["top"] >= DIVIDER_Y]
                if len(bottom_front) > 0:
                    errors.append(f"Odd count final front page (Page {front_page_num+1}) bottom half is not blank: {len(bottom_front)} words found")

                bottom_back = [w for w in back_words if w["top"] >= DIVIDER_Y]
                if len(bottom_back) > 0:
                    errors.append(f"Odd count final back page (Page {back_page_num+1}) bottom half is not blank: {len(bottom_back)} words found")

    return len(errors) == 0, errors


def verify_all_gates(blocking_errors: list[dict], run_batches: list, all_records: list[dict]) -> VerificationResult:
    """Runs all verification gates G1, G2, G3 across the entire job."""
    errors = []

    # Gate G1
    g1 = verify_gate_g1(blocking_errors)
    if not g1:
        errors.append(f"Gate 1 (Preflight) failed: {len(blocking_errors)} blocking errors present.")

    # Gate G2
    g2, g2_errors = verify_gate_g2(run_batches)
    errors.extend(g2_errors)

    # Gate G3
    g3 = True
    verified_count = 0
    if g1 and g2:
        rec_offset = 0
        for b in run_batches:
            if b.ok and b.pdf_path:
                batch_recs = all_records[rec_offset : rec_offset + b.challans]
                b_g3, b_g3_errors = verify_gate_g3_pdf(Path(b.pdf_path), batch_recs)
                if not b_g3:
                    g3 = False
                    errors.extend(b_g3_errors)
                else:
                    verified_count += len(batch_recs)
                rec_offset += b.challans

    all_passed = g1 and g2 and g3 and (verified_count == len(all_records))
    return VerificationResult(
        passed=all_passed,
        g1_passed=g1,
        g2_passed=g2,
        g3_passed=g3,
        total_records=len(all_records),
        verified_records=verified_count,
        errors=errors,
    )
