"""
test_comprehensive.py  –  Item 1 evidence tests
================================================
Covers every item from the completion brief that was not already in the
existing test files:

  Scale:         1000 (even), 1001 (odd crossing batch boundary), 999 (odd)
  Odd edge:      1 record, 3 records  (page math only – blanking already in
                 test_odd_count.py)
  Edge data:     131-char address, Hindi text, &/<>/'/" XML specials,
                 control character, pipe multi-section
  Preflight:     header-row-as-data blocked, duplicate-challan blocked
  Layout detect: Sheet1/row8, Sheet1/row2, Bike/row8
  G3 negative:   swapped records list AND physically page-swapped PDF
"""
from __future__ import annotations

import csv
import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import input_inspector as ii
import notice_lib as nl
import verifier

TEMPLATE_PATH = ROOT / "templates" / "Formate_with_Sign_SectionSplit.docx"
SOFFICE_PATH = os.environ.get(
    "SOFFICE_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_rec(n: int, *, address=None, section=None, name=None, challan=None) -> dict:
    return {
        "SERIAL_NUMBER": str(n),
        "E_CHALLAN_NUMBER": challan or f"CGTC{n:012d}",
        "VEHICLE_NUMBER": f"CG07TC{n:04d}",
        "VIOLATION_DATE": "18/07/2026",
        "NAME": name or f"TEST PERSON {n}",
        "ADDRESS": address or f"HOUSE {n} WARD 1 BHILAI DURG CG 490006",
        "MOBILE_NUMBER": str(9800000000 + (n % 9999999)),
        "FINE_AMOUNT": "500",
        "SECTION": section or "129/194(D)Driving without helmet",
    }


def _write_csv(path: Path, records: list) -> Path:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=nl.FIELD_KEYS)
        w.writeheader()
        w.writerows(records)
    return path


def _write_xlsx(path: Path, records: list, sheet_name="Bike", start_row=8) -> Path:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    # Header row one above data
    headers = ["Serial Number", "E-Challan Number", "Vehicle Number",
               "Violation Date", "Name", "Address", "Mobile Number",
               "Fine Amount", "Section"]
    for ci, h in enumerate(headers, 2):
        ws.cell(start_row - 1, ci, h)
    for ri, rec in enumerate(records):
        r = start_row + ri
        serial = rec["SERIAL_NUMBER"]
        ws.cell(r, 2, int(serial) if str(serial).isdigit() else serial)
        ws.cell(r, 3, rec["E_CHALLAN_NUMBER"])
        ws.cell(r, 4, rec["VEHICLE_NUMBER"])
        ws.cell(r, 5, rec["VIOLATION_DATE"])
        ws.cell(r, 6, rec["NAME"])
        ws.cell(r, 7, rec["ADDRESS"])
        ws.cell(r, 8, rec["MOBILE_NUMBER"])
        fine = rec["FINE_AMOUNT"]
        ws.cell(r, 9, int(fine) if str(fine).isdigit() else fine)
        ws.cell(r, 10, rec["SECTION"])
    wb.save(path)
    return path


# ===========================================================================
# 1.  SCALE RUNS  (1000, 1001, 999)
# ===========================================================================

class TestScaleRuns(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = ROOT / "tmp" / "test_scale"
        cls.tmp.mkdir(parents=True, exist_ok=True)

    def _run(self, n: int, workers: int = 1):
        recs = [_make_rec(i + 1) for i in range(n)]
        csv_p = _write_csv(self.tmp / f"scale_{n}.csv", recs)
        out_d = self.tmp / f"out_{n}_w{workers}"
        run = nl.generate_all(
            csv_p, TEMPLATE_PATH, out_d,
            batch_size=500, soffice_path=SOFFICE_PATH, workers=workers,
        )
        return run, recs

    def test_1000_even(self):
        run, recs = self._run(1000)
        total_challans = sum(b.challans for b in run.batches)
        total_pages = sum(b.actual_pages for b in run.ok_batches)
        self.assertEqual(len(run.failed_batches), 0)
        self.assertEqual(total_challans, 1000)
        self.assertEqual(total_pages, 1000)   # 1000 / 2 = 500 pairs * 2 pages
        res = verifier.verify_all_gates([], run.batches, recs)
        self.assertTrue(res.passed, f"G3 failed: {res.errors[:3]}")
        print(f"\n[EVIDENCE] 1000 records: {total_pages} pages, "
              f"{sum(b.seconds for b in run.ok_batches):.1f}s total, "
              f"verified={res.verified_records}")

    def test_999_odd(self):
        run, recs = self._run(999)
        total_pages = sum(b.actual_pages for b in run.ok_batches)
        self.assertEqual(len(run.failed_batches), 0)
        self.assertEqual(total_pages, 1000)   # ceil(999/2)*2 = 1000
        res = verifier.verify_all_gates([], run.batches, recs)
        self.assertTrue(res.passed, f"G3 failed: {res.errors[:3]}")
        print(f"\n[EVIDENCE] 999 records: {total_pages} pages, verified={res.verified_records}")

    def test_1001_odd_batch_boundary(self):
        """1001 records spans 3 batches of 500 (500+500+1); last batch is odd."""
        run, recs = self._run(1001)
        total_pages = sum(b.actual_pages for b in run.ok_batches)
        self.assertEqual(len(run.failed_batches), 0)
        self.assertEqual(total_pages, 1002)   # ceil(1001/2)*2 = 1002
        res = verifier.verify_all_gates([], run.batches, recs)
        self.assertTrue(res.passed, f"G3 failed: {res.errors[:3]}")
        print(f"\n[EVIDENCE] 1001 records: {total_pages} pages, "
              f"{len(run.ok_batches)} batches, verified={res.verified_records}")


# ===========================================================================
# 2.  ODD EDGE: 1 and 3 records
# ===========================================================================

class TestOddEdgeCounts(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = ROOT / "tmp" / "test_odd_edge"
        cls.tmp.mkdir(parents=True, exist_ok=True)

    def _pages(self, n: int) -> int:
        recs = [_make_rec(i + 1) for i in range(n)]
        csv_p = _write_csv(self.tmp / f"oe_{n}.csv", recs)
        out_d = self.tmp / f"out_oe_{n}"
        run = nl.generate_all(csv_p, TEMPLATE_PATH, out_d,
                              batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        return sum(b.actual_pages for b in run.ok_batches)

    def test_1_record_gives_2_pages(self):
        p = self._pages(1)
        self.assertEqual(p, 2)
        print(f"\n[EVIDENCE] 1 record -> {p} pages")

    def test_3_records_give_4_pages(self):
        p = self._pages(3)
        self.assertEqual(p, 4)
        print(f"\n[EVIDENCE] 3 records -> {p} pages")


# ===========================================================================
# 3.  EDGE-CASE DATA
# ===========================================================================

class TestEdgeCaseData(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = ROOT / "tmp" / "test_edge"
        cls.tmp.mkdir(parents=True, exist_ok=True)

    def _gen2(self, rec1, rec2, tag):
        csv_p = _write_csv(self.tmp / f"{tag}.csv", [rec1, rec2])
        out_d = self.tmp / f"out_{tag}"
        run = nl.generate_all(csv_p, TEMPLATE_PATH, out_d,
                              batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        self.assertEqual(len(run.failed_batches), 0,
                         f"{tag}: {[b.error for b in run.failed_batches]}")
        return run

    def test_131_char_address_does_not_block_and_generates(self):
        base = "HOUSE NO 12 AASTHA COLONY NEAR RAILWAY STATION MAIN ROAD WARD 42 BHILAI NAGAR DURG CHHATTISGARH INDIA 490006"
        addr = (base + " EXTRA PADDING TEXT HERE")[:131]
        self.assertEqual(len(addr), 131)
        blocking, warnings = ii.run_preflight([_make_rec(1, address=addr)])
        self.assertEqual(len(blocking), 0, f"131-char address must not block: {blocking}")
        run = self._gen2(_make_rec(1, address=addr), _make_rec(2), "addr131")
        sz = Path(run.ok_batches[0].pdf_path).stat().st_size
        self.assertGreater(sz, 0)
        print(f"\n[EVIDENCE] 131-char address: len={len(addr)}, blocking=0, pdf_size={sz:,} bytes")

    def test_hindi_text_generates_without_crash(self):
        rec = _make_rec(10,
                        name="राकेश कुमार वर्मा",
                        address="मकान नं 42 शांति नगर वार्ड 7 भिलाई दुर्ग छत्तीसगढ़ 490006")
        run = self._gen2(rec, _make_rec(11), "hindi")
        sz = Path(run.ok_batches[0].pdf_path).stat().st_size
        print(f"\n[EVIDENCE] Hindi text: pdf_size={sz:,} bytes, pages={run.ok_batches[0].actual_pages}")

    def test_xml_special_chars_escaped(self):
        """& < > ' \" in challan fields must be XML-escaped, not crash."""
        rec = _make_rec(20,
                        name='A & B < C > D "E" \'F\'',
                        address='WARD 1 & 2 ROAD < BYPASS > BHILAI')
        run = self._gen2(rec, _make_rec(21), "xmlspecial")
        print(f"\n[EVIDENCE] XML specials (& < > ' \"): PDF OK, pages={run.ok_batches[0].actual_pages}")

    def test_pipe_multi_section_not_flagged(self):
        multi = "128/194(C)Triple Riding|129/194(D) Driving without helmet"
        blocking, warnings = ii.run_preflight([_make_rec(30, section=multi)])
        self.assertEqual(len(blocking), 0, f"Pipe section should not block: {blocking}")
        run = self._gen2(_make_rec(30, section=multi), _make_rec(31), "pipe")
        print(f"\n[EVIDENCE] Pipe multi-section: blocking=0, PDF OK")


# ===========================================================================
# 4.  PREFLIGHT EXTENDED
# ===========================================================================

class TestPreflightExtended(unittest.TestCase):

    def test_header_row_blocked(self):
        rec = {
            "SERIAL_NUMBER": "SERIAL_NUMBER",
            "E_CHALLAN_NUMBER": "E_CHALLAN_NUMBER",
            "VEHICLE_NUMBER": "CG07A0001", "VIOLATION_DATE": "18/07/2026",
            "NAME": "TEST", "ADDRESS": "ADDR", "MOBILE_NUMBER": "9800000001",
            "FINE_AMOUNT": "500", "SECTION": "129/194(D)test",
        }
        blocking, _ = ii.run_preflight([rec])
        reasons = [b["reason"] for b in blocking]
        self.assertTrue(any("Header row" in r for r in reasons),
                        f"Header row must be blocked. Got: {reasons}")
        print(f"\n[EVIDENCE] Header-row-as-data blocked: {[r for r in reasons if 'Header' in r]}")

    def test_duplicate_challan_blocking(self):
        r1 = _make_rec(101, challan="CG_DUPTEST_001")
        r2 = _make_rec(102, challan="CG_DUPTEST_001")
        blocking, _ = ii.run_preflight([r1, r2])
        reasons = [b["reason"] for b in blocking]
        self.assertTrue(any("Duplicate challan" in r for r in reasons),
                        f"Duplicate challan must be blocked. Got: {reasons}")
        print(f"\n[EVIDENCE] Duplicate challan blocking: {[r for r in reasons if 'Duplicate' in r]}")

    def test_control_char_blocking(self):
        rec = _make_rec(200)
        rec["NAME"] = "RAMESH\x00SAHU"
        blocking, _ = ii.run_preflight([rec])
        self.assertTrue(any("Illegal XML control character" in b["reason"] for b in blocking))
        print(f"\n[EVIDENCE] Control char \\x00 in NAME: blocked")


# ===========================================================================
# 5.  LAYOUT DETECTION
# ===========================================================================

class TestLayoutDetection(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = ROOT / "tmp" / "test_layouts"
        cls.tmp.mkdir(parents=True, exist_ok=True)
        cls.recs = [_make_rec(i + 1) for i in range(3)]

    def test_layout_sheet1_row8(self):
        path = self.tmp / "layout_s1r8.xlsx"
        _write_xlsx(path, self.recs, sheet_name="Sheet1", start_row=8)
        sheet, row, _ = ii.detect_excel_data_block(path)
        self.assertEqual(sheet, "Sheet1")
        self.assertEqual(row, 8)
        print(f"\n[EVIDENCE] Layout Sheet1/row8: detected sheet='{sheet}', start_row={row}")

    def test_layout_sheet1_row2(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        for ri, rec in enumerate(self.recs):
            r = 2 + ri
            ws.cell(r, 2, int(rec["SERIAL_NUMBER"]))
            ws.cell(r, 3, rec["E_CHALLAN_NUMBER"])
            for ci, k in enumerate(["VEHICLE_NUMBER","VIOLATION_DATE","NAME",
                                     "ADDRESS","MOBILE_NUMBER","FINE_AMOUNT","SECTION"], 4):
                ws.cell(r, ci, rec[k])
        path = self.tmp / "layout_s1r2.xlsx"
        wb.save(path)
        sheet, row, _ = ii.detect_excel_data_block(path)
        self.assertEqual(sheet, "Sheet1")
        self.assertEqual(row, 2)
        print(f"\n[EVIDENCE] Layout Sheet1/row2: detected sheet='{sheet}', start_row={row}")

    def test_layout_bike_row8(self):
        path = self.tmp / "layout_bike_r8.xlsx"
        _write_xlsx(path, self.recs, sheet_name="Bike", start_row=8)
        sheet, row, _ = ii.detect_excel_data_block(path)
        self.assertEqual(sheet, "Bike")
        self.assertEqual(row, 8)
        print(f"\n[EVIDENCE] Layout Bike/row8: detected sheet='{sheet}', start_row={row}")


# ===========================================================================
# 6.  G3 NEGATIVE TESTS
# ===========================================================================

class TestG3Negative(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.tmp = ROOT / "tmp" / "test_g3neg"
        cls.tmp.mkdir(parents=True, exist_ok=True)

        cls.rec1 = _make_rec(4001)
        cls.rec2 = _make_rec(4002)
        csv_p = _write_csv(cls.tmp / "g3neg.csv", [cls.rec1, cls.rec2])
        out_d = cls.tmp / "out_g3neg"
        cls.gen_run = nl.generate_all(csv_p, TEMPLATE_PATH, out_d,
                                      batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        assert len(cls.gen_run.ok_batches) == 1, "setUpClass generation failed"
        cls.good_pdf = Path(cls.gen_run.ok_batches[0].pdf_path)

    def test_g3_passes_on_correct_pdf(self):
        res = verifier.verify_all_gates([], self.gen_run.batches, [self.rec1, self.rec2])
        self.assertTrue(res.passed, f"Good PDF should pass G3: {res.errors}")
        print(f"\n[EVIDENCE] G3 on correct PDF: PASSED, verified={res.verified_records}")

    def test_g3_fails_on_swapped_records_list(self):
        """Pass records in opposite order -- G3 must fail."""
        res = verifier.verify_all_gates([], self.gen_run.batches, [self.rec2, self.rec1])
        self.assertFalse(res.passed)
        self.assertGreater(len(res.errors), 0)
        print(f"\n[EVIDENCE] G3 swapped-records-list: FAILED (errors={len(res.errors)}): "
              f"{res.errors[:2]}")

    def test_g3_fails_on_physically_swapped_pdf_pages(self):
        """Create a tampered PDF: repeat the back page (p1) as page 0,
        so the front notice data is gone. G3 must detect missing challan."""
        import pypdf
        tampered = self.tmp / "g3_tampered.pdf"
        reader = pypdf.PdfReader(str(self.good_pdf))
        # 2-record run -> 2 pages: [0=front, 1=back]
        # Tamper: back page appears as page 0 -- front notice data missing
        writer = pypdf.PdfWriter()
        writer.add_page(reader.pages[1])  # back page instead of front
        writer.add_page(reader.pages[1])  # back page for back too
        with open(tampered, "wb") as fh:
            writer.write(fh)

        fake_batch = nl.BatchResult(
            batch_index=1, challans=2, pairs=1,
            expected_pages=2, actual_pages=2,
            pdf_path=str(tampered), ok=True,
        )
        res = verifier.verify_all_gates([], [fake_batch], [self.rec1, self.rec2])
        self.assertFalse(res.passed, "G3 must FAIL on physically tampered PDF")
        print(f"\n[EVIDENCE] G3 physically-tampered PDF: FAILED (errors={len(res.errors)}): "
              f"{res.errors[:2]}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
