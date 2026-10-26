import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pdfplumber
import pypdf
import notice_lib as nl

TEMPLATE_PATH = ROOT / "templates" / "Formate_with_Sign_SectionSplit.docx"
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe")


class TestOddCountEngine(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp_dir = ROOT / "tmp" / "test_odd_runs"
        cls.tmp_dir.mkdir(parents=True, exist_ok=True)

        cls.rec1 = {
            "SERIAL_NUMBER": "1001",
            "E_CHALLAN_NUMBER": "CG5010992607181301",
            "VEHICLE_NUMBER": "CG07A1001",
            "VIOLATION_DATE": "18/07/2026",
            "NAME": "ASHOK KUMAR",
            "ADDRESS": "HOUSE 101 WARD 5 CIVIL LINES RAIPUR CG 492001",
            "MOBILE_NUMBER": "9826100001",
            "FINE_AMOUNT": "500",
            "SECTION": "129/194(D)Driving without helmet",
        }
        cls.rec2 = {
            "SERIAL_NUMBER": "1002",
            "E_CHALLAN_NUMBER": "CG5010992607181302",
            "VEHICLE_NUMBER": "CG07A1002",
            "VIOLATION_DATE": "19/07/2026",
            "NAME": "BHUPESH PATEL",
            "ADDRESS": "QR NO 12-B SECTOR 4 BHILAI DURG CG 490006",
            "MOBILE_NUMBER": "9826100002",
            "FINE_AMOUNT": "1000",
            "SECTION": "190(2). For 2 wheeler without valid documents",
        }
        cls.rec3 = {
            "SERIAL_NUMBER": "1003",
            "E_CHALLAN_NUMBER": "CG5010992607181303",
            "VEHICLE_NUMBER": "CG07A1003",
            "VIOLATION_DATE": "20/07/2026",
            "NAME": "CHETAN VERMA",
            "ADDRESS": "PLOT 88 TALAPARA MAIN ROAD BILASPUR CG 495001",
            "MOBILE_NUMBER": "9826100003",
            "FINE_AMOUNT": "1500",
            "SECTION": "184/177 Dangerous driving / unsafe driving",
        }

    def _write_csv(self, filename: str, records: list[dict]) -> Path:
        p = self.tmp_dir / filename
        lines = ["SERIAL_NUMBER,E_CHALLAN_NUMBER,VEHICLE_NUMBER,VIOLATION_DATE,NAME,ADDRESS,MOBILE_NUMBER,FINE_AMOUNT,SECTION"]
        for r in records:
            lines.append(f'{r["SERIAL_NUMBER"]},{r["E_CHALLAN_NUMBER"]},{r["VEHICLE_NUMBER"]},{r["VIOLATION_DATE"]},'
                         f'"{r["NAME"]}","{r["ADDRESS"]}",{r["MOBILE_NUMBER"]},{r["FINE_AMOUNT"]},"{r["SECTION"]}"')
        p.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
        return p

    def test_single_record_odd(self):
        """A single record run produces 2 pages with top notice filled and bottom half completely blank."""
        csv_path = self._write_csv("odd_1.csv", [self.rec1])
        out_dir = self.tmp_dir / "out_odd_1"
        run = nl.generate_all(csv_path, TEMPLATE_PATH, out_dir, batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        self.assertEqual(len(run.ok_batches), 1)
        pdf_path = Path(run.ok_batches[0].pdf_path)

        # 1. Page count is 2
        self.assertEqual(nl.count_pdf_pages(pdf_path), 2)

        # 2. Bottom half on front (page 1) and back (page 2) contains NO text
        with pdfplumber.open(pdf_path) as pdf:
            p1 = pdf.pages[0]
            p2 = pdf.pages[1]
            height = p1.height
            mid_y = height / 2.0  # ~420.9 pt

            # Top half should contain words
            top_words_p1 = [w for w in p1.extract_words() if w["top"] < mid_y]
            self.assertTrue(len(top_words_p1) > 10, "Top notice words not found on page 1")

            # Bottom half of page 1 MUST have no text at all
            bottom_words_p1 = [w for w in p1.extract_words() if w["top"] >= mid_y]
            self.assertEqual(len(bottom_words_p1), 0, f"Expected blank bottom half on page 1, found: {bottom_words_p1}")

            # Top half of page 2 should contain back slip words
            top_words_p2 = [w for w in p2.extract_words() if w["top"] < mid_y]
            self.assertTrue(len(top_words_p2) > 5, "Back slip words not found on page 2")

            # Bottom half of page 2 MUST have no text at all
            bottom_words_p2 = [w for w in p2.extract_words() if w["top"] >= mid_y]
            self.assertEqual(len(bottom_words_p2), 0, f"Expected blank bottom half on page 2, found: {bottom_words_p2}")

    def test_odd_top_notice_identical_to_even_top_notice(self):
        """The top notice of an odd run's last sheet has identical text and coordinates (within 1.0 pt)
        to the same record rendered as the top notice of an even run."""
        # 2-record run: rec3 + rec2
        csv_even = self._write_csv("even_2_rec3_first.csv", [self.rec3, self.rec2])
        out_even = self.tmp_dir / "out_even_2"
        run_even = nl.generate_all(csv_even, TEMPLATE_PATH, out_even, batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        pdf_even = Path(run_even.ok_batches[0].pdf_path)

        # 3-record run: rec1 + rec2 + rec3 (rec3 is the 3rd, odd record)
        csv_odd = self._write_csv("odd_3_rec3_last.csv", [self.rec1, self.rec2, self.rec3])
        out_odd = self.tmp_dir / "out_odd_3"
        run_odd = nl.generate_all(csv_odd, TEMPLATE_PATH, out_odd, batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        pdf_odd = Path(run_odd.ok_batches[0].pdf_path)

        self.assertEqual(nl.count_pdf_pages(pdf_even), 2)
        self.assertEqual(nl.count_pdf_pages(pdf_odd), 4)

        with pdfplumber.open(pdf_even) as p_even, pdfplumber.open(pdf_odd) as p_odd:
            # rec3 is on Page 1 (top half) of even run
            # rec3 is on Page 3 (top half) of odd run
            page_even_front = p_even.pages[0]
            page_odd_front = p_odd.pages[2]

            # The top notice signature ends before y=370; the divider line is at y~385
            # Notice 2 in an even run begins at y~404 pt.
            DIVIDER_Y = 385.0
            words_even_top = [w for w in page_even_front.extract_words() if w["top"] < DIVIDER_Y]
            words_odd_top = [w for w in page_odd_front.extract_words() if w["top"] < DIVIDER_Y]

            self.assertEqual(len(words_even_top), len(words_odd_top),
                             f"Word count mismatch in top notice: {len(words_even_top)} vs {len(words_odd_top)}")

            for w_idx, (we, wo) in enumerate(zip(words_even_top, words_odd_top)):
                self.assertEqual(we["text"], wo["text"], f"Word {w_idx} text mismatch: {we['text']} vs {wo['text']}")
                for coord in ("x0", "top", "x1", "bottom"):
                    diff = abs(we[coord] - wo[coord])
                    self.assertLessEqual(
                        diff, 1.0,
                        f"Word '{we['text']}' {coord} differs by {diff:.2f} pt between odd and even run"
                    )

            # Check bottom half of odd run Page 3 and Page 4 has zero text
            bottom_odd_front = [w for w in page_odd_front.extract_words() if w["top"] >= DIVIDER_Y]
            self.assertEqual(len(bottom_odd_front), 0, f"Odd last front page bottom half is not blank: {bottom_odd_front}")

            page_odd_back = p_odd.pages[3]
            # Back slip is in top half; bottom half below DIVIDER_Y has zero text
            bottom_odd_back = [w for w in page_odd_back.extract_words() if w["top"] >= DIVIDER_Y]
            self.assertEqual(len(bottom_odd_back), 0, f"Odd last back page bottom half is not blank: {bottom_odd_back}")

    def test_odd_count_multiprocessing_workers_2(self):
        """Odd count runs produce identical outputs and page counts whether workers=1 or workers=2."""
        csv_path = self._write_csv("odd_3_workers.csv", [self.rec1, self.rec2, self.rec3])

        out_w1 = self.tmp_dir / "out_workers_1"
        run_w1 = nl.generate_all(csv_path, TEMPLATE_PATH, out_w1, batch_size=2, soffice_path=SOFFICE_PATH, workers=1)
        self.assertEqual(len(run_w1.ok_batches), 2)  # 2 batches: batch 1 has 2 recs, batch 2 has 1 rec

        out_w2 = self.tmp_dir / "out_workers_2"
        run_w2 = nl.generate_all(csv_path, TEMPLATE_PATH, out_w2, batch_size=2, soffice_path=SOFFICE_PATH, workers=2)
        self.assertEqual(len(run_w2.ok_batches), 2)

        # Batch 2 is the odd batch in both runs
        pdf_w1 = Path(run_w1.ok_batches[1].pdf_path)
        pdf_w2 = Path(run_w2.ok_batches[1].pdf_path)

        self.assertEqual(nl.count_pdf_pages(pdf_w1), 2)
        self.assertEqual(nl.count_pdf_pages(pdf_w2), 2)

        with pdfplumber.open(pdf_w1) as p1, pdfplumber.open(pdf_w2) as p2:
            words1 = p1.pages[0].extract_words()
            words2 = p2.pages[0].extract_words()
            self.assertEqual(len(words1), len(words2))
            for w1, w2 in zip(words1, words2):
                self.assertEqual(w1["text"], w2["text"])
                for coord in ("x0", "top", "x1", "bottom"):
                    self.assertLessEqual(abs(w1[coord] - w2[coord]), 1.0)


if __name__ == "__main__":
    unittest.main()
