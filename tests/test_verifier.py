import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import verifier
import notice_lib as nl

TEMPLATE_PATH = ROOT / "templates" / "Formate_with_Sign_SectionSplit.docx"
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe")


class TestVerifierGates(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp_dir = ROOT / "tmp" / "test_verifier"
        cls.tmp_dir.mkdir(parents=True, exist_ok=True)

        cls.rec1 = {
            "SERIAL_NUMBER": "2001",
            "E_CHALLAN_NUMBER": "CG5010992607182001",
            "VEHICLE_NUMBER": "CG07A2001",
            "VIOLATION_DATE": "18/07/2026",
            "NAME": "SANJAY SAHU",
            "ADDRESS": "HOUSE NO 10 SECTOR 1 BHILAI DURG CG 490006",
            "MOBILE_NUMBER": "9826120001",
            "FINE_AMOUNT": "500",
            "SECTION": "129/194(D)Driving without helmet",
        }
        cls.rec2 = {
            "SERIAL_NUMBER": "2002",
            "E_CHALLAN_NUMBER": "CG5010992607182002",
            "VEHICLE_NUMBER": "CG07A2002",
            "VIOLATION_DATE": "19/07/2026",
            "NAME": "VIKRAM SINGH",
            "ADDRESS": "HOUSE NO 20 SECTOR 2 BHILAI DURG CG 490006",
            "MOBILE_NUMBER": "9826120002",
            "FINE_AMOUNT": "1000",
            "SECTION": "190(2). For 2 wheeler without valid documents",
        }

    def _write_csv(self, filename: str, records: list[dict]) -> Path:
        p = self.tmp_dir / filename
        lines = ["SERIAL_NUMBER,E_CHALLAN_NUMBER,VEHICLE_NUMBER,VIOLATION_DATE,NAME,ADDRESS,MOBILE_NUMBER,FINE_AMOUNT,SECTION"]
        for r in records:
            lines.append(f'{r["SERIAL_NUMBER"]},{r["E_CHALLAN_NUMBER"]},{r["VEHICLE_NUMBER"]},{r["VIOLATION_DATE"]},'
                         f'"{r["NAME"]}","{r["ADDRESS"]}",{r["MOBILE_NUMBER"]},{r["FINE_AMOUNT"]},"{r["SECTION"]}"')
        p.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
        return p

    def test_g3_passes_on_correct_generation(self):
        csv_p = self._write_csv("ver_2.csv", [self.rec1, self.rec2])
        out_d = self.tmp_dir / "out_ver_2"
        run = nl.generate_all(csv_p, TEMPLATE_PATH, out_d, batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        self.assertEqual(len(run.ok_batches), 1)

        res = verifier.verify_all_gates(blocking_errors=[], run_batches=run.batches, all_records=[self.rec1, self.rec2])
        self.assertTrue(res.passed, f"Expected G3 to pass but got errors: {res.errors}")
        self.assertEqual(res.verified_records, 2)

    def test_g3_negative_test_swapped_records_detected(self):
        """Negative test: swap two records in the expected records list and prove G3 catches it!"""
        csv_p = self._write_csv("ver_swap.csv", [self.rec1, self.rec2])
        out_d = self.tmp_dir / "out_ver_swap"
        run = nl.generate_all(csv_p, TEMPLATE_PATH, out_d, batch_size=1000, soffice_path=SOFFICE_PATH, workers=1)
        self.assertEqual(len(run.ok_batches), 1)

        # Deliberately swap rec1 and rec2 when checking against the PDF
        swapped_records = [self.rec2, self.rec1]
        res = verifier.verify_all_gates(blocking_errors=[], run_batches=run.batches, all_records=swapped_records)

        self.assertFalse(res.passed, "G3 should have caught swapped records!")
        self.assertTrue(any("missing from correct front half" in e or "leaked into OPPOSITE" in e for e in res.errors),
                        f"Expected missing or leaked error in {res.errors}")


if __name__ == "__main__":
    unittest.main()
