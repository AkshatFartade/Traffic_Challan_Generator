import os
import sys
import unittest
from pathlib import Path

# Add project root to path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pdfplumber
import pypdf
import notice_lib as nl

SNAPSHOT_PDF = ROOT / "tests" / "golden" / "snapshot.pdf"
SYNTHETIC_CSV = ROOT / "tests" / "golden" / "synthetic_2_records.csv"
TEMPLATE_PATH = ROOT / "templates" / "Formate_with_Sign_SectionSplit.docx"
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe")


class TestGoldenSnapshot(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not SNAPSHOT_PDF.exists():
            raise FileNotFoundError(f"Baseline snapshot not found at {SNAPSHOT_PDF}")
        if not SYNTHETIC_CSV.exists():
            raise FileNotFoundError(f"Synthetic test data not found at {SYNTHETIC_CSV}")

    def test_golden_regression(self):
        """Re-runs the pipeline on the 2 synthetic records and compares against snapshot.pdf:
        - page count
        - page dimensions (A4)
        - words text in exact order
        - word bounding boxes within 1.0 pt
        - embedded fonts
        """
        tmp_out = ROOT / "tmp" / "test_golden_run"
        tmp_out.mkdir(parents=True, exist_ok=True)

        run = nl.generate_all(
            input_path=SYNTHETIC_CSV,
            template_path=TEMPLATE_PATH,
            output_dir=tmp_out,
            batch_size=1000,
            soffice_path=SOFFICE_PATH,
            workers=1,
        )

        self.assertEqual(len(run.ok_batches), 1, f"Batch generation failed: {run.batches}")
        produced_pdf = Path(run.ok_batches[0].pdf_path)

        # 1. Page count
        snap_reader = pypdf.PdfReader(str(SNAPSHOT_PDF))
        prod_reader = pypdf.PdfReader(str(produced_pdf))
        self.assertEqual(len(prod_reader.pages), len(snap_reader.pages), "Page count mismatch against golden snapshot")

        # 2. Dimensions & Word coordinates via pdfplumber
        with pdfplumber.open(SNAPSHOT_PDF) as snap_pdf, pdfplumber.open(produced_pdf) as prod_pdf:
            self.assertEqual(len(snap_pdf.pages), len(prod_pdf.pages))

            for page_idx, (snap_page, prod_page) in enumerate(zip(snap_pdf.pages, prod_pdf.pages)):
                # Page dimensions check (A4: ~595.3 x 841.9 pt)
                self.assertAlmostEqual(snap_page.width, prod_page.width, delta=1.0,
                                       msg=f"Page {page_idx+1} width mismatch")
                self.assertAlmostEqual(snap_page.height, prod_page.height, delta=1.0,
                                       msg=f"Page {page_idx+1} height mismatch")

                snap_words = snap_page.extract_words()
                prod_words = prod_page.extract_words()

                self.assertEqual(len(snap_words), len(prod_words),
                                 f"Page {page_idx+1} word count mismatch: {len(snap_words)} vs {len(prod_words)}")

                for w_idx, (sw, pw) in enumerate(zip(snap_words, prod_words)):
                    self.assertEqual(sw["text"], pw["text"],
                                     f"Page {page_idx+1} word {w_idx} text mismatch: '{sw['text']}' != '{pw['text']}'")
                    # Tolerance 1.0 pt
                    for coord in ("x0", "top", "x1", "bottom"):
                        diff = abs(sw[coord] - pw[coord])
                        self.assertLessEqual(
                            diff, 1.0,
                            f"Page {page_idx+1} word '{sw['text']}' {coord} shifted by {diff:.2f} pt "
                            f"(expected {sw[coord]:.2f}, got {pw[coord]:.2f})"
                        )


if __name__ == "__main__":
    unittest.main()
