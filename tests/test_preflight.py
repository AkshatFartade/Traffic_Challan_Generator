import unittest
from pathlib import Path
import input_inspector as ii

class TestPreflightRules(unittest.TestCase):
    def test_pipe_separated_multi_sections_not_flagged(self):
        """Pipe-separated multi-section values such as
        '128/194(C)Triple Riding on Two Wheeler|129/194(D) Driving without helmet'
        are normal and must NOT be flagged as blocking or warning."""
        recs = [{
            "SERIAL_NUMBER": "101",
            "E_CHALLAN_NUMBER": "CG5010992607180101",
            "VEHICLE_NUMBER": "CG07A1101",
            "VIOLATION_DATE": "18/07/2026",
            "NAME": "RAMESH SAHU",
            "ADDRESS": "HOUSE 1 WARD 2 BHILAI DURG CG 490006",
            "MOBILE_NUMBER": "9826101101",
            "FINE_AMOUNT": "1000",
            "SECTION": "128/194(C)Triple Riding on Two Wheeler|129/194(D) Driving without helmet"
        }]
        blocking, warnings = ii.run_preflight(recs)
        self.assertEqual(len(blocking), 0, f"Unexpected blocking error: {blocking}")
        self.assertEqual(len(warnings), 0, f"Unexpected warning: {warnings}")

    def test_blocking_missing_challan(self):
        recs = [{
            "SERIAL_NUMBER": "101",
            "E_CHALLAN_NUMBER": "",
            "VEHICLE_NUMBER": "CG07A1101",
            "VIOLATION_DATE": "18/07/2026",
            "NAME": "RAMESH SAHU",
            "ADDRESS": "HOUSE 1 WARD 2 BHILAI DURG CG 490006",
            "MOBILE_NUMBER": "9826101101",
            "FINE_AMOUNT": "500",
            "SECTION": "129/194(D)Driving without helmet"
        }]
        blocking, _ = ii.run_preflight(recs)
        self.assertTrue(any("Missing challan number" in b["reason"] for b in blocking))

    def test_blocking_duplicate_challan(self):
        recs = [
            {
                "SERIAL_NUMBER": "101",
                "E_CHALLAN_NUMBER": "CG_DUP_12345",
                "VEHICLE_NUMBER": "CG07A1101",
                "VIOLATION_DATE": "18/07/2026",
                "NAME": "RAMESH SAHU",
                "ADDRESS": "HOUSE 1 WARD 2 BHILAI DURG CG 490006",
                "MOBILE_NUMBER": "9826101101",
                "FINE_AMOUNT": "500",
                "SECTION": "129/194(D)Driving without helmet"
            },
            {
                "SERIAL_NUMBER": "102",
                "E_CHALLAN_NUMBER": "CG_DUP_12345",
                "VEHICLE_NUMBER": "CG07A1102",
                "VIOLATION_DATE": "19/07/2026",
                "NAME": "SURESH SAHU",
                "ADDRESS": "HOUSE 2 WARD 2 BHILAI DURG CG 490006",
                "MOBILE_NUMBER": "9826101102",
                "FINE_AMOUNT": "500",
                "SECTION": "129/194(D)Driving without helmet"
            }
        ]
        blocking, _ = ii.run_preflight(recs)
        self.assertTrue(any("Duplicate challan number" in b["reason"] for b in blocking))

    def test_blocking_numeric_cells_float_and_scientific(self):
        recs = [
            {
                "SERIAL_NUMBER": "101",
                "E_CHALLAN_NUMBER": "CG5010992607180101",
                "VEHICLE_NUMBER": "CG07A1101",
                "VIOLATION_DATE": "18/07/2026",
                "NAME": "RAMESH SAHU",
                "ADDRESS": "HOUSE 1 WARD 2 BHILAI DURG CG 490006",
                "MOBILE_NUMBER": "9.1e+09",
                "FINE_AMOUNT": "500.0",
                "SECTION": "129/194(D)Driving without helmet"
            }
        ]
        blocking, _ = ii.run_preflight(recs)
        reasons = [b["reason"] for b in blocking]
        self.assertTrue(any("Scientific notation" in r for r in reasons))
        self.assertTrue(any("Float decimal in fine amount" in r for r in reasons))

    def test_blocking_control_characters(self):
        recs = [{
            "SERIAL_NUMBER": "101",
            "E_CHALLAN_NUMBER": "CG5010992607180101",
            "VEHICLE_NUMBER": "CG07A1101",
            "VIOLATION_DATE": "18/07/2026",
            "NAME": "RAMESH\x00SAHU",
            "ADDRESS": "HOUSE 1\x0bWARD 2",
            "MOBILE_NUMBER": "9826101101",
            "FINE_AMOUNT": "500",
            "SECTION": "129/194(D)Driving without helmet"
        }]
        blocking, _ = ii.run_preflight(recs)
        self.assertTrue(any("Illegal XML control character" in b["reason"] for b in blocking))

    def test_warnings_mobile_and_long_address(self):
        recs = [{
            "SERIAL_NUMBER": "101",
            "E_CHALLAN_NUMBER": "CG5010992607180101",
            "VEHICLE_NUMBER": "CG07A1101",
            "VIOLATION_DATE": "18/07/2026",
            "NAME": "RAMESH SAHU",
            "ADDRESS": "A" * 140,  # Longer than 135 chars
            "MOBILE_NUMBER": "982610",  # Only 6 digits
            "FINE_AMOUNT": "500",
            "SECTION": "129/194(D)Driving without helmet"
        }]
        _, warnings = ii.run_preflight(recs)
        w_reasons = [w["reason"] for w in warnings]
        self.assertTrue(any("not 10 digits" in r for r in w_reasons))
        self.assertTrue(any("Address exceeds standard length" in r for r in w_reasons))

    def test_xls_legacy_rejection(self):
        with self.assertRaises(ValueError) as ctx:
            ii.check_magic_bytes(Path("legacy_file.xls"))
        self.assertIn("Excel 97-2003 (.xls) format is not supported", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
