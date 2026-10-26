#!/usr/bin/env python3
import argparse
import time
from pathlib import Path
import notice_lib as nl

def main():
    ap = argparse.ArgumentParser(description="Create one DOCX batch without LibreOffice conversion.")
    ap.add_argument("--input", required=True)
    ap.add_argument("--template", default="templates/Formate_with_Sign_SectionSplit.docx")
    ap.add_argument("--output", default="tmp/direct_lo_test.docx")
    ap.add_argument("--sheet", default="Bike")
    ap.add_argument("--start-row", type=int, default=8)
    ap.add_argument("--batch-size", type=int, default=1000)
    args = ap.parse_args()

    input_path = Path(args.input)
    template_path = Path(args.template)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    t0 = time.perf_counter()
    records = nl.read_records(input_path, sheet_name=args.sheet, start_row=args.start_row)
    records = records[:args.batch_size]
    t1 = time.perf_counter()

    parts = nl.load_template_parts(template_path)
    pairs = nl.make_pairs(records)
    t2 = time.perf_counter()

    nl.build_batch_docx(template_path, parts, pairs, output_path)
    t3 = time.perf_counter()

    print(f"Records read   : {len(records)}")
    print(f"Read time      : {t1-t0:.2f}s")
    print(f"Template/pairs : {t2-t1:.2f}s")
    print(f"DOCX build     : {t3-t2:.2f}s")
    print(f"Total          : {t3-t0:.2f}s")
    print(f"DOCX created   : {output_path}")
    print(f"DOCX size      : {output_path.stat().st_size/1e6:.2f} MB")

if __name__ == "__main__":
    main()
