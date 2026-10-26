#!/usr/bin/env python3
"""
generate_notices.py — CLI for the Python Traffic Notice generator.

Examples
--------
  python generate_notices.py --input records.csv --template templates/Formate_with_Sign_SectionSplit.docx --output out/
  python generate_notices.py --input daily.xlsx --sheet Bike --start-row 8 --batch-size 1000 --zip --output out/
"""
import argparse
import json
import logging
import sys
from pathlib import Path

import notice_lib as nl


def main():
    ap = argparse.ArgumentParser(description="Generate traffic notice PDFs from Excel/CSV challan data.")
    ap.add_argument("--input", required=True, help="Path to .csv or .xlsx challan data")
    ap.add_argument("--template", default="templates/Formate_with_Sign_SectionSplit.docx")
    ap.add_argument("--output", default="output", help="Directory to write batch PDFs into")
    ap.add_argument("--batch-size", type=int, default=1000, help="Challans per output PDF")
    ap.add_argument("--sheet", default="Bike", help="Excel sheet name (Excel input only)")
    ap.add_argument("--start-row", type=int, default=8, help="Excel start row (Excel input only)")
    ap.add_argument("--zip", action="store_true", help="Also zip all batch PDFs into one file")
    ap.add_argument("--soffice", default="soffice", help="Path to the soffice/LibreOffice executable")
    ap.add_argument("--workers", type=int, default=1, help="Convert this many batches in parallel (use your CPU core count)")
    ap.add_argument("--summary-json", default=None, help="Optional path to write a machine-readable run summary")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    def progress(batch_result, run):
        status = "OK" if batch_result.ok else "FAILED"
        print(
            f"  batch {batch_result.batch_index:>4} | {batch_result.challans:>5} challans | "
            f"{batch_result.seconds:>6.1f}s | {status}"
            + (f" | {batch_result.error}" if batch_result.error else "")
        )

    print(f"Reading {args.input} ...")
    run = nl.generate_all(
        input_path=args.input,
        template_path=args.template,
        output_dir=args.output,
        batch_size=args.batch_size,
        sheet_name=args.sheet,
        start_row=args.start_row,
        make_zip=args.zip,
        soffice_path=args.soffice,
        progress_cb=progress,
        workers=args.workers,
    )

    print()
    print(f"Total challans   : {run.total_challans}")
    print(f"Batches OK       : {len(run.ok_batches)}")
    print(f"Batches FAILED   : {len(run.failed_batches)}")
    print(f"Total time       : {run.total_seconds:.1f}s")
    if run.total_challans:
        print(f"Throughput       : {run.total_challans / run.total_seconds:.1f} challans/sec")
    if run.zip_path:
        print(f"Zip file         : {run.zip_path}")

    if args.summary_json:
        Path(args.summary_json).write_text(json.dumps(nl.run_result_to_dict(run), indent=2, ensure_ascii=False))

    if run.failed_batches:
        print("\nSome batches FAILED — see messages above. Nothing else was skipped silently.")
        sys.exit(1)


if __name__ == "__main__":
    main()
