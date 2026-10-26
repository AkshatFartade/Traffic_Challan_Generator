"""
job_runner.py — Standalone subprocess runner for generating traffic notices.
Isolates LibreOffice & ProcessPoolExecutor from the Flask web server.
Writes real-time status.json, enforces verification gates, and generates ZIP/manifest.
"""
from __future__ import annotations

import argparse
import datetime
import json
import logging
import os
import signal
import sys
import time
import zipfile
from pathlib import Path

# Add project root to path
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import config
import notice_lib as nl
import verifier

logger = logging.getLogger("job_runner")


def update_status(job_dir: Path, data: dict):
    status_file = job_dir / "status.json"
    temp_file = job_dir / "status.json.tmp"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    temp_file.replace(status_file)


def run_job(job_dir: Path):
    job_dir = Path(job_dir)
    spec_file = job_dir / "job_spec.json"
    if not spec_file.exists():
        raise FileNotFoundError(f"Missing job specification file: {spec_file}")

    with open(spec_file, "r", encoding="utf-8") as f:
        spec = json.load(f)

    input_path = Path(spec["input_path"])
    template_path = Path(spec.get("template_path", str(config.TEMPLATE_PATH)))
    batch_size = int(spec.get("batch_size", config.DEFAULT_BATCH_SIZE))
    workers = min(int(spec.get("workers", config.DEFAULT_WORKERS)), config.MAX_WORKERS)
    sheet_name = spec.get("sheet_name", "Bike")
    start_row = int(spec.get("start_row", 8))
    soffice_path = spec.get("soffice_path", config.SOFFICE_PATH)
    blocking_errors = spec.get("blocking_errors", [])

    output_dir = job_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    tmp_dir = job_dir / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    t_start = time.perf_counter()

    # Read records
    records = nl.read_records(input_path, sheet_name=sheet_name, start_row=start_row)
    total_challans = len(records)
    total_batches = (total_challans + batch_size - 1) // batch_size

    # Check if single batch
    is_single_batch = (total_batches == 1)

    status_data = {
        "job_id": job_dir.name,
        "state": "generating",
        "stage": "generating",
        "total_challans": total_challans,
        "batch_size": batch_size,
        "total_batches": total_batches,
        "completed_batches": 0,
        "batches": [],
        "elapsed_seconds": 0.0,
        "measured_speed": 0.0,
        "estimated_remaining_seconds": None,
        "message": "Converting to PDF. About 2 to 3 minutes for 1,000 notices on this computer." if is_single_batch else "Starting batch generation...",
        "verified": False,
        "verification_result": None,
        "download_ready": False,
        "download_url": None,
        "zip_name": None,
        "error": None,
    }
    update_status(job_dir, status_data)

    def progress_callback(batch_res: nl.BatchResult, run_res: nl.RunResult):
        elapsed = time.perf_counter() - t_start
        done = len(run_res.batches)
        challans_done = sum(b.challans for b in run_res.batches)
        speed = (challans_done / elapsed) if elapsed > 0 else 0.0

        rem_challans = total_challans - challans_done
        rem_seconds = (rem_challans / speed) if speed > 0 else None

        status_data["completed_batches"] = done
        status_data["elapsed_seconds"] = round(elapsed, 1)
        status_data["measured_speed"] = round(speed, 1)
        status_data["estimated_remaining_seconds"] = round(rem_seconds, 1) if rem_seconds is not None else None
        status_data["message"] = f"Generating: completed batch {done} of {total_batches} ({round(speed, 1)} notices/sec)"
        status_data["batches"].append({
            "batch_index": batch_res.batch_index,
            "challans": batch_res.challans,
            "expected_pages": batch_res.expected_pages,
            "actual_pages": batch_res.actual_pages,
            "ok": batch_res.ok,
            "seconds": round(batch_res.seconds, 1),
            "error": batch_res.error,
            "pdf_name": Path(batch_res.pdf_path).name if batch_res.pdf_path else None,
        })
        update_status(job_dir, status_data)

    try:
        # Run generation via notice_lib.generate_all
        run_res = nl.generate_all(
            input_path=input_path,
            template_path=template_path,
            output_dir=output_dir,
            batch_size=batch_size,
            sheet_name=sheet_name,
            start_row=start_row,
            make_zip=False,  # We assemble the zip and manifest ourselves after verification gates
            soffice_path=soffice_path,
            tmp_dir=tmp_dir,
            progress_cb=progress_callback,
            workers=workers,
        )

        # STAGE: VERIFYING
        status_data["state"] = "verifying"
        status_data["stage"] = "verifying"
        status_data["message"] = "Checking that every notice is in the right place (Gate G1, G2, G3)..."
        update_status(job_dir, status_data)

        ver_res = verifier.verify_all_gates(
            blocking_errors=blocking_errors,
            run_batches=run_res.batches,
            all_records=records,
        )

        status_data["verification_result"] = {
            "passed": ver_res.passed,
            "g1_passed": ver_res.g1_passed,
            "g2_passed": ver_res.g2_passed,
            "g3_passed": ver_res.g3_passed,
            "total_records": ver_res.total_records,
            "verified_records": ver_res.verified_records,
            "errors": ver_res.errors,
        }

        if not ver_res.passed:
            status_data["state"] = "verification_failed"
            status_data["stage"] = "verification_failed"
            status_data["message"] = f"Verification failed: {len(ver_res.errors)} problem(s) detected. Download disabled."
            status_data["download_ready"] = False
            update_status(job_dir, status_data)
            return

        # ALL GATES PASSED! Prepare outputs
        today_str = datetime.date.today().strftime("%d-%m-%Y")

        if is_single_batch and run_res.ok_batches:
            # Single batch PDF naming: Traffic_Notices_<DD-MM-YYYY>_<N>_challans.pdf
            pdf_src = Path(run_res.ok_batches[0].pdf_path)
            single_name = f"Traffic_Notices_{today_str}_{total_challans}_challans.pdf"
            single_dst = output_dir / single_name
            if pdf_src != single_dst:
                pdf_src.rename(single_dst)
                run_res.ok_batches[0].pdf_path = str(single_dst)

            status_data["download_name"] = single_name
            status_data["download_type"] = "pdf"
        else:
            # Multi-batch: Traffic_Notices_<DD-MM-YYYY>_<N>_challans.zip with manifest.txt
            zip_name = f"Traffic_Notices_{today_str}_{total_challans}_challans.zip"
            zip_path = output_dir / zip_name
            manifest_lines = [
                f"Traffic Notices Run Manifest",
                f"Date: {today_str}",
                f"Total Challans: {total_challans}",
                f"Total Batches: {len(run_res.ok_batches)}",
                f"Verification Gate G1: PASSED",
                f"Verification Gate G2: PASSED",
                f"Verification Gate G3: PASSED",
                f"",
                f"{'File':<55} | {'Serials':<15} | {'Count':<8} | {'Status'}",
                "-" * 95,
            ]
            rec_offset = 0
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for b in run_res.ok_batches:
                    p = Path(b.pdf_path)
                    batch_recs = records[rec_offset : rec_offset + b.challans]
                    first_s = batch_recs[0].get("SERIAL_NUMBER", "")
                    last_s = batch_recs[-1].get("SERIAL_NUMBER", "")
                    serial_range = f"{first_s} - {last_s}"
                    manifest_lines.append(f"{p.name:<55} | {serial_range:<15} | {b.challans:<8} | VERIFIED OK")
                    zf.write(p, arcname=p.name)
                    rec_offset += b.challans

                manifest_text = "\n".join(manifest_lines) + "\n"
                zf.writestr("manifest.txt", manifest_text)

            status_data["download_name"] = zip_name
            status_data["download_type"] = "zip"

        status_data["state"] = "completed"
        status_data["stage"] = "completed"
        status_data["verified"] = True
        status_data["download_ready"] = True
        status_data["message"] = f"{total_challans} of {total_challans} notices verified against your Excel data."
        update_status(job_dir, status_data)

    except Exception as e:
        status_data["state"] = "failed"
        status_data["stage"] = "failed"
        status_data["error"] = str(e)
        status_data["message"] = f"Job generation failed: {e}"
        update_status(job_dir, status_data)
        logger.exception("Job %s failed", job_dir.name)


def main():
    parser = argparse.ArgumentParser(description="Challan notice generation subprocess runner")
    parser.add_argument("--job-dir", required=True, help="Path to the job working directory")
    args = parser.parse_args()

    run_job(Path(args.job_dir))


if __name__ == "__main__":
    main()
