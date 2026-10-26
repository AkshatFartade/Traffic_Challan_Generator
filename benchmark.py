#!/usr/bin/env python3
"""
benchmark.py — generates synthetic challan CSVs at increasing sizes and runs
the generator against each one, producing a documented results table
(BENCHMARK_RESULTS.md) with timing, throughput, and pass/fail per batch.

Usage:
  python3 benchmark.py --sizes 1000 2000 3000 4000 5000 6000 7000 8000 9000 10000
  python3 benchmark.py --sizes 70000 --batch-size 1000   # single stress test
"""
import argparse
import csv
import platform
import random
import shutil
import time
from pathlib import Path

import notice_lib as nl

try:
    import psutil
    HAVE_PSUTIL = True
except ImportError:
    HAVE_PSUTIL = False

FIRST_NAMES = ["RAKESH", "JITENDRA", "SUNITA", "ANIL", "PRIYA", "RAVI", "DEEPAK", "SANGEETA", "VIKAS", "POOJA"]
LAST_NAMES = ["VISHVAKARMA", "KUMAR", "SHARMA", "VERMA", "PATEL", "SINGH", "SAHU", "YADAV", "DAS", "GUPTA"]
SECTIONS = [
    "129/194(D)Driving without helmet",
    "190(2). For 2 wheeler without valid documents",
    "(146/196)Driving an uninsured vehicle",
    "21(25)/177 Using mobile phone while driving",
]
CITIES = ["BHILAI DURG CG 490006", "KAWARDHA CG 491995", "RAIPUR CG 492001", "DURG CG 491001"]


def make_synthetic_csv(path: Path, n: int, seed: int = 42):
    rnd = random.Random(seed)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=nl.FIELD_KEYS)
        w.writeheader()
        for i in range(1, n + 1):
            w.writerow(
                {
                    "SERIAL_NUMBER": i,
                    "E_CHALLAN_NUMBER": f"CG{rnd.randint(10**15, 10**16-1)}",
                    "VEHICLE_NUMBER": f"CG07{rnd.choice('ABCDEFG')}{rnd.randint(1000,9999)}",
                    "VIOLATION_DATE": f"{rnd.randint(1,28):02d}/{rnd.randint(1,12):02d}/2026",
                    "NAME": f"{rnd.choice(FIRST_NAMES)} {rnd.choice(LAST_NAMES)}",
                    "ADDRESS": f"HOUSE NO {rnd.randint(1,999)} WARD {rnd.randint(1,40)} {rnd.choice(CITIES)}",
                    "MOBILE_NUMBER": f"9{rnd.randint(100000000,999999999)}",
                    "FINE_AMOUNT": rnd.choice([500, 1000, 1500]),
                    "SECTION": rnd.choice(SECTIONS),
                }
            )


def run_one(size: int, batch_size: int, template: Path, work_dir: Path, soffice: str) -> dict:
    csv_path = work_dir / f"synthetic_{size}.csv"
    out_dir = work_dir / f"out_{size}"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    make_synthetic_csv(csv_path, size)

    proc = psutil.Process() if HAVE_PSUTIL else None
    mem_before = proc.memory_info().rss / 1e6 if proc else None

    t0 = time.perf_counter()
    run = nl.generate_all(
        input_path=csv_path,
        template_path=template,
        output_dir=out_dir,
        batch_size=batch_size,
        make_zip=True,
        soffice_path=soffice,
    )
    elapsed = time.perf_counter() - t0
    mem_after = proc.memory_info().rss / 1e6 if proc else None

    total_pdf_bytes = sum(Path(b.pdf_path).stat().st_size for b in run.ok_batches if b.pdf_path)

    return {
        "challans": size,
        "batch_size": batch_size,
        "batches_total": len(run.batches),
        "batches_ok": len(run.ok_batches),
        "batches_failed": len(run.failed_batches),
        "seconds": round(elapsed, 1),
        "challans_per_sec": round(size / elapsed, 2) if elapsed else None,
        "output_mb": round(total_pdf_bytes / 1e6, 1),
        "mem_before_mb": round(mem_before, 1) if mem_before else "n/a",
        "mem_after_mb": round(mem_after, 1) if mem_after else "n/a",
        "failed_details": "; ".join(f"batch {b.batch_index}: {b.error}" for b in run.failed_batches),
    }


def write_report(results: list[dict], out_path: Path):
    lines = []
    lines.append("# Load-test results — Traffic Notice Generator (Python/LibreOffice)")
    lines.append("")
    lines.append(f"- Machine: {platform.platform()} — {platform.processor() or 'unknown CPU'}")
    if HAVE_PSUTIL:
        lines.append(f"- Logical CPUs: {psutil.cpu_count()}  |  RAM: {round(psutil.virtual_memory().total/1e9,1)} GB")
    lines.append(f"- Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")
    lines.append("| Challans | Batch size | Batches OK/Total | Time (s) | Challans/sec | Output size (MB) | Notes |")
    lines.append("|---:|---:|:---:|---:|---:|---:|---|")
    for r in results:
        notes = "OK" if r["batches_failed"] == 0 else f"FAILED: {r['failed_details']}"
        lines.append(
            f"| {r['challans']:,} | {r['batch_size']:,} | {r['batches_ok']}/{r['batches_total']} | "
            f"{r['seconds']} | {r['challans_per_sec']} | {r['output_mb']} | {notes} |"
        )
    out_path.write_text("\n".join(lines), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sizes", type=int, nargs="+", default=[1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000])
    ap.add_argument("--batch-size", type=int, default=1000)
    ap.add_argument("--template", default="templates/Formate_with_Sign_SectionSplit.docx")
    ap.add_argument("--work-dir", default="bench_work")
    ap.add_argument("--soffice", default="soffice")
    ap.add_argument("--report", default="BENCHMARK_RESULTS.md")
    args = ap.parse_args()

    work_dir = Path(args.work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    template = Path(args.template)

    results = []
    for size in args.sizes:
        print(f"\n=== Testing {size} challans (batch size {args.batch_size}) ===")
        r = run_one(size, args.batch_size, template, work_dir, args.soffice)
        print(r)
        results.append(r)
        write_report(results, Path(args.report))  # write after every size so partial progress is saved

    print(f"\nReport written to {args.report}")


if __name__ == "__main__":
    main()
