#!/usr/bin/env python3
"""Deterministic, privacy-safe synthetic challan dataset generator.

Generates CSV or Excel files matching the Traffic Notice Generator schema.
No source personal records are read. A fixed seed makes datasets reproducible.
"""
from __future__ import annotations

import argparse, csv, random
from datetime import date, timedelta
from pathlib import Path

FIELD_KEYS = [
    "SERIAL_NUMBER", "E_CHALLAN_NUMBER", "VEHICLE_NUMBER", "VIOLATION_DATE",
    "NAME", "ADDRESS", "MOBILE_NUMBER", "FINE_AMOUNT", "SECTION",
]

# Fictional/synthetic token pools chosen to resemble common Indian notice text
# lengths and language patterns without copying any real record.
FIRST = ["AARAV", "VIHAAN", "KABIR", "ISHAN", "NEEL", "RITVIK", "TANAY", "ADITYA",
         "ANAYA", "DIYAA", "MYRA", "NAVYA", "RIYA", "TARA", "IRA", "KAVYA"]
LAST = ["KESHAR", "DEVAN", "RAJVAN", "MEHRAJ", "SAINIK", "VERDAN", "KAPOORAN", "MALVIK"]
LOCALITIES = ["SHASTRI NAGAR", "NEHRU WARD", "GANDHI CHOWK", "CIVIL LINES", "STATION ROAD",
              "VIDYA NAGAR", "RAM NAGAR", "SHANTI VIHAR", "KANCHAN BAGH", "GREEN PARK"]
SECTIONS = [
    ("129/194(D)Driving without helmet", 0.42),
    ("190(2). For 2 wheeler without valid documents", 0.24),
    ("(146/196)Driving an uninsured vehicle", 0.18),
    ("21(25)/177 Using mobile phone while driving", 0.10),
    ("184/177 Dangerous driving / unsafe driving", 0.06),
]
FINES = [(500, 0.48), (1000, 0.32), (1500, 0.14), (2000, 0.04), (3000, 0.02)]
CITIES = ["RAIPUR CG 492001", "DURG CG 491001", "BHILAI CG 490006", "KAWARDHA CG 491995"]


def weighted(rnd, items):
    values, weights = zip(*items)
    return rnd.choices(values, weights=weights, k=1)[0]


def make_record(rnd, serial, base_date=date(2026, 1, 1)):
    day = base_date + timedelta(days=rnd.randrange(0, 365))
    # Mix short/medium/long addresses to exercise layout realistically.
    addr_style = rnd.choices(["short", "medium", "long"], weights=[0.20, 0.65, 0.15], k=1)[0]
    if addr_style == "short":
        address = f"HOUSE {rnd.randint(1, 999)}, {rnd.choice(LOCALITIES)}"
    elif addr_style == "long":
        address = (f"HOUSE NO {rnd.randint(1, 999)}, GALI {rnd.randint(1, 12)}, "
                   f"WARD {rnd.randint(1, 50)}, {rnd.choice(LOCALITIES)}, {rnd.choice(CITIES)}")
    else:
        address = f"HOUSE NO {rnd.randint(1, 999)}, WARD {rnd.randint(1, 50)}, {rnd.choice(LOCALITIES)}, {rnd.choice(CITIES)}"

    # Stable synthetic identifiers. They are deliberately not derived from real data.
    challan = f"SYN26{serial:012d}"
    vehicle = f"CG{rnd.randint(1, 99):02d}{rnd.choice('ABCDEFGH')}{rnd.randint(1000, 9999):04d}"
    mobile = f"9{rnd.randint(100000000, 999999999)}"
    name = f"{rnd.choice(FIRST)} {rnd.choice(LAST)}"

    return {
        "SERIAL_NUMBER": serial,
        "E_CHALLAN_NUMBER": challan,
        "VEHICLE_NUMBER": vehicle,
        "VIOLATION_DATE": day.strftime("%d/%m/%Y"),
        "NAME": name,
        "ADDRESS": address,
        "MOBILE_NUMBER": mobile,
        "FINE_AMOUNT": weighted(rnd, FINES),
        "SECTION": weighted(rnd, SECTIONS),
    }


def generate_csv(path: Path, n: int, seed: int):
    rnd = random.Random(seed)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELD_KEYS)
        w.writeheader()
        for i in range(1, n + 1):
            w.writerow(make_record(rnd, i))


def generate_xlsx(path: Path, n: int, seed: int, sheet: str = "Bike", start_row: int = 8):
    from openpyxl import Workbook
    rnd = random.Random(seed)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    # Match production Excel placement: fields B:J, first data row 8.
    for row in range(1, start_row):
        ws.cell(row=row, column=1, value="")
    for col, key in enumerate(FIELD_KEYS, start=2):
        ws.cell(row=start_row - 1, column=col, value=key)
    for i in range(1, n + 1):
        rec = make_record(rnd, i)
        for col, key in enumerate(FIELD_KEYS, start=2):
            ws.cell(row=start_row + i - 1, column=col, value=rec[key])
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="Generate privacy-safe synthetic challan data")
    ap.add_argument("--sizes", nargs="+", type=int, default=[2, 10, 100, 500, 1000])
    ap.add_argument("--out-dir", default="synthetic_data")
    ap.add_argument("--seed", type=int, default=20260911)
    ap.add_argument("--format", choices=["csv", "xlsx", "both"], default="both")
    args = ap.parse_args()
    out = Path(args.out_dir)
    for n in args.sizes:
        if n < 1:
            raise SystemExit("size must be >= 1")
        if args.format in ("csv", "both"):
            generate_csv(out / f"synthetic_{n}.csv", n, args.seed)
        if args.format in ("xlsx", "both"):
            generate_xlsx(out / f"synthetic_{n}.xlsx", n, args.seed)
        print(f"generated {n:,} records (seed={args.seed})")

if __name__ == "__main__":
    main()
