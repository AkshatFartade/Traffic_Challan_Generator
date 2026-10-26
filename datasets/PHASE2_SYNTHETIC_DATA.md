# Phase 2 — Synthetic Test Data

## Privacy rule

All records are generated from synthetic token pools. No source challan workbook is read and no real personal record is copied.

## Reproducibility

Default seed: `20260911`

The same seed + size produces the same dataset content.

## Schema

`SERIAL_NUMBER,E_CHALLAN_NUMBER,VEHICLE_NUMBER,VIOLATION_DATE,NAME,ADDRESS,MOBILE_NUMBER,FINE_AMOUNT,SECTION`

For Excel output the fields are placed in **B:J** with the header on row 7 and data beginning on row 8, worksheet `Bike`, matching the production reader.

## Distribution model

The generator intentionally approximates the *structure* and text-length behavior of the known challan workflow. It does **not** claim to reproduce an empirical distribution from real source data because the synthetic generator must not ingest personal records.

- Section mix: weighted common offence patterns.
- Fine amounts: weighted toward 500/1000/1500.
- Names: short/medium two-token names.
- Addresses: 20% short, 65% medium, 15% long.
- Dates: uniformly distributed across 2026.
- Vehicle/challan/mobile values: unique synthetic identifiers.
- English/Hinglish notice text: representative of the existing section formats.

## Initial test sizes

2, 10, 100, 500, 1,000, 5,000, 10,000, 20,000, 30,000, 50,000, 70,000.

Do not pre-generate 70,000 until the smaller PDF tests are passing; generate it on demand during the load-test phase.
