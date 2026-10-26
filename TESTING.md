# Testing protocol — Traffic Notice Generator (Python edition)

## Step 0 — Correctness check (already done during development)

Before any load testing, the output was verified against the approved
template using the real `tests/Traffic_Notice_Test_4_Records.xlsx` file:
- Front page: 2 challans per page (top/bottom half), correct name/address/
  vehicle number/date/section-number/e-challan number/fine amount, correct
  section-number-vs-description split, signature in place.
- Back page: correct name/address/mobile number/section description,
  correct box position for each of the two challans.
- Page count exactly 2 pages per pair, as required.

Do this again yourself first, on your machine, before trusting any bigger
run: `python generate_notices.py --input tests/Traffic_Notice_Test_4_Records.xlsx ...`
then open the PDF and eyeball it.

## Step 1 — Load testing (1,000 → 10,000 → 70,000)

`benchmark.py` generates synthetic challan data at each size, runs the
full pipeline, and checks:
- Did every batch produce a PDF? (batches OK / batches total)
- Does every batch have the exact expected page count?
- How long did it take, and what's the challans/sec rate?
- Did memory usage stay stable (no runaway growth as size increases)?

Run it yourself:
```
python benchmark.py --sizes 1000 2000 3000 4000 5000 6000 7000 8000 9000 10000 --workers 1
python benchmark.py --sizes 70000 --workers 4          # the real target volume
```
It writes `BENCHMARK_RESULTS.md` after every size (so a long run's
progress is never lost even if a later size fails), and with `psutil`
installed it also records RAM before/after each size.

Try `--workers 4` (or however many cores your CPU has) once you've
confirmed `--workers 1` works cleanly — batches are independent, so this
should give a close-to-linear speedup on a multi-core machine.

### Results gathered so far (this development sandbox)

Machine used below: single-core Linux sandbox, 4.2 GB RAM — **weaker than
your i5-10th-gen-U laptop**, so treat these as a worst-case floor, not
what you should expect on your own hardware.

| Challans | Batch size | Batches OK/Total | Time (s) | Challans/sec | Output size (MB) | Notes |
|---:|---:|:---:|---:|---:|---:|---|
| 1,000 (real test file) | 1,000 | 1/1 | 34.1 | 29.3 | 16.0 | OK |
| 1,000 (synthetic) | 1,000 | 1/1 | 31.8 | 31.4 | 16.0 | OK |
| 2,000 | 1,000 | 2/2 | 61.9 | 32.3 | 32.0 | OK |
| 3,000 | 1,000 | 3/3 | 94.0 | 31.9 | 47.9 | OK |
| 4,000 | 1,000 | 4/4 | 124.4 | 32.2 | 63.9 | OK |
| 5,000 | 1,000 | 5/5 | 155.7 | 32.1 | 79.9 | OK |

Observations from this range:
- Throughput is flat at **~32 challans/sec** regardless of size — no
  slowdown as volume grows, and no batch ever failed a page-count check.
- Memory grew by roughly 15-30 MB per extra 1,000-challan batch and never
  spiked, which is expected (each batch is independent and its temp files
  are cleaned up immediately after conversion).
- At this rate, 10,000 challans ≈ 5 minutes, and 70,000 challans ≈
  36 minutes **on a single CPU core with no parallelism**. `--workers 4`
  on your quad-core laptop should cut that roughly to 9-10 minutes for
  70,000 — please confirm this with the real run below.

### Your turn — fill this in on your laptop (i5-10th gen U, 8 GB RAM)

| Challans | Batch size | Workers | Batches OK/Total | Time | Challans/sec | Notes |
|---:|---:|---:|:---:|---:|---:|---|
| 1,000  |       |   |   |   |   |   |
| 2,000  |       |   |   |   |   |   |
| 3,000  |       |   |   |   |   |   |
| 4,000  |       |   |   |   |   |   |
| 5,000  |       |   |   |   |   |   |
| 6,000  |       |   |   |   |   |   |
| 7,000  |       |   |   |   |   |   |
| 8,000  |       |   |   |   |   |   |
| 9,000  |       |   |   |   |   |   |
| 10,000 |       |   |   |   |   |   |
| 70,000 |       |   |   |   |   |   |

Things worth watching specifically at 70,000 on 8 GB RAM:
- Task Manager memory for `python.exe` and any `soffice.exe` processes —
  if it climbs steadily instead of staying flat, note it here.
- Whether any individual batch fails the page-count check (this usually
  means one record's NAME/ADDRESS text is unusually long and needs a
  smaller font or the fixed layout revisited — the run will tell you
  exactly which batch and can be re-run for just that slice).
- Total wall-clock time and whether `--workers 4` actually gives a
  speedup over `--workers 1` on your machine (UHD integrated graphics
  don't matter here — this is CPU/RAM-bound, not GPU-bound).

## Step 2 — Spot-check the actual PDFs, not just the numbers

Load testing tells you it *ran*; it doesn't tell you the *printed output*
is right. For at least the 1,000 and 70,000 tests, open a handful of
pages from the middle and end of the output (not just page 1) and check
against a couple of real challan records by hand.
