# Traffic Notice Generator --- Developer & Operations Documentation

**Baseline:** `A:\TWE_Challans_Generator`\
**Primary package:** `A:\TWE_Challans_Generator\notice_generator`\
**Approved template baseline:** `Formate_with_Sign_SectionSplit.docx`\
**Status:** Documentation of the existing Python/LibreOffice
implementation. **No changes are being made to the baseline project as
part of this documentation pass.**

------------------------------------------------------------------------

## 1. Purpose and scope

The Traffic Notice Generator converts traffic/e-challan records from
Excel or CSV into official-format PDF notices.

The current implementation is a Python rewrite of the older PowerShell +
Microsoft Word COM workflow. Its key architectural change is batching:

1.  Read challan records.
2.  Pair records two at a time.
3.  Fill the approved Word template for each pair.
4.  Combine many pairs into one temporary DOCX per batch.
5.  Convert each batch DOCX to PDF with headless LibreOffice.
6.  Verify that the PDF has exactly the expected number of pages.
7.  Keep successful batch PDFs and report failed batches independently.
8.  Optionally create one ZIP containing the successful batch PDFs.

The approved document layout is template-driven. The generator does not
redraw the notice, reconstruct borders, or manually recreate the
signature/layout.

### Important production principle

The approved template is a controlled artifact. Treat it as part of the
application binary/configuration, not as an ordinary document that users
should casually edit.

------------------------------------------------------------------------

## 2. Current architecture

### High-level flow

``` text
Excel (.xlsx/.xlsm/.xls)
        |
        |  openpyxl, sheet + start row
        v
Normalized Python records
        |
        |  or
        |
CSV (.csv) --------------------+
        |                       |
        v                       v
        +----> record validation / normalization
                         |
                         v
                  Pair records (2)
                         |
                         v
              Fill approved DOCX template
                         |
                         v
             Batch temporary DOCX
            (default: 1000 challans)
                         |
                         v
             Headless LibreOffice
                         |
                         v
                  Batch PDF
                         |
                         v
                pypdf page check
                         |
             +-----------+-----------+
             |                       |
          correct                  wrong/error
             |                       |
             v                       v
        final batch PDF          failed batch
             |
             v
       optional ZIP of
       successful PDFs
```

### Main components

  -------------------------------------------------------------------------------------
  Component                                         Responsibility
  ------------------------------------------------- -----------------------------------
  `generate_notices.py`                             Command-line interface and
                                                    progress/reporting

  `notice_lib.py`                                   Core data loading, template
                                                    manipulation, batching, LibreOffice
                                                    conversion, validation and result
                                                    objects

  `templates/Formate_with_Sign_SectionSplit.docx`   Approved visual/legal document
                                                    baseline

  `benchmark.py`                                    Synthetic CSV generation and
                                                    automated load-test runner

  `requirements.txt`                                Python package dependencies

  `tests/`                                          Small correctness fixtures

  `bench_work/`                                     Benchmark-generated input/output
                                                    working area

  `output/`                                         Normal production output directory

  `output_word/`                                    Historical/development output area;
                                                    not required by the current
                                                    production path

  `tmp/`                                            Temporary/development area; runtime
                                                    also creates its own temporary
                                                    directory by default
  -------------------------------------------------------------------------------------

------------------------------------------------------------------------

## 3. Folder/file structure

The inspected project snapshot contains the following relevant
structure:

``` text
A:\TWE_Challans_Generator\
└── notice_generator\
    ├── .venv\
    ├── benchmark.py
    ├── BENCHMARK_RESULTS.md
    ├── bench_work\
    │   ├── synthetic_1000.csv
    │   ├── synthetic_2000.csv
    │   └── out_2000\
    ├── generate_notices.py
    ├── notice_lib.py
    ├── output\
    │   └── .gitkeep
    ├── output_word\
    │   └── Traffic_Notices_batch_0001_1-4.pdf
    ├── README.md
    ├── requirements.txt
    ├── sample_benchmark_output_from_dev_sandbox.md
    ├── templates\
    │   └── Formate_with_Sign_SectionSplit.docx
    ├── TESTING.md
    ├── tests\
    │   ├── Traffic_Notice_Test_4_Records.xlsx
    │   └── Traffic_Notice_Test_1000_Records.xlsx
    └── tmp\
```

The ZIP snapshot also contains the local `.venv` and Python cache files.
A new developer should normally create a fresh virtual environment
rather than copy or depend on the existing `.venv`.

### Files that matter most

**Do not casually modify:**

-   `templates/Formate_with_Sign_SectionSplit.docx`
-   `notice_lib.py`
-   `generate_notices.py`

**Safe working areas:**

-   `output/`
-   `bench_work/`
-   `tmp/`

Benchmark documentation/results can be updated deliberately during the
later performance-testing phase.

------------------------------------------------------------------------

## 4. Python environment

The project is Windows-oriented in its current tested deployment, while
the core design is intentionally not tied to Microsoft Word.

The project snapshot contains a `.venv`, including Python 3.13 compiled
cache artifacts. The project code itself is written for modern Python
and the project documentation specifies Python 3.10+.

### Recommended fresh setup

Open PowerShell:

``` powershell
cd A:\TWE_Challans_Generator\notice_generator

py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1

python --version
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If PowerShell blocks script activation, the environment can still be
used directly:

``` powershell
A:\TWE_Challans_Generator\notice_generator\.venv\Scripts\python.exe --version
```

and packages can be installed with:

``` powershell
A:\TWE_Challans_Generator\notice_generator\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Python dependencies

Current `requirements.txt`:

``` text
openpyxl>=3.1
pypdf>=5.0
psutil>=5.9   # used by benchmark.py for memory statistics
```

The implementation uses:

-   **openpyxl** --- reads Excel workbooks.
-   **pypdf** --- counts/validates generated PDF pages.
-   **psutil** --- optional benchmark resource reporting; the benchmark
    can run without it, but RAM reporting is then unavailable.

The rest of the generator uses Python standard-library modules such as
`csv`, `zipfile`, `subprocess`, `tempfile`, `pathlib`,
`concurrent.futures`, `dataclasses`, and `time`.

------------------------------------------------------------------------

## 5. LibreOffice dependency

Microsoft Word is **not** part of the current PDF-generation path.

LibreOffice is required because `notice_lib.py` invokes its headless
command-line converter.

The known Windows installation used during development is:

``` text
C:\Program Files\LibreOffice\program\soffice.com
```

### Verify LibreOffice

PowerShell:

``` powershell
& "C:\Program Files\LibreOffice\program\soffice.com" --version
```

If LibreOffice is on `PATH`:

``` powershell
soffice --version
```

### Recommended explicit path

For repeatable production/benchmark runs, use the explicit executable
path:

``` powershell
$soffice = "C:\Program Files\LibreOffice\program\soffice.com"
```

and pass:

``` text
--soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

### Why the implementation uses a private LibreOffice profile

Every conversion creates a temporary LibreOffice user profile and passes
it through:

``` text
-env:UserInstallation=...
```

This is important when multiple batches are converted in parallel. It
prevents concurrent LibreOffice workers from fighting over the same
profile/lock files.

The implementation also uses:

``` text
--headless
--norestore
--nologo
--nofirststartwizard
--convert-to pdf
```

------------------------------------------------------------------------

## 6. Excel input format

The existing daily Excel compatibility mode intentionally mirrors the
earlier PowerShell workflow.

### Workbook rules

-   File types supported by the current reader: `.xlsx`, `.xlsm`,
    `.xls`.
-   Default worksheet: `Bike`.
-   Default first data row: **8**.
-   The reader takes columns **B through J**.
-   Rows where both `SERIAL_NUMBER` and `E_CHALLAN_NUMBER` are empty are
    skipped.

### Excel column mapping

  Excel column   Field                Meaning
  -------------- -------------------- --------------------------------
  B              `SERIAL_NUMBER`      Record serial number
  C              `E_CHALLAN_NUMBER`   E-challan identifier
  D              `VEHICLE_NUMBER`     Vehicle registration
  E              `VIOLATION_DATE`     Violation date/time/date value
  F              `NAME`               Person name
  G              `ADDRESS`            Address
  H              `MOBILE_NUMBER`      Mobile number
  I              `FINE_AMOUNT`        Fine amount
  J              `SECTION`            Offence/section text

The reader uses
`openpyxl.load_workbook(..., read_only=True, data_only=True)`, which is
appropriate for large input workbooks because it avoids loading the full
editable workbook model into memory.

### Date normalization

The generator accepts Excel serial date values, Python date/datetime
values, and several common text formats.

The normal output format is:

``` text
dd/MM/yyyy
```

If a date string cannot be recognized, the original text is retained
rather than silently discarded.

------------------------------------------------------------------------

## 7. CSV input format

CSV is the forward-looking interface for the eventual web application.

The CSV must contain these exact column names, in any order:

``` text
SERIAL_NUMBER
E_CHALLAN_NUMBER
VEHICLE_NUMBER
VIOLATION_DATE
NAME
ADDRESS
MOBILE_NUMBER
FINE_AMOUNT
SECTION
```

Example header:

``` csv
SERIAL_NUMBER,E_CHALLAN_NUMBER,VEHICLE_NUMBER,VIOLATION_DATE,NAME,ADDRESS,MOBILE_NUMBER,FINE_AMOUNT,SECTION
```

The CSV reader accepts UTF-8 with optional BOM (`utf-8-sig`).

Rows with neither a serial number nor e-challan number are skipped.

------------------------------------------------------------------------

## 8. Internal record normalization

Regardless of Excel or CSV input, the generator converts each record
into the same nine-key structure:

``` text
SERIAL_NUMBER
E_CHALLAN_NUMBER
VEHICLE_NUMBER
VIOLATION_DATE
NAME
ADDRESS
MOBILE_NUMBER
FINE_AMOUNT
SECTION
```

Text values are stripped of leading/trailing whitespace.

XML-sensitive characters are escaped before insertion into the Word
document XML. This is important because the template is edited at the
DOCX XML level rather than through the high-level Word object model.

------------------------------------------------------------------------

## 9. Approved template behavior

The baseline template is:

``` text
templates\Formate_with_Sign_SectionSplit.docx
```

It is a DOCX ZIP package containing the official document XML and
supporting assets, including the approved formatting/signature content.

The generator reads `word/document.xml` from the template package.

### Template sections

The template must contain at least two Word section properties:

1.  Front section
2.  Back section

The code extracts the approved pair content and the final section
definition.

It then repeats the pair structure for every two records.

### Placeholder system

The template contains placeholders in the form:

``` text
{{SERIAL_NUMBER}}
{{E_CHALLAN_NUMBER}}
{{VEHICLE_NUMBER}}
{{VIOLATION_DATE}}
{{NAME}}
{{ADDRESS}}
{{MOBILE_NUMBER}}
{{FINE_AMOUNT}}
{{SECTION}}
```

The code replaces these placeholders directly in `word/document.xml`.

### Important formatting behavior

The template/legal content is preserved.

Only the inserted data runs are normalized to the generator's current
data font size, **10 pt**. This prevents long data values from
inheriting an unsuitable placeholder formatting while leaving
official/static template text untouched.

The placeholder text is initially marked red in the approved template;
when real data is inserted, the relevant text is changed to normal
black.

------------------------------------------------------------------------

## 10. SECTION field processing

`SECTION` is special because the same source field supplies different
content to the front and back of the notice.

The generator splits section text into:

-   **section number** for the front
-   **description** for the back

Examples supported by the current regular expression include:

``` text
129/194(D)Driving without helmet
```

becoming approximately:

``` text
Front: 129/194(D)
Back : Driving without helmet
```

and:

``` text
190(2). For 2 wheeler without valid documents
```

becoming:

``` text
Front: 190(2)
Back : For 2 wheeler without valid documents
```

Multiple section components separated with `|` are handled as multiple
parts.

If a section value does not match the expected structure, the
implementation deliberately keeps the whole value on the front rather
than silently dropping official information.

------------------------------------------------------------------------

## 11. Front/back record mapping

Records are processed sequentially in pairs:

``` text
Record 1 + Record 2 -> pair 1
Record 3 + Record 4 -> pair 2
Record 5 + Record 6 -> pair 3
...
```

For each pair:

-   Front page contains the two challans in the approved top/bottom
    layout.
-   The immediately following back page contains the corresponding back
    information for those same two challans.

This gives the physical packet sequence:

``` text
Pair 1 front
Pair 1 back
Pair 2 front
Pair 2 back
Pair 3 front
Pair 3 back
...
```

### Odd record count

If the total number of records is odd, the final record is paired with
an internal empty record.

Therefore:

``` text
pairs = ceil(challans / 2)
expected pages = pairs * 2
```

For example:

    Challans   Pairs   Expected pages
  ---------- ------- ----------------
           1       1                2
           2       1                2
           3       2                4
           4       2                4
        1000     500             1000

------------------------------------------------------------------------

## 12. Batch processing

The default batch size is:

``` text
1000 challans
```

A batch is independent from every other batch.

For 70,000 challans with batch size 1,000:

``` text
70 batches
35,000 record pairs
70,000 expected PDF pages
```

Each batch becomes one temporary DOCX and one final PDF.

### Why batching exists

The old approach repeatedly opened Word and exported a tiny PDF for each
pair. At large volumes, thousands of Word startup/automation/export
operations dominate runtime.

The current approach reduces that overhead to roughly:

``` text
one LibreOffice conversion per batch
```

So with 1,000-challan batches:

``` text
70,000 challans -> 70 LibreOffice conversions
```

rather than tens of thousands of conversions.

### Batch independence

A failure in one batch does not intentionally invalidate successful
batches.

The run result records:

-   batch number
-   challan count
-   pair count
-   expected pages
-   actual pages
-   elapsed seconds
-   output path
-   success/failure
-   error message

------------------------------------------------------------------------

## 13. PDF generation workflow

For each batch:

### Step 1 --- Load template

The template is parsed and validated before batch processing begins.

### Step 2 --- Build pair content

Each two-record pair is substituted into the template.

### Step 3 --- Preserve Word section behavior

Intermediate pair section breaks are embedded in paragraphs so the
resulting merged document maintains the correct section boundaries. The
final section remains the document-level section property.

This is important to the front/back page layout.

### Step 4 --- Build one DOCX

The template ZIP is copied into a new DOCX package.

Only:

``` text
word/document.xml
```

is replaced with the generated merged XML.

Other template package assets are copied from the approved template.

### Step 5 --- Convert with LibreOffice

LibreOffice runs headlessly and creates the PDF.

### Step 6 --- Validate page count

Expected:

``` text
number_of_pairs * 2
```

If the actual PDF page count differs, the batch is failed.

### Step 7 --- Promote the PDF

Only after successful conversion and page-count validation is the
generated PDF moved into the final output directory.

### Step 8 --- Cleanup

The temporary batch DOCX is deleted. The overall temporary runtime
directory is removed when the run finishes.

------------------------------------------------------------------------

## 14. Page-count safety check

The page-count check is a key production safety mechanism.

For a batch containing `N` challans:

``` text
pairs = ceil(N / 2)
expected_pages = pairs * 2
```

The PDF is accepted only when:

``` text
actual_pages == expected_pages
```

A mismatch can indicate that content has overflowed the fixed layout,
for example because of unusually long name/address text.

This check does not prove every visual detail is correct; it proves that
the basic expected two-pages-per-pair structure survived PDF conversion.

For production acceptance, visual spot-checking remains necessary.

------------------------------------------------------------------------

## 15. Output naming

The current naming scheme is:

``` text
Traffic_Notices_batch_<batch-number>_<first-serial>-<last-serial>.pdf
```

Example:

``` text
Traffic_Notices_batch_0001_1-1000.pdf
Traffic_Notices_batch_0002_1001-2000.pdf
```

The serial values are sanitized for safe filenames.

If `--zip` is used, the successful PDFs are also placed into:

``` text
Traffic_Notices_all_batches.zip
```

Only successful batches are included in the ZIP.

------------------------------------------------------------------------

## 16. Command-line usage

### Basic CSV run

From:

``` text
A:\TWE_Challans_Generator\notice_generator
```

run:

``` powershell
python generate_notices.py `
  --input records.csv `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output
```

### Excel run using the production defaults

``` powershell
python generate_notices.py `
  --input daily.xlsx `
  --sheet Bike `
  --start-row 8 `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output
```

### Recommended explicit LibreOffice path

``` powershell
python generate_notices.py `
  --input daily.xlsx `
  --sheet Bike `
  --start-row 8 `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output `
  --batch-size 1000 `
  --workers 1 `
  --soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

### Production run with ZIP

``` powershell
python generate_notices.py `
  --input daily.xlsx `
  --sheet Bike `
  --start-row 8 `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output `
  --batch-size 1000 `
  --workers 1 `
  --zip `
  --soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

### Parallel conversion

After correctness and single-worker performance are established:

``` powershell
python generate_notices.py `
  --input daily.xlsx `
  --sheet Bike `
  --start-row 8 `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output `
  --batch-size 1000 `
  --workers 4 `
  --zip `
  --soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

Parallelism should be treated as a measured configuration choice, not an
assumption. LibreOffice conversion, CPU saturation, RAM pressure, disk
I/O, and temporary-file behavior all need to be measured.

### Machine-readable run summary

Add:

``` text
--summary-json output\run_summary.json
```

Example:

``` powershell
python generate_notices.py `
  --input daily.xlsx `
  --sheet Bike `
  --start-row 8 `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output `
  --batch-size 1000 `
  --summary-json output\run_summary.json `
  --soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

The JSON includes run totals and per-batch details.

------------------------------------------------------------------------

## 17. CLI options

  Option             Default                  Purpose
  ------------------ ------------------------ ---------------------------------
  `--input`          required                 `.csv` or Excel input
  `--template`       approved template path   DOCX template
  `--output`         `output`                 final output directory
  `--batch-size`     `1000`                   challans per PDF
  `--sheet`          `Bike`                   Excel worksheet
  `--start-row`      `8`                      first Excel data row
  `--zip`            off                      create ZIP of successful PDFs
  `--soffice`        `soffice`                LibreOffice executable
  `--workers`        `1`                      parallel batch worker processes
  `--summary-json`   none                     optional JSON result file

------------------------------------------------------------------------

## 18. Error handling and retries

The generator is intentionally batch-oriented.

### LibreOffice conversion

Each batch conversion has:

-   a default timeout of **600 seconds**
-   up to **2 attempts**
-   increasing delay between failed attempts

A conversion is successful only if:

1.  LibreOffice exits successfully, and
2.  the expected PDF file exists.

If conversion fails after the retries, the batch is marked failed.

### Page-count failure

If PDF conversion succeeds but page count is wrong, the batch is still
marked failed.

Typical message:

``` text
page count mismatch: expected X, got Y
```

The likely cause is layout overflow or another document-structure
problem.

### Overall process

The CLI exits with a non-zero status when any batch failed.

Successful batches are retained, so the operator can investigate and
rerun only the failed slice rather than automatically discarding the
complete run.

------------------------------------------------------------------------

## 19. Troubleshooting

### Problem: `soffice` not found

Check:

``` powershell
& "C:\Program Files\LibreOffice\program\soffice.com" --version
```

Then supply the explicit path with `--soffice`.

### Problem: LibreOffice exits with code 1

Check:

-   LibreOffice installation.
-   Whether the DOCX can be opened normally by LibreOffice.
-   Available disk space.
-   Whether another process is locking the input/output files.
-   Whether the batch temporary directory is writable.
-   Whether the machine is under heavy resource pressure.

The current implementation already uses a private LibreOffice profile
per conversion to reduce profile-lock conflicts.

### Problem: page-count mismatch

Check the failing batch's records, especially:

-   unusually long `NAME`
-   unusually long `ADDRESS`
-   unusual `SECTION` text
-   unexpected characters
-   malformed data

Do not immediately change the template. First identify the exact
record(s) causing the overflow.

### Problem: Excel says sheet not found

The default sheet is:

``` text
Bike
```

Check workbook sheet names or pass:

``` text
--sheet "<actual sheet name>"
```

### Problem: Excel records are missing

The Excel compatibility reader begins at row 8 and reads B:J.

Check:

-   correct worksheet
-   correct start row
-   serial/e-challan values
-   whether data is actually in B:J

### Problem: no eligible records

The generator raises an error when no usable records are found.

For Excel/CSV, a row is considered eligible when at least one of:

``` text
SERIAL_NUMBER
E_CHALLAN_NUMBER
```

is present.

### Problem: output directory contains old PDFs

Use a clean output directory for benchmark runs and controlled
production runs. Do not mix outputs from different test configurations.

For benchmarks, the existing benchmark script removes its per-size
output directory before generating the next result.

------------------------------------------------------------------------

## 20. Correctness testing before load testing

The small four-record test is the first gate.

Known development fixture:

``` text
tests\Traffic_Notice_Test_4_Records.xlsx
```

Run:

``` powershell
python generate_notices.py `
  --input tests\Traffic_Notice_Test_4_Records.xlsx `
  --sheet Bike `
  --start-row 8 `
  --template templates\Formate_with_Sign_SectionSplit.docx `
  --output output `
  --batch-size 1000 `
  --workers 1 `
  --soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

Expected structure for four challans:

``` text
4 challans
2 pairs
4 PDF pages
```

The four-page output should be visually checked for:

-   front-page top/bottom placement
-   matching front/back records
-   names
-   addresses
-   vehicle numbers
-   dates
-   e-challan numbers
-   fine amounts
-   section-number split
-   back-page section descriptions
-   mobile number
-   signature placement
-   pagination
-   font/spacing consistency

A successful page count alone is not sufficient for legal/operational
acceptance.

------------------------------------------------------------------------

## 21. Existing benchmark tooling

`benchmark.py` is already present in the project.

Its current role is to:

1.  Generate synthetic CSV data.
2.  Run the full generator.
3.  Record timing.
4.  Calculate challans/second.
5.  Record output size.
6.  Record process RSS before/after when `psutil` is available.
7.  Record successful/failed batches.
8.  Rewrite `BENCHMARK_RESULTS.md` after each completed size so partial
    progress survives a later failure.

Example:

``` powershell
python benchmark.py --sizes 1000 2000 3000 4000 5000 --batch-size 1000
```

For a target stress test:

``` powershell
python benchmark.py --sizes 70000 --batch-size 1000
```

The later systematic benchmark plan should supersede the current
development-only benchmark results.

------------------------------------------------------------------------

## 22. Existing benchmark evidence --- important distinction

The project contains historical benchmark material from development.

The development sandbox reported approximately:

``` text
~32 challans/sec
```

over 1,000--5,000-challan tests on a single-core Linux sandbox, with no
batch failures in those reported runs.

That environment was explicitly weaker than the target Windows laptop,
so those numbers are useful as a reference floor, **not as production
performance evidence for the target machine**.

The Windows `BENCHMARK_RESULTS.md` snapshot also contains a **failed
1,000-challan run** caused by LibreOffice conversion returning exit
code 1. Therefore the project must not claim that 1,000-challan Windows
performance has already been successfully established.

There is a separate successful four-record development run from the
current baseline work. That confirms the basic pipeline but is too small
to characterize large-volume performance.

### Reporting rule for the final performance report

Always separate:

**Measured** - actually executed on the specified machine - exact input
size - exact batch size - exact worker count - observed runtime -
observed output - observed resource behavior

from:

**Estimated** - extrapolated from measured throughput - model-based
projection - not directly executed at that volume

------------------------------------------------------------------------

## 23. Synthetic data strategy for the next phase

The later benchmark dataset should not use real personal information.

The synthetic generator should reproduce structural characteristics
rather than identities:

-   same nine columns
-   same data types
-   realistic vehicle/challan/date/amount formats
-   realistic Hindi/English text patterns
-   realistic short/medium/long names
-   realistic short/medium/long addresses
-   realistic section distribution
-   realistic fine-amount distribution
-   realistic record lengths
-   controlled long-text cases to test layout boundaries

The dataset generator should use a fixed random seed for
reproducibility.

A benchmark input should be identifiable by:

``` text
dataset version
seed
record count
generation timestamp
```

This makes performance results repeatable and auditable.

------------------------------------------------------------------------

## 24. Recommended systematic load-test progression

Do not begin with an unvalidated 70,000-challan run.

Recommended sequence:

``` text
2
10
100
500
1,000
5,000
10,000
20,000
30,000
50,000
70,000
```

At each significant step, confirm:

1.  All expected batches completed.
2.  All successful PDFs exist.
3.  Page counts are exact.
4.  Output size is plausible.
5.  Runtime is recorded.
6.  CPU/RAM behavior is recorded where possible.
7.  No sustained resource growth is appearing.
8.  LibreOffice processes terminate normally.
9.  No unexplained retries/failures occurred.

If a run becomes unstable, stop escalating and investigate before
increasing volume.

------------------------------------------------------------------------

## 25. Benchmark metrics

For every significant run record:

  Metric               Definition
  -------------------- -----------------------------------------
  Challans             Number of input records
  Batch size           Challans assigned to each batch
  Batches              Number of generated batches
  Workers              ProcessPool worker count
  Runtime              End-to-end wall-clock seconds
  Avg sec/challan      Runtime / challans
  Throughput           Challans / runtime
  Throughput/min       Challans/minute
  Successful batches   Completed and validated
  Failed batches       Failed after retry handling
  Expected pages       `ceil(challans/2) * 2`
  Actual pages         Sum of validated batch PDF pages
  Output size          Total PDF bytes
  ZIP size             Optional combined ZIP size
  CPU                  Observed utilization if measurable
  RAM                  Baseline/peak/post-run where measurable
  Disk                 Temporary/output storage behavior
  LibreOffice          Conversion failures/retries/timeouts
  Notes                Any anomalies

### Throughput formulas

``` text
challans/sec = challans / total_seconds
```

``` text
challans/min = challans / total_seconds * 60
```

``` text
seconds/challan = total_seconds / challans
```

### 70,000 estimate

If a stable measured throughput `T` challans/sec has been established:

``` text
estimated_seconds = 70,000 / T
estimated_minutes = estimated_seconds / 60
estimated_hours = estimated_minutes / 60
```

The final report should show the calculation and identify the
measurement(s) used.

------------------------------------------------------------------------

## 26. Resource monitoring

For production-scale testing, process RSS alone is not enough.

Where practical, monitor:

-   Python process CPU
-   LibreOffice process CPU
-   total system CPU
-   Python RSS
-   LibreOffice RSS
-   total system RAM used
-   available RAM
-   temporary directory size
-   output directory size
-   disk free space

On Windows, Task Manager can provide an operator-level cross-check.

`psutil` can be used in benchmark instrumentation for more reproducible
measurements.

### Important interpretation

The current `benchmark.py` records the benchmark process's RSS
before/after. That is useful but does **not** automatically capture peak
RAM consumed by all concurrent LibreOffice workers.

For the final production-scale benchmark, peak system/process behavior
should be measured explicitly rather than inferred from before/after
Python RSS.

------------------------------------------------------------------------

## 27. Scaling behavior to investigate

There are three important scaling dimensions.

### A. More challans, same batch size

Example:

``` text
1,000 -> 5,000 -> 10,000 -> 20,000 -> 70,000
```

If throughput stays approximately stable, the architecture is scaling
linearly with batch count.

### B. Different batch sizes

Test selected batch sizes such as:

``` text
250
500
1,000
2,000
5,000
```

Larger batches reduce the number of LibreOffice launches, but increase:

-   temporary DOCX size
-   conversion duration per batch
-   failure blast radius
-   memory pressure
-   retry cost

The optimal production batch size therefore must be measured.

### C. Different worker counts

For the target machine, compare:

``` text
workers = 1
workers = 2
workers = 4
```

Potential outcomes:

-   near-linear improvement
-   diminishing returns
-   RAM saturation
-   LibreOffice contention
-   disk I/O bottleneck

Do not assume four workers is automatically optimal.

------------------------------------------------------------------------

## 28. Production operating procedure

### Before a production run

1.  Confirm the approved template hash/version.
2.  Confirm Python environment.
3.  Confirm required packages.
4.  Confirm LibreOffice version/path.
5.  Confirm input workbook name and worksheet.
6.  Confirm start row.
7.  Confirm available disk space.
8.  Use a clean output directory.
9.  Run a small validation sample when the template/data source has
    changed.
10. Record the run configuration.

### During the run

Watch:

-   batch progress
-   failed batches
-   LibreOffice errors/retries
-   CPU
-   RAM
-   disk space
-   output growth

Do not repeatedly launch overlapping large production jobs on the same
machine unless concurrency has been specifically tested.

### After the run

Confirm:

``` text
successful batches == expected batches
failed batches == 0
actual pages == expected pages
```

Then:

1.  Inspect sample PDFs.
2.  Inspect a middle batch.
3.  Inspect the final batch.
4.  Confirm ZIP contents if ZIP was requested.
5.  Preserve the run summary/log.
6.  Archive the input, configuration, template version/hash, and result
    summary according to operational policy.

------------------------------------------------------------------------

## 29. Production acceptance checklist

``` text
[ ] Correct input workbook/CSV
[ ] Correct worksheet
[ ] Correct start row
[ ] Approved template
[ ] Python environment verified
[ ] LibreOffice verified
[ ] Output directory clean
[ ] Small correctness test passed when required
[ ] Batch size recorded
[ ] Worker count recorded
[ ] Run started
[ ] No unexplained conversion failures
[ ] All expected batches successful
[ ] All expected pages generated
[ ] PDF spot-check completed
[ ] Output/ZIP verified
[ ] Run summary archived
```

------------------------------------------------------------------------

## 30. Current limitations / engineering observations

1.  The Excel reader is intentionally tied to the existing B:J / row-8 /
    `Bike` layout for backward compatibility.
2.  CSV is the cleaner forward-looking interface.
3.  The current page-count validation detects structural pagination
    failures but cannot replace visual inspection.
4.  Peak system resource monitoring is not yet comprehensive in the
    existing benchmark script.
5.  Parallel processing has been implemented but must be benchmarked on
    the target hardware before being declared the production default.
6.  A failed batch can be retained/recovered independently, but a formal
    operator workflow for identifying and rerunning only failed ranges
    should be documented after the load tests.
7.  The existing benchmark history contains both successful development
    results and a failed Windows LibreOffice run; these must not be
    mixed together as if they were one clean benchmark series.
8.  The current architecture still creates a potentially large temporary
    DOCX for each batch. This should be observed during 50,000--70,000
    tests rather than assumed harmless.
9.  The eventual website can call the same core `generate_all()` engine,
    but web-server orchestration, job queues, authentication, file
    isolation, quotas, and asynchronous job handling are separate
    production concerns.

------------------------------------------------------------------------

## 31. Developer extension point

The main reusable API is:

``` python
notice_lib.generate_all(...)
```

This makes the core engine suitable for later integration into a
Flask/FastAPI-style upload service.

A future web layer should ideally:

``` text
Upload
  -> validate
  -> store temporary input
  -> call generate_all()
  -> expose progress/status
  -> return generated ZIP/PDFs
  -> clean temporary files
```

The document-generation engine should remain separate from HTTP/UI code.

------------------------------------------------------------------------

## 32. Baseline integrity rule

The current approved baseline should be treated as frozen during
performance characterization.

**Do not change the template or generation algorithm merely to improve
benchmark numbers.**

If a change becomes necessary:

1.  record the reason,
2.  create a clearly versioned change,
3.  rerun the correctness suite,
4.  rerun affected benchmarks,
5.  distinguish pre-change and post-change measurements.

This prevents performance testing from accidentally becoming an
uncontrolled software-change exercise.

------------------------------------------------------------------------

## 33. Final performance-report structure

The eventual report should contain:

### Executive summary

-   purpose
-   production target
-   measured best/stable configuration
-   measured throughput
-   measured maximum tested volume
-   evidence-based 70,000 estimate

### System under test

-   hardware
-   OS
-   Python version
-   LibreOffice version
-   dependency versions
-   storage environment

### Application architecture

-   data flow
-   batching
-   template processing
-   PDF conversion
-   page validation

### Test methodology

-   synthetic-data design
-   seeds
-   test volumes
-   batch sizes
-   worker counts
-   measurement method

### Results

A table such as:

  ---------------------------------------------------------------------------------------------
    Challans   Batch   Workers   Batches    Time   Challans/s   Pages   Output    Peak Result
                                                                                   RAM 
  ---------- ------- --------- --------- ------- ------------ ------- -------- ------- --------

  ---------------------------------------------------------------------------------------------

### Failure/retry analysis

-   failed batches
-   conversion retries
-   page mismatches
-   causes
-   recovery behavior

### Scaling analysis

-   throughput vs volume
-   throughput vs batch size
-   throughput vs workers
-   resource behavior

### 70,000-challan assessment

Clearly split into:

**Measured:** largest volume actually executed.

**Estimated:** 70,000 projection based on the measured stable region.

### Production recommendation

-   recommended batch size
-   recommended worker count
-   expected runtime range
-   operational procedure
-   monitoring
-   recovery procedure

------------------------------------------------------------------------

## 34. Current baseline evidence

The available development artifacts include:

-   approved template `Formate_with_Sign_SectionSplit.docx`
-   Python project ZIP snapshot
-   four-record Excel correctness fixture
-   1,000-record Excel test fixture
-   benchmark scripts and benchmark documentation
-   successful small PDF outputs
-   a historical benchmark result set

The approved template and existing working generator remain the baseline
for the next phases.

------------------------------------------------------------------------

## 35. Phase transition rule

The project should move to synthetic-data generation only after this
documentation/baseline-understanding phase is accepted.

Then:

``` text
Phase 1: Documentation
        |
        v
Phase 2: Synthetic dataset
        |
        v
Phase 3: Progressive load tests
        |
        v
Phase 4: Benchmark analysis
        |
        v
Phase 5: Technical & Performance Report
```

The final report must never present an extrapolated 70,000 result as a
directly measured result.
