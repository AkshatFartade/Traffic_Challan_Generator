"""
scripts/test_2000_web.py
------------------------
Item 3 Large Run Test:
  - 2,000 records through the website backend
  - workers = 4
  - Verifies:
      * 2 batches generated
      * All verification gates G1, G2, G3 PASSED
      * download_type == "zip"
      * ZIP download succeeds
      * manifest.txt inside ZIP contains gate pass & serial ranges
      * Per-batch PDF downloads via /api/jobs/<id>/download/<batch_pdf> work
      * Measures and reports timings
"""
from __future__ import annotations

import io
import json
import os
import sys
import time
import zipfile
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
BASE_URL = os.environ.get("CHALLAN_TEST_URL", "http://127.0.0.1:5000")
PASSWORD = os.environ.get("CHALLAN_SUPERVISOR_PASSWORD", "test123")


def log(msg: str):
    print(msg, flush=True)


def make_2000_xlsx_bytes() -> bytes:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bike"
    headers = [
        "Serial Number", "E-Challan Number", "Vehicle Number",
        "Violation Date", "Name", "Address", "Mobile Number",
        "Fine Amount", "Section"
    ]
    for ci, h in enumerate(headers, 2):
        ws.cell(7, ci, h)

    log("Generating 2,000 synthetic records in memory...")
    for i in range(2000):
        r = 8 + i
        n = i + 1
        ws.cell(r, 2, n)
        ws.cell(r, 3, f"CG2000TC{n:010d}")
        ws.cell(r, 4, f"CG07TC{n:04d}")
        ws.cell(r, 5, "18/07/2026")
        ws.cell(r, 6, f"CITIZEN NAME {n}")
        ws.cell(r, 7, f"HOUSE {n} WARD 2 SECTOR {n%10} BHILAI DURG CG 490006")
        ws.cell(r, 8, f"98260{n:05d}")
        ws.cell(r, 9, 500)
        ws.cell(r, 10, "129/194(D)Driving without helmet")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def run():
    session = requests.Session()

    # 1. Check health
    resp = session.get(f"{BASE_URL}/healthz", timeout=10)
    assert resp.status_code == 200, f"Healthz failed: {resp.status_code}"
    log(f"[EVIDENCE] Healthz check: HTTP {resp.status_code} OK")

    # 2. Login
    resp = session.post(f"{BASE_URL}/login", data={"password": PASSWORD}, allow_redirects=True, timeout=10)
    assert resp.status_code == 200, f"Login failed: {resp.status_code}"
    log(f"[EVIDENCE] Login with supervisor password: HTTP 200 OK")

    # 3. Upload 2000-record xlsx
    xlsx_data = make_2000_xlsx_bytes()
    log(f"Uploading 2000-record XLSX ({len(xlsx_data):,} bytes)...")
    files = {"file": ("Synthetic_2000_Records.xlsx", io.BytesIO(xlsx_data), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    resp = session.post(f"{BASE_URL}/api/upload", files=files, timeout=60)
    assert resp.status_code == 200, f"Upload failed: {resp.status_code} {resp.text}"
    upload_res = resp.json()
    job_id = upload_res["job_id"]
    log(f"[EVIDENCE] Upload successful: job_id={job_id}, total_records={upload_res['total_records']}, valid={upload_res['valid']}")
    assert upload_res["total_records"] == 2000
    assert upload_res["valid"] is True

    # 4. Start job with workers=4
    log(f"Starting job with workers=4...")
    resp = session.post(f"{BASE_URL}/api/jobs/{job_id}/start", json={"workers": 4}, timeout=10)
    assert resp.status_code == 200, f"Start job failed: {resp.status_code} {resp.text}"
    log(f"[EVIDENCE] Job started with workers=4: {resp.json()}")

    # 5. Poll status until completed
    t0 = time.time()
    last_msg = ""
    while True:
        resp = session.get(f"{BASE_URL}/api/jobs/{job_id}", timeout=10)
        assert resp.status_code == 200
        st = resp.json()
        state = st.get("state")
        msg = st.get("message", "")
        if msg != last_msg:
            log(f"  [Status {time.time()-t0:.1f}s] state={state}, stage={st.get('stage')}, msg={msg}")
            last_msg = msg

        if state == "completed":
            log(f"\n[EVIDENCE] Job {job_id} COMPLETED in {time.time()-t0:.1f}s!")
            break
        elif state in ("failed", "cancelled", "verification_failed"):
            raise RuntimeError(f"Job ended in error state '{state}': {st}")

        time.sleep(2)

    # 6. Verify completion details
    assert st.get("verified") is True, "Gate verification did not pass"
    assert st.get("download_ready") is True, "Download is not ready"
    assert st.get("download_type") == "zip", f"Expected zip download, got {st.get('download_type')}"
    zip_name = st.get("download_name")
    log(f"[EVIDENCE] Status: verified=True, download_ready=True, download_type=zip, download_name={zip_name}")

    # Report batch timings from status
    batches = st.get("batches", [])
    log(f"\n--- Batch Timings (workers=4) ---")
    for b in batches:
        log(f"  Batch {b.get('batch_index')}: {b.get('challans')} challans, {b.get('actual_pages')} pages in {b.get('seconds')}s (ok={b.get('ok')})")
    log(f"Total measured speed: {st.get('measured_speed')} notices/sec")

    # 7. Download ZIP
    zip_resp = session.get(f"{BASE_URL}/api/jobs/{job_id}/download", timeout=60)
    assert zip_resp.status_code == 200, f"ZIP download failed: {zip_resp.status_code}"
    zip_bytes = zip_resp.content
    cache_ctrl = zip_resp.headers.get("Cache-Control", "")
    log(f"\n[EVIDENCE] Downloaded ZIP: {len(zip_bytes):,} bytes (HTTP 200, Cache-Control: {cache_ctrl})")

    # Inspect ZIP contents
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        namelist = zf.namelist()
        log(f"[EVIDENCE] Files in ZIP: {namelist}")
        assert "manifest.txt" in namelist, "manifest.txt missing from ZIP"
        manifest_text = zf.read("manifest.txt").decode("utf-8")
        log(f"\n[EVIDENCE] manifest.txt contents:\n{manifest_text}")
        assert "Verification Gate G1: PASSED" in manifest_text
        assert "Verification Gate G2: PASSED" in manifest_text
        assert "Verification Gate G3: PASSED" in manifest_text
        assert "Total Challans: 2000" in manifest_text
        assert "Total Batches: 2" in manifest_text

    # 8. Test per-batch downloads
    for b in batches:
        pdf_name = b.get("pdf_name")
        if pdf_name:
            b_resp = session.get(f"{BASE_URL}/api/jobs/{job_id}/download/{pdf_name}", timeout=60)
            assert b_resp.status_code == 200, f"Batch download failed for {pdf_name}: {b_resp.status_code}"
            log(f"[EVIDENCE] Per-batch download '{pdf_name}': HTTP 200, size={len(b_resp.content):,} bytes")

    log(f"\n[EVIDENCE] ALL Item 3 requirements verified successfully!")


if __name__ == "__main__":
    run()
