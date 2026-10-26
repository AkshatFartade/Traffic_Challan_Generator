"""
routes.py — Flask blueprints and API endpoints for Traffic Notice Generator.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import time
from pathlib import Path

from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
import pypdfium2 as pdfium

import auth
import config
import input_inspector as ii

main_bp = Blueprint("main", __name__)
api_bp = Blueprint("api", __name__, url_prefix="/api")
auth_bp = Blueprint("auth", __name__)

# Active background subprocess: {job_id: subprocess.Popen}
_ACTIVE_PROCESSES: dict[str, subprocess.Popen] = {}


def get_job_dir(job_id: str) -> Path:
    """Safe job directory resolver with strict path traversal protection."""
    if not re.match(r"^[A-Za-z0-9_-]{16,64}$", job_id):
        abort(400, description="Invalid job identifier format.")
    job_dir = config.JOBS_DIR / job_id
    if not job_dir.exists():
        abort(404, description="Job directory not found or expired.")
    return job_dir


# --- MAIN & HEALTH ROUTES ---

@main_bp.route("/")
@auth.login_required
def index():
    return render_template("index.html")


@main_bp.route("/healthz")
def healthz():
    return "ok", 200, {"Content-Type": "text/plain"}


# --- AUTH ROUTES ---

@auth_bp.route("/login", methods=["GET", "POST"])
def login_page():
    if not config.SUPERVISOR_PASSWORD:
        return render_template(
            "login.html",
            error="Supervisor password is not configured on this server. Set CHALLAN_SUPERVISOR_PASSWORD environment variable.",
            unconfigured=True,
        )

    if request.method == "POST":
        ip = request.remote_addr or "127.0.0.1"
        allowed, remaining = auth.check_rate_limit(ip)
        if not allowed:
            return render_template(
                "login.html",
                error=f"Too many failed login attempts. Please wait {remaining} seconds.",
                locked=True,
            ), 429

        password = request.form.get("password", "")
        valid, err_msg = auth.verify_supervisor_password(password)
        if valid:
            auth.clear_failed_attempts(ip)
            session["is_supervisor"] = True
            session.permanent = True
            next_url = request.args.get("next") or url_for("main.index")
            return redirect(next_url)
        else:
            auth.record_failed_attempt(ip)
            return render_template("login.html", error=err_msg), 401

    return render_template("login.html")


@auth_bp.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("auth.login_page"))


# --- API ROUTES ---

@api_bp.route("/upload", methods=["POST"])
@auth.login_required
def upload_file():
    """Stores uploaded file, runs Input Inspector & Preflight, returns report."""
    if "file" not in request.files:
        return jsonify({"error": "No file part in request"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No selected file"}), 400

    filename = Path(file.filename).name
    # Check extension
    suffix = Path(filename).suffix.lower()
    if suffix not in (".xlsx", ".csv"):
        if suffix == ".xls":
            return jsonify({
                "error": "Unsupported file format",
                "message": "Excel 97-2003 (.xls) format is not supported. Please save as modern Excel (.xlsx) or CSV (.csv)."
            }), 400
        return jsonify({
            "error": "Unsupported file format",
            "message": f"Only .xlsx and .csv files are supported. Received '{suffix}'."
        }), 400

    # Create random job folder
    job_id = secrets.token_hex(16)
    job_dir = config.JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    input_path = job_dir / f"input{suffix}"
    file.save(input_path)

    # Check size
    if input_path.stat().st_size > config.MAX_UPLOAD_SIZE_BYTES:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({
            "error": "File too large",
            "message": f"Uploaded file exceeds maximum limit of {config.MAX_UPLOAD_SIZE_BYTES // (1024*1024)} MB."
        }), 400

    preferred_sheet = request.form.get("sheet_name")

    try:
        report = ii.inspect_file(input_path, preferred_sheet=preferred_sheet)
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": "Inspection failed", "message": str(e)}), 400

    # Export problems .xlsx if any flagged rows exist
    problems_path = None
    if report.blocking_errors or report.warnings:
        problems_file = job_dir / "Flagged_Rows_Report.xlsx"
        ii.export_problems_xlsx(report.blocking_errors, report.warnings, problems_file)
        problems_path = str(problems_file)

    # Save job spec
    spec = {
        "job_id": job_id,
        "input_path": str(input_path),
        "filename": filename,
        "file_type": report.file_type,
        "sheet_name": report.sheet_name,
        "start_row": report.start_row,
        "total_records": report.total_records,
        "blocking_errors": report.blocking_errors,
        "warnings": report.warnings,
        "problems_file": problems_path,
        "batch_size": config.DEFAULT_BATCH_SIZE,
        "workers": config.DEFAULT_WORKERS,
    }
    with open(job_dir / "job_spec.json", "w", encoding="utf-8") as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)

    return jsonify({
        "job_id": job_id,
        "filename": filename,
        "file_type": report.file_type,
        "sheet_name": report.sheet_name,
        "available_sheets": report.available_sheets,
        "needs_sheet_selection": report.needs_sheet_selection,
        "start_row": report.start_row,
        "total_records": report.total_records,
        "valid": report.valid,
        "blocking_count": len(report.blocking_errors),
        "blocking_errors": report.blocking_errors[:50],  # cap preview list
        "warnings_count": len(report.warnings),
        "warnings": report.warnings[:50],
        "first_5": report.first_5,
        "last_5": report.last_5,
        "has_problems_file": problems_path is not None,
    })


@api_bp.route("/jobs/<job_id>/start", methods=["POST"])
@auth.login_required
def start_job(job_id: str):
    """Launches the standalone job runner in a separate process."""
    job_dir = get_job_dir(job_id)

    # Check if a job is already actively running
    for pid, proc in list(_ACTIVE_PROCESSES.items()):
        if proc.poll() is None:
            if pid == job_id:
                return jsonify({"status": "already_running"}), 200
            return jsonify({
                "error": "Server busy",
                "message": "Another generation job is currently running. Please wait for it to finish."
            }), 429
        else:
            _ACTIVE_PROCESSES.pop(pid, None)

    spec_file = job_dir / "job_spec.json"
    if not spec_file.exists():
        return jsonify({"error": "Job spec not found"}), 404

    with open(spec_file, "r", encoding="utf-8") as f:
        spec = json.load(f)

    if spec.get("blocking_errors"):
        return jsonify({
            "error": "Blocking errors exist",
            "message": "Cannot start generation while blocking preflight errors are present."
        }), 400

    # Optional worker count override
    req_workers = request.json.get("workers") if request.is_json else None
    if req_workers is not None:
        try:
            w_int = int(req_workers)
            if w_int > config.MAX_WORKERS:
                return jsonify({
                    "error": "Invalid workers",
                    "message": f"Worker count cannot exceed {config.MAX_WORKERS} on this machine."
                }), 400
            spec["workers"] = max(1, w_int)
            with open(spec_file, "w", encoding="utf-8") as f:
                json.dump(spec, f, indent=2, ensure_ascii=False)
        except ValueError:
            pass

    # Launch subprocess runner
    cmd = [
        sys.executable,
        str(config.BASE_DIR / "job_runner.py"),
        "--job-dir",
        str(job_dir),
    ]

    proc = subprocess.Popen(
        cmd,
        cwd=str(config.BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    _ACTIVE_PROCESSES[job_id] = proc

    return jsonify({"status": "started", "job_id": job_id})


@api_bp.route("/jobs/<job_id>", methods=["GET"])
@auth.login_required
def get_job_status(job_id: str):
    """Returns current real progress, stage, measured speed, and verification results."""
    job_dir = get_job_dir(job_id)
    status_file = job_dir / "status.json"

    if not status_file.exists():
        # Subprocess hasn't written status.json yet or just queued
        return jsonify({
            "job_id": job_id,
            "state": "queued",
            "stage": "queued",
            "message": "Job queued for generation...",
            "completed_batches": 0,
            "download_ready": False,
        })

    try:
        with open(status_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        return jsonify(data)
    except Exception:
        return jsonify({"state": "reading", "message": "Updating status..."})


@api_bp.route("/jobs/<job_id>/cancel", methods=["POST"])
@auth.login_required
def cancel_job(job_id: str):
    """Terminates the running job subprocess and LibreOffice children cleanly."""
    job_dir = get_job_dir(job_id)
    proc = _ACTIVE_PROCESSES.pop(job_id, None)

    if proc and proc.poll() is None:
        try:
            # On Windows, kill process tree
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
        except Exception:
            proc.kill()

    # Update status.json
    status_file = job_dir / "status.json"
    status_data = {
        "job_id": job_id,
        "state": "cancelled",
        "stage": "cancelled",
        "message": "Generation was cancelled by the supervisor.",
        "download_ready": False,
    }
    with open(status_file, "w", encoding="utf-8") as f:
        json.dump(status_data, f, indent=2, ensure_ascii=False)

    return jsonify({"status": "cancelled"})


@api_bp.route("/jobs/<job_id>/preview", methods=["GET"])
@auth.login_required
def get_job_preview(job_id: str):
    """Returns a PNG preview of Page 1 rendered via pypdfium2."""
    job_dir = get_job_dir(job_id)
    preview_png = job_dir / "preview_page1.png"

    if preview_png.exists():
        return send_file(preview_png, mimetype="image/png")

    # Find the first output PDF
    output_dir = job_dir / "output"
    pdfs = list(output_dir.glob("*.pdf"))
    if not pdfs:
        return jsonify({"error": "No generated PDF available for preview"}), 404

    pdf_path = pdfs[0]
    try:
        doc = pdfium.PdfDocument(str(pdf_path))
        if len(doc) == 0:
            return jsonify({"error": "Empty PDF document"}), 500
        page = doc[0]
        image = page.render(scale=1.5).to_pil()
        image.save(preview_png, format="PNG")
        return send_file(preview_png, mimetype="image/png")
    except Exception as e:
        return jsonify({"error": "Preview rendering failed", "message": str(e)}), 500


@api_bp.route("/jobs/<job_id>/download", methods=["GET"])
@auth.login_required
def download_main(job_id: str):
    """Streams the verified final output PDF or ZIP with Cache-Control: no-store."""
    job_dir = get_job_dir(job_id)
    status_file = job_dir / "status.json"
    if not status_file.exists():
        abort(404, description="Job status not found.")

    with open(status_file, "r", encoding="utf-8") as f:
        status = json.load(f)

    # Verification gate check: Never offer download unless verified
    if not status.get("verified") or not status.get("download_ready"):
        abort(403, description="Verification gates have not passed. Output download is disabled.")

    download_name = status.get("download_name")
    if not download_name:
        abort(404, description="Download file not specified.")

    target_file = job_dir / "output" / download_name
    if not target_file.exists():
        abort(404, description="Target download file does not exist.")

    response = send_file(
        target_file,
        as_attachment=True,
        download_name=download_name,
        conditional=True,
    )
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


@api_bp.route("/jobs/<job_id>/download/<batch_name>", methods=["GET"])
@auth.login_required
def download_batch(job_id: str, batch_name: str):
    """Streams an individual verified batch PDF."""
    job_dir = get_job_dir(job_id)
    status_file = job_dir / "status.json"
    if not status_file.exists():
        abort(404)

    with open(status_file, "r", encoding="utf-8") as f:
        status = json.load(f)

    if not status.get("verified"):
        abort(403, description="Verification gates have not passed.")

    # Strict filename sanitization
    safe_name = Path(batch_name).name
    if not safe_name.endswith(".pdf") or ".." in safe_name:
        abort(400, description="Invalid batch PDF filename.")

    target_file = job_dir / "output" / safe_name
    if not target_file.exists():
        abort(404, description="Batch PDF not found.")

    response = send_file(
        target_file,
        as_attachment=True,
        download_name=safe_name,
        conditional=True,
    )
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return response


@api_bp.route("/jobs/<job_id>/problems", methods=["GET"])
@auth.login_required
def download_problems(job_id: str):
    """Downloads the flagged preflight rows Excel workbook."""
    job_dir = get_job_dir(job_id)
    problems_file = job_dir / "Flagged_Rows_Report.xlsx"
    if not problems_file.exists():
        abort(404, description="No flagged rows report found for this job.")

    return send_file(
        problems_file,
        as_attachment=True,
        download_name="Flagged_Rows_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
