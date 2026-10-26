"""
app.py — Flask application factory for Traffic Notice Generator.
Includes startup self-checks, security headers, CSRF validation, and clean error handling.
"""
from __future__ import annotations

import glob
import hashlib
import os
import shutil
import time
from pathlib import Path

from flask import Flask, jsonify, render_template, request, session

import config
from routes import api_bp, auth_bp, main_bp


def verify_template_hash():
    """Rule 3: Refuse to start if template does not match pinned SHA-256 hash."""
    if not config.TEMPLATE_PATH.exists():
        raise RuntimeError(f"Approved template not found at {config.TEMPLATE_PATH}")

    h = hashlib.sha256()
    with open(config.TEMPLATE_PATH, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    actual_hash = h.hexdigest().lower()
    expected_hash = config.PINNED_TEMPLATE_HASH.lower()

    if actual_hash != expected_hash:
        raise RuntimeError(
            f"CRITICAL: Template SHA-256 hash mismatch!\n"
            f"Expected: {expected_hash}\n"
            f"Actual:   {actual_hash}\n"
            "Server startup aborted to prevent unauthorized or modified template printing."
        )


def verify_libreoffice():
    """Startup self-check: Ensure LibreOffice executable is available."""
    if not Path(config.SOFFICE_PATH).exists() and not shutil.which(config.SOFFICE_PATH):
        raise RuntimeError(
            f"LibreOffice executable not found at '{config.SOFFICE_PATH}'. "
            "Please install LibreOffice or configure SOFFICE_PATH environment variable."
        )


def cleanup_expired_jobs():
    """Startup & periodic retention cleaner: removes job directories older than retention minutes."""
    config.JOBS_DIR.mkdir(parents=True, exist_ok=True)
    now = time.time()
    cutoff_seconds = config.JOB_RETENTION_MINUTES * 60

    for job_dir in config.JOBS_DIR.iterdir():
        if job_dir.is_dir():
            try:
                mtime = job_dir.stat().st_mtime
                if now - mtime > cutoff_seconds:
                    shutil.rmtree(job_dir, ignore_errors=True)
            except Exception:
                pass


def create_app() -> Flask:
    # Rule 3: Enforce template integrity before accepting any requests
    verify_template_hash()
    verify_libreoffice()
    cleanup_expired_jobs()

    app = Flask(
        __name__,
        template_folder=str(config.BASE_DIR / "templates_html"),
        static_folder=str(config.BASE_DIR / "static"),
    )

    app.config["SECRET_KEY"] = config.SECRET_KEY
    app.config["SESSION_COOKIE_NAME"] = config.SESSION_COOKIE_NAME
    app.config["SESSION_COOKIE_HTTPONLY"] = config.SESSION_COOKIE_HTTPONLY
    app.config["SESSION_COOKIE_SAMESITE"] = config.SESSION_COOKIE_SAMESITE
    app.config["SESSION_COOKIE_SECURE"] = config.SESSION_COOKIE_SECURE
    app.config["MAX_CONTENT_LENGTH"] = config.MAX_UPLOAD_SIZE_BYTES

    # Register blueprints
    app.register_blueprint(main_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(api_bp)

    # Security Headers Middleware
    @app.after_request
    def set_security_headers(response):
        # Strict offline CSP: no external fonts, scripts, or CDNs allowed
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self'; "
            "style-src 'self' 'unsafe-inline'; "
            "font-src 'self'; "
            "img-src 'self' data:; "
            "connect-src 'self'; "
            "frame-ancestors 'none'; "
            "form-action 'self';"
        )
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        return response

    # Error Handlers (never exposing tracebacks to client)
    @app.errorhandler(400)
    def handle_400(e):
        if request.path.startswith("/api/"):
            return jsonify({"error": "Bad Request", "message": str(e.description)}), 400
        return render_template("404.html", error_title="Invalid Request", error_message=str(e.description)), 400

    @app.errorhandler(403)
    def handle_403(e):
        if request.path.startswith("/api/"):
            return jsonify({"error": "Forbidden", "message": str(e.description)}), 403
        return render_template("404.html", error_title="Access Forbidden", error_message=str(e.description)), 403

    @app.errorhandler(404)
    def handle_404(e):
        if request.path.startswith("/api/"):
            return jsonify({"error": "Not Found", "message": str(e.description)}), 404
        return render_template("404.html", error_title="Notice Not Found", error_message=str(e.description)), 404

    @app.errorhandler(413)
    def handle_413(e):
        msg = f"Uploaded file is too large. Maximum supported size is {config.MAX_UPLOAD_SIZE_BYTES // (1024*1024)} MB."
        if request.path.startswith("/api/"):
            return jsonify({"error": "File Too Large", "message": msg}), 413
        return render_template("404.html", error_title="File Too Large", error_message=msg), 413

    @app.errorhandler(500)
    def handle_500(e):
        if request.path.startswith("/api/"):
            return jsonify({"error": "Internal Error", "message": "An unexpected error occurred during processing."}), 500
        return render_template("404.html", error_title="System Error", error_message="An internal error occurred."), 500

    return app


if __name__ == "__main__":
    app = create_app()
    print(f"Starting Traffic Notice Generator server on http://{config.HOST}:{config.PORT}")
    app.run(host=config.HOST, port=config.PORT, debug=config.DEBUG)
