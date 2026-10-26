"""
config.py — Central configuration for Traffic Notice Generator.
All paths, ports, folders, limits, and settings are defined here.
No machine-specific values are hard-coded elsewhere in the application.
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path

# Base Paths
BASE_DIR = Path(__file__).resolve().parent
JOBS_DIR = Path(os.environ.get("CHALLAN_JOBS_DIR", str(BASE_DIR / "jobs")))
TEMPLATE_PATH = Path(os.environ.get("CHALLAN_TEMPLATE_PATH", str(BASE_DIR / "templates" / "Formate_with_Sign_SectionSplit.docx")))
STATIC_DIR = BASE_DIR / "static"
FONTS_DIR = STATIC_DIR / "fonts"

# Pinned Template SHA-256 Hash (Rule 3)
PINNED_TEMPLATE_HASH = "2b3c563af36b29526bcf467f5a11a9fdff6361740f9990002e6ea7e9e4e8ff8b"

# Network & Server Settings
HOST = os.environ.get("CHALLAN_HOST", "127.0.0.1")  # Default to localhost only; LAN is explicit opt-in
PORT = int(os.environ.get("CHALLAN_PORT", "5000"))
DEBUG = os.environ.get("CHALLAN_DEBUG", "0") == "1"

# Security & Authentication
# Supervisor password must be provided via CHALLAN_SUPERVISOR_PASSWORD
SUPERVISOR_PASSWORD = os.environ.get("CHALLAN_SUPERVISOR_PASSWORD", "")
SECRET_KEY = os.environ.get("CHALLAN_SECRET_KEY", "")
if not SECRET_KEY:
    # Use a persistent local secret key file so restart doesn't log out active sessions
    secret_file = BASE_DIR / ".session_secret"
    if secret_file.exists():
        SECRET_KEY = secret_file.read_text(encoding="utf-8").strip()
    else:
        import secrets
        SECRET_KEY = secrets.token_hex(32)
        try:
            secret_file.write_text(SECRET_KEY, encoding="utf-8")
        except Exception:
            pass

SESSION_COOKIE_NAME = "challan_session"
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = "Strict"
SESSION_COOKIE_SECURE = False  # Set to True in HTTPS deployments

# Generation & Batching Rules
DEFAULT_BATCH_SIZE = 1000
MAX_WORKERS = 4  # Hard ceiling on laptop (6 is no faster, 8 causes LO timeouts)
DEFAULT_WORKERS = int(os.environ.get("CHALLAN_WORKERS", "4"))
if DEFAULT_WORKERS > MAX_WORKERS:
    DEFAULT_WORKERS = MAX_WORKERS

# Retention & Cleanup
JOB_RETENTION_MINUTES = int(os.environ.get("CHALLAN_JOB_RETENTION_MINUTES", "60"))
MAX_UPLOAD_SIZE_BYTES = int(os.environ.get("CHALLAN_MAX_UPLOAD_SIZE_MB", "50")) * 1024 * 1024

# LibreOffice Auto-Detection
STANDARD_LO_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/usr/bin/libreoffice",
]


def detect_soffice_path() -> str:
    env_path = os.environ.get("SOFFICE_PATH")
    if env_path and Path(env_path).exists():
        return env_path
    shutil_which = shutil.which("soffice") or shutil.which("libreoffice")
    if shutil_which:
        return shutil_which
    for cand in STANDARD_LO_PATHS:
        if Path(cand).exists():
            return cand
    return "soffice"


SOFFICE_PATH = detect_soffice_path()
