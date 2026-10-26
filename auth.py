"""
auth.py — Supervisor authentication, rate-limiting, and session management.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import time
from functools import wraps
from typing import Optional
from flask import session, request, redirect, url_for, jsonify, current_app

import config

# In-memory rate limiting: {ip_address: [(timestamp)]}
_FAILED_ATTEMPTS: dict[str, list[float]] = {}
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_SECONDS = 600  # 10 minutes


def check_rate_limit(ip: str) -> tuple[bool, int]:
    """Returns (is_allowed, seconds_remaining)."""
    now = time.time()
    attempts = _FAILED_ATTEMPTS.get(ip, [])
    # Filter attempts within lockout window
    valid_attempts = [t for t in attempts if now - t < LOCKOUT_SECONDS]
    _FAILED_ATTEMPTS[ip] = valid_attempts

    if len(valid_attempts) >= MAX_FAILED_ATTEMPTS:
        remaining = int(LOCKOUT_SECONDS - (now - valid_attempts[0]))
        return False, max(1, remaining)
    return True, 0


def record_failed_attempt(ip: str):
    now = time.time()
    attempts = _FAILED_ATTEMPTS.get(ip, [])
    attempts.append(now)
    _FAILED_ATTEMPTS[ip] = attempts


def clear_failed_attempts(ip: str):
    _FAILED_ATTEMPTS.pop(ip, None)


def verify_supervisor_password(provided_password: str) -> tuple[bool, Optional[str]]:
    """Validates the provided password against config.SUPERVISOR_PASSWORD using constant-time comparison."""
    expected = config.SUPERVISOR_PASSWORD
    if not expected:
        return False, "Supervisor password is not configured in the server environment (CHALLAN_SUPERVISOR_PASSWORD)."

    # Constant time comparison
    if hmac.compare_digest(provided_password.strip(), expected.strip()):
        return True, None
    return False, "Incorrect password. Please try again."


def is_authenticated() -> bool:
    """Checks whether the current session is authenticated as supervisor."""
    # If no supervisor password is set in environment, auth can be disabled or prompt setup
    if not config.SUPERVISOR_PASSWORD:
        return False
    return session.get("is_supervisor") is True


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_authenticated():
            if request.path.startswith("/api/"):
                return jsonify({"error": "Unauthorized", "message": "Please log in to continue."}), 401
            return redirect(url_for("auth.login_page", next=request.path))
        return f(*args, **kwargs)
    return decorated
