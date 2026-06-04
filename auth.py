"""
Authentication for CloudFuze Migrate Incentive Calculator.

Session-based login with bcrypt for password hashing.
"""

import re
from typing import Optional

try:
    import bcrypt
except ImportError:
    bcrypt = None  # type: ignore

from database import create_user, get_user_by_email, get_user_by_id


def hash_password(password: str) -> str:
    """Hash password using bcrypt. Falls back to plaintext if bcrypt not installed (dev only)."""
    if bcrypt:
        return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    return password  # Dev fallback - NOT for production


def verify_password(password: str, stored_hash: str) -> bool:
    """Verify password against stored hash."""
    if bcrypt:
        try:
            return bcrypt.checkpw(
                password.encode("utf-8"),
                stored_hash.encode("utf-8"),
            )
        except Exception:
            return False
    return password == stored_hash  # Dev fallback


def validate_email(email: str) -> bool:
    """Basic email format validation."""
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, email.strip()))


def authenticate(email: str, password: str) -> Optional[dict]:
    """
    Authenticate user by email and password.

    Returns:
        User dict (without password_hash) if valid, else None.
    """
    user = get_user_by_email(email.strip().lower())
    if not user:
        return None
    if not verify_password(password, user["password_hash"]):
        return None
    return {k: v for k, v in user.items() if k != "password_hash"}


def ensure_admin_user() -> None:
    """
    Create default admin user if no users exist.
    Email: admin@cloudfuze.com, Password: Admin@123 (change in production!)
    """
    from database import db_cursor

    with db_cursor(commit=False) as c:
        c.execute("SELECT COUNT(*) as cnt FROM users")
        if c.fetchone()["cnt"] > 0:
            return

    pw = hash_password("Admin@123")
    create_user(
        full_name="System Admin",
        email="admin@cloudfuze.com",
        password_hash=pw,
        role="ADMIN",
        team_id=None,
    )
