"""
Database access layer for CloudFuze Migrate Incentive Calculator.

Manages all interactions with PostgreSQL: schema, teams, users, deals,
incentive slabs, rep/team incentives, uploads, and audit logs.
"""

import os
from contextlib import contextmanager
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

import psycopg2
from psycopg2.extras import RealDictCursor


DB_HOST_ENV = "COMP_DB_HOST"
DB_PORT_ENV = "COMP_DB_PORT"
DB_NAME_ENV = "COMP_DB_NAME"
DB_USER_ENV = "COMP_DB_USER"
DB_PASSWORD_ENV = "COMP_DB_PASSWORD"


def _is_duplicate_column_error(e: Exception) -> bool:
    """True if error indicates column already exists (PostgreSQL or MySQL style)."""
    msg = str(e).lower()
    return "duplicate column" in msg or "already exists" in msg


def get_db_config() -> Dict[str, Any]:
    """Build PostgreSQL connection config from environment variables."""
    host = os.getenv(DB_HOST_ENV, "localhost")
    port = int(os.getenv(DB_PORT_ENV, "5432"))
    database = os.getenv(DB_NAME_ENV)
    user = os.getenv(DB_USER_ENV)
    password = os.getenv(DB_PASSWORD_ENV)

    if not all([database, user, password]):
        raise RuntimeError(
            "Database configuration incomplete. "
            "Set COMP_DB_NAME, COMP_DB_USER, COMP_DB_PASSWORD."
        )

    return {
        "host": host,
        "port": port,
        "database": database,
        "user": user,
        "password": password,
    }


def get_connection():
    """Create and return a new PostgreSQL connection."""
    return psycopg2.connect(**get_db_config())


@contextmanager
def db_cursor(commit: bool = False):
    """Context manager for cursor with optional commit. Cursor returns dict-like rows."""
    conn = get_connection()
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    try:
        yield cursor
        if commit:
            conn.commit()
    finally:
        cursor.close()
        conn.close()


def _migrate_user_compensation_group_constraint() -> None:
    """Ensure ``compensation_group`` allows SMB_CHITRADIP (replace legacy CHECK if present)."""
    try:
        with db_cursor(commit=False) as cursor:
            cursor.execute(
                """
                SELECT c.conname
                FROM pg_constraint c
                JOIN pg_class t ON c.conrelid = t.oid
                WHERE t.relname = 'users' AND c.contype = 'c'
                  AND pg_get_constraintdef(c.oid) LIKE '%compensation_group%'
                """
            )
            names = [r["conname"] for r in (cursor.fetchall() or [])]
        for name in names:
            try:
                with db_cursor(commit=True) as cursor:
                    cursor.execute(f'ALTER TABLE users DROP CONSTRAINT "{name}"')
            except psycopg2.Error:
                pass
    except Exception:
        pass
    try:
        with db_cursor(commit=True) as cursor:
            cursor.execute(
                """ALTER TABLE users ADD CONSTRAINT users_compensation_group_chk
                CHECK (compensation_group IS NULL OR compensation_group IN ('SMB_A', 'SMB_B', 'SMB_CHITRADIP'))"""
            )
    except psycopg2.Error as e:
        if "already exists" not in str(e).lower():
            pass


def initialize_schema() -> None:
    """Create all tables required for CloudFuze Migrate Incentive Calculator (PostgreSQL)."""
    statements = [
        """
        CREATE TABLE IF NOT EXISTS teams (
            team_id SERIAL PRIMARY KEY,
            team_name VARCHAR(100) NOT NULL UNIQUE,
            team_goal DECIMAL(18, 2) NULL,
            is_active BOOLEAN DEFAULT TRUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS users (
            user_id SERIAL PRIMARY KEY,
            full_name VARCHAR(255) NOT NULL,
            email VARCHAR(255) NOT NULL UNIQUE,
            password_hash VARCHAR(255) NOT NULL,
            role VARCHAR(50) NOT NULL DEFAULT 'SALES_REP' CHECK (role IN ('ADMIN', 'SALES_REP', 'SALES_MANAGER')),
            team_id INT NULL,
            is_active BOOLEAN DEFAULT TRUE,
            eligible_for_compensation BOOLEAN NOT NULL DEFAULT TRUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (team_id) REFERENCES teams(team_id) ON DELETE SET NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS excel_uploads (
            upload_id SERIAL PRIMARY KEY,
            uploaded_by INT NOT NULL,
            file_name VARCHAR(255) NOT NULL,
            upload_status VARCHAR(50) NOT NULL DEFAULT 'DRAFT' CHECK (upload_status IN ('DRAFT', 'VALIDATED', 'FINALIZED')),
            records_processed INT DEFAULT 0,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            finalized_at TIMESTAMP NULL,
            FOREIGN KEY (uploaded_by) REFERENCES users(user_id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS incentive_slabs (
            slab_id SERIAL PRIMARY KEY,
            min_revenue DECIMAL(18, 2) NOT NULL,
            max_revenue DECIMAL(18, 2) NULL,
            incentive_percentage DECIMAL(5, 2) NOT NULL,
            is_active BOOLEAN DEFAULT TRUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS deals (
            deal_id SERIAL PRIMARY KEY,
            deal_name VARCHAR(255) NOT NULL,
            deal_owner_id INT NOT NULL,
            team_id INT NOT NULL,
            amount DECIMAL(18, 2) NOT NULL,
            paid_amount DECIMAL(18, 2) NOT NULL DEFAULT 0,
            payment_status VARCHAR(50) NOT NULL DEFAULT 'UNPAID' CHECK (payment_status IN ('PAID', 'UNPAID', 'PARTIALLY_PAID')),
            upload_id INT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (deal_owner_id) REFERENCES users(user_id),
            FOREIGN KEY (team_id) REFERENCES teams(team_id),
            FOREIGN KEY (upload_id) REFERENCES excel_uploads(upload_id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS rep_incentives (
            rep_incentive_id SERIAL PRIMARY KEY,
            user_id INT NOT NULL,
            team_id INT NOT NULL,
            total_deals_closed INT NOT NULL,
            total_revenue DECIMAL(18, 2) NOT NULL,
            total_paid_amount DECIMAL(18, 2) NOT NULL DEFAULT 0,
            slab_id INT NOT NULL,
            incentive_percentage DECIMAL(5, 2) NOT NULL,
            incentive_amount DECIMAL(18, 2) NOT NULL,
            payment_status VARCHAR(50) DEFAULT 'UNPAID' CHECK (payment_status IN ('PAID', 'UNPAID', 'PARTIALLY_PAID', 'N/A')),
            calculation_period VARCHAR(50) NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(user_id),
            FOREIGN KEY (team_id) REFERENCES teams(team_id),
            FOREIGN KEY (slab_id) REFERENCES incentive_slabs(slab_id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS team_incentives (
            team_incentive_id SERIAL PRIMARY KEY,
            team_id INT NOT NULL,
            team_lead_id INT NOT NULL,
            total_team_deals INT NOT NULL,
            total_team_revenue DECIMAL(18, 2) NOT NULL,
            total_paid_amount DECIMAL(18, 2) NOT NULL DEFAULT 0,
            slab_id INT NOT NULL,
            incentive_percentage DECIMAL(5, 2) NOT NULL,
            incentive_amount DECIMAL(18, 2) NOT NULL,
            payment_status VARCHAR(50) DEFAULT 'UNPAID' CHECK (payment_status IN ('PAID', 'UNPAID', 'PARTIALLY_PAID')),
            calculation_period VARCHAR(50) NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (team_id) REFERENCES teams(team_id),
            FOREIGN KEY (team_lead_id) REFERENCES users(user_id),
            FOREIGN KEY (slab_id) REFERENCES incentive_slabs(slab_id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS audit_logs (
            audit_id SERIAL PRIMARY KEY,
            action_type VARCHAR(100) NOT NULL,
            performed_by INT NULL,
            entity_type VARCHAR(100) NOT NULL,
            entity_id VARCHAR(100) NULL,
            details TEXT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (performed_by) REFERENCES users(user_id)
        )
        """,
    ]

    # Run CREATE TABLEs in one transaction (PostgreSQL: one failure aborts the whole block)
    with db_cursor(commit=True) as cursor:
        for sql in statements:
            cursor.execute(sql)

    # Run each ALTER in its own transaction so a duplicate-column error doesn't abort the rest
    def _run_alter(alter_sql: str) -> None:
        try:
            with db_cursor(commit=True) as cursor:
                cursor.execute(alter_sql)
        except psycopg2.Error as e:
            if not _is_duplicate_column_error(e) and "already exists" not in str(e).lower():
                raise

    _run_alter("ALTER TABLE deals ADD COLUMN paid_amount DECIMAL(18, 2) NOT NULL DEFAULT 0")
    for tbl, col in [("rep_incentives", "total_paid_amount"), ("team_incentives", "total_paid_amount")]:
        _run_alter(f"ALTER TABLE {tbl} ADD COLUMN {col} DECIMAL(18, 2) NOT NULL DEFAULT 0")
    _run_alter("ALTER TABLE teams ADD COLUMN team_goal DECIMAL(18, 2) NULL")
    _run_alter("ALTER TABLE users ADD COLUMN eligible_for_compensation BOOLEAN NOT NULL DEFAULT TRUE")
    _run_alter("ALTER TABLE users ADD COLUMN compensation_group VARCHAR(20) NULL")
    _run_alter(
        "ALTER TABLE incentive_slabs ADD COLUMN slab_set VARCHAR(32) NOT NULL DEFAULT 'DEFAULT'"
    )
    try:
        with db_cursor(commit=True) as cursor:
            cursor.execute(
                "ALTER TABLE rep_incentives ADD CONSTRAINT rep_incentives_payment_status_check "
                "CHECK (payment_status IN ('PAID', 'UNPAID', 'PARTIALLY_PAID', 'N/A'))"
            )
    except psycopg2.Error as e:
        if "already exists" not in str(e).lower():
            raise
    for col_def in [
        "ALTER TABLE rep_incentives ADD COLUMN close_date DATE NULL",
        "ALTER TABLE rep_incentives ADD COLUMN incentive_eligibility VARCHAR(50) NULL",
        "ALTER TABLE deals ADD COLUMN close_date DATE NULL",
        "ALTER TABLE deals ADD COLUMN incentive_eligibility VARCHAR(50) NULL",
    ]:
        _run_alter(col_def)

    _run_alter("ALTER TABLE deals ADD COLUMN license_resale_exclusion BOOLEAN NOT NULL DEFAULT FALSE")
    _run_alter("ALTER TABLE rep_incentives ADD COLUMN quota DECIMAL(18, 2) NULL")
    _run_alter("ALTER TABLE users ADD COLUMN hubspot_quota_usd DECIMAL(18, 2) NULL")
    _run_alter("ALTER TABLE users ADD COLUMN hubspot_quota_period VARCHAR(32) NULL")

    _migrate_user_compensation_group_constraint()

    with db_cursor(commit=True) as cursor:
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS outbound_meetings (
                outbound_id SERIAL PRIMARY KEY,
                user_id INT NOT NULL REFERENCES users(user_id) ON DELETE CASCADE,
                region VARCHAR(32) NOT NULL,
                meeting_date DATE NOT NULL,
                incentive_amount DECIMAL(18, 2) NOT NULL DEFAULT 0,
                notes TEXT NULL,
                created_by INT NULL REFERENCES users(user_id) ON DELETE SET NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

    _ensure_deals_payment_status_includes_refunds()

    _seed_default_data()
    ensure_outbound_team()
    apply_commission_policy_team_targets()
    ensure_slab_sets_from_policy()


def _ensure_deals_payment_status_includes_refunds() -> None:
    """Allow REFUNDED / PARTIALLY_REFUNDED on deals (HubSpot enumeration)."""
    try:
        with db_cursor(commit=True) as cursor:
            cursor.execute("ALTER TABLE deals DROP CONSTRAINT IF EXISTS deals_payment_status_check")
            cursor.execute(
                """ALTER TABLE deals ADD CONSTRAINT deals_payment_status_check
                   CHECK (payment_status IN (
                       'PAID', 'UNPAID', 'PARTIALLY_PAID', 'REFUNDED', 'PARTIALLY_REFUNDED'
                   ))"""
            )
    except psycopg2.Error:
        pass


def apply_commission_policy_team_targets() -> None:
    """Apply quarterly team targets from ``commission_policy`` / policy JSON (incl. SMB A+B sum)."""
    from commission_policy import TEAM_QUARTERLY_TARGETS_USD, smb_team_goal_from_subgroups

    smb_split = smb_team_goal_from_subgroups()

    with db_cursor(commit=True) as cursor:
        for team_name, goal in TEAM_QUARTERLY_TARGETS_USD.items():
            if team_name == "SMB" and smb_split is not None:
                goal = smb_split
            cursor.execute(
                "UPDATE teams SET team_goal = %s WHERE team_name = %s",
                (goal, team_name),
            )


def ensure_outbound_team() -> None:
    """Idempotent: ensure the Outbound team row exists (User Management Team dropdown, etc.)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "INSERT INTO teams (team_name) VALUES ('Outbound') ON CONFLICT (team_name) DO NOTHING"
        )


def _seed_default_data() -> None:
    """Seed teams, incentive slabs, and default admin if tables are empty."""
    with db_cursor(commit=False) as cursor:
        cursor.execute("SELECT COUNT(*) as cnt FROM teams")
        if cursor.fetchone()["cnt"] > 0:
            return

    # Seed teams (PostgreSQL: ON CONFLICT DO NOTHING instead of INSERT IGNORE)
    teams = [("SMB",), ("Enterprise",), ("Account Management",), ("Outbound",)]
    with db_cursor(commit=True) as cursor:
        for t in teams:
            cursor.execute(
                "INSERT INTO teams (team_name) VALUES (%s) ON CONFLICT (team_name) DO NOTHING", t
            )

    from commission_policy import REP_SLAB_ROWS_FOR_DB

    slabs = REP_SLAB_ROWS_FOR_DB
    with db_cursor(commit=True) as cursor:
        for min_rev, max_rev, pct in slabs:
            cursor.execute(
                """INSERT INTO incentive_slabs (min_revenue, max_revenue, incentive_percentage, slab_set)
                   VALUES (%s, %s, %s, 'DEFAULT')""",
                (min_rev, max_rev, pct),
            )


def ensure_slab_sets_from_policy() -> None:
    """
    Insert missing incentive slab rows for each policy slab set (DEFAULT, SMB_A, SMB_B).
    Safe to run on every startup; only inserts when a slab_set has zero rows.
    """
    from commission_policy import all_slab_sets_for_db

    sets = all_slab_sets_for_db()
    with db_cursor(commit=True) as cursor:
        for slab_set, rows in sets.items():
            cursor.execute(
                "SELECT COUNT(*) AS c FROM incentive_slabs WHERE slab_set = %s",
                (slab_set,),
            )
            cnt = cursor.fetchone()
            n = int(cnt["c"] if cnt else 0)
            if n > 0:
                continue
            for min_rev, max_rev, pct in rows:
                cursor.execute(
                    """INSERT INTO incentive_slabs (min_revenue, max_revenue, incentive_percentage, slab_set)
                       VALUES (%s, %s, %s, %s)""",
                    (min_rev, max_rev, pct, slab_set),
                )


# --- Teams ---
def get_all_teams(active_only: bool = True) -> List[Dict[str, Any]]:
    """Fetch all teams (includes team_goal)."""
    sql = "SELECT team_id, team_name, team_goal, is_active FROM teams"
    if active_only:
        sql += " WHERE is_active = TRUE"
    sql += " ORDER BY team_name"
    with db_cursor(commit=False) as cursor:
        cursor.execute(sql)
        return cursor.fetchall() or []


def update_team_goal(team_id: int, team_goal: float | None) -> None:
    """Set or clear team goal for team incentives (achievement-based commission)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "UPDATE teams SET team_goal = %s WHERE team_id = %s",
            (team_goal if team_goal is not None and team_goal >= 0 else None, team_id),
        )


def get_team_by_id(team_id: int) -> Optional[Dict[str, Any]]:
    """Get team by ID (includes team_goal)."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            "SELECT team_id, team_name, team_goal, is_active FROM teams WHERE team_id = %s",
            (team_id,),
        )
        return cursor.fetchone()


def get_team_by_name(name: str) -> Optional[Dict[str, Any]]:
    """Get team by name."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            "SELECT team_id, team_name FROM teams WHERE team_name = %s",
            (name.strip(),),
        )
        return cursor.fetchone()


# --- Users ---
def get_user_by_email(email: str) -> Optional[Dict[str, Any]]:
    """Get user by email."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT user_id, full_name, email, password_hash, role, team_id, is_active
               FROM users WHERE email = %s AND is_active = TRUE""",
            (email.strip().lower(),),
        )
        return cursor.fetchone()


def get_user_by_id(user_id: int) -> Optional[Dict[str, Any]]:
    """Get user by ID."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT u.user_id, u.full_name, u.email, u.role, u.team_id, u.is_active,
                      u.compensation_group, t.team_name,
                      u.hubspot_quota_usd, u.hubspot_quota_period
               FROM users u
               LEFT JOIN teams t ON u.team_id = t.team_id
               WHERE u.user_id = %s""",
            (user_id,),
        )
        return cursor.fetchone()


def create_user(
    full_name: str,
    email: str,
    password_hash: str,
    role: str,
    team_id: Optional[int] = None,
    compensation_group: Optional[str] = None,
) -> int:
    """Create a new user. Returns user_id. ``compensation_group``: SMB_A / SMB_B for SMB reps only."""
    sql = """
    INSERT INTO users (full_name, email, password_hash, role, team_id, compensation_group)
    VALUES (%s, %s, %s, %s, %s, %s)
    RETURNING user_id
    """
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            sql,
            (
                full_name.strip(),
                email.strip().lower(),
                password_hash,
                role,
                team_id,
                compensation_group,
            ),
        )
        row = cursor.fetchone()
        return int(row["user_id"])


def update_user_password(user_id: int, password_hash: str) -> None:
    """Update a user's password."""
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "UPDATE users SET password_hash = %s WHERE user_id = %s",
            (password_hash, user_id),
        )


def update_user_profile(
    user_id: int,
    full_name: str,
    email: str,
    role: str,
    team_id: Optional[int] = None,
    compensation_group: Optional[str] = None,
    password_hash: Optional[str] = None,
) -> None:
    """
    Update user fields. Pass ``password_hash`` only when changing password.

    Raises:
        ValueError: if email is already used by another active user.
    """
    email_norm = email.strip().lower()
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            "SELECT user_id FROM users WHERE email = %s AND is_active = TRUE",
            (email_norm,),
        )
        row = cursor.fetchone()
        if row and int(row["user_id"]) != user_id:
            raise ValueError("That email is already in use by another user.")

    fields = [
        "full_name = %s",
        "email = %s",
        "role = %s",
        "team_id = %s",
        "compensation_group = %s",
    ]
    params: List[Any] = [
        full_name.strip(),
        email_norm,
        role,
        team_id,
        compensation_group,
    ]
    if password_hash is not None:
        fields.append("password_hash = %s")
        params.append(password_hash)
    params.append(user_id)
    sql = f"UPDATE users SET {', '.join(fields)} WHERE user_id = %s"
    with db_cursor(commit=True) as cursor:
        cursor.execute(sql, tuple(params))


def update_user_compensation_group(user_id: int, compensation_group: Optional[str]) -> None:
    """Set SMB subgroup (SMB_A / SMB_B / SMB_CHITRADIP) or clear."""
    g = compensation_group if compensation_group in (None, "SMB_A", "SMB_B", "SMB_CHITRADIP") else None
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "UPDATE users SET compensation_group = %s WHERE user_id = %s",
            (g, user_id),
        )


def update_user_hubspot_quota(
    user_id: int,
    hubspot_quota_usd: Optional[float],
    hubspot_quota_period: Optional[str],
) -> None:
    """Store individual sales target synced from HubSpot Goals (Forecast) for the given period label (e.g. 2026-Q1)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "UPDATE users SET hubspot_quota_usd = %s, hubspot_quota_period = %s WHERE user_id = %s",
            (hubspot_quota_usd, hubspot_quota_period, user_id),
        )


def bulk_update_users_hubspot_quotas(
    rows: List[Tuple[int, Optional[float], Optional[str]]],
) -> None:
    """Set HubSpot-synced quota for many users: list of (user_id, hubspot_quota_usd, hubspot_quota_period)."""
    if not rows:
        return
    with db_cursor(commit=True) as cursor:
        for uid, amt, per in rows:
            cursor.execute(
                "UPDATE users SET hubspot_quota_usd = %s, hubspot_quota_period = %s WHERE user_id = %s",
                (amt, per, uid),
            )


def update_user_eligible_for_compensation(user_id: int, eligible: bool) -> None:
    """Set whether user is eligible for compensation (quota achieved). Rep-only; affects rep incentive generation."""
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "UPDATE users SET eligible_for_compensation = %s WHERE user_id = %s",
            (eligible, user_id),
        )


def get_user_ids_not_eligible_for_compensation() -> List[int]:
    """Return user_ids where eligible_for_compensation = FALSE (not eligible; quota not achieved). For rep incentive logic only."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            "SELECT user_id FROM users WHERE is_active = TRUE AND COALESCE(eligible_for_compensation, TRUE) = FALSE"
        )
        return [row["user_id"] for row in (cursor.fetchall() or [])]


def get_users_for_team(team_id: int) -> List[Dict[str, Any]]:
    """Get all users in a team."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT user_id, full_name, email, role FROM users
               WHERE team_id = %s AND is_active = TRUE""",
            (team_id,),
        )
        return cursor.fetchall() or []


def get_all_users_with_teams(active_only: bool = True) -> List[Dict[str, Any]]:
    """Get all users with team info (for validation during upload). Newest at end."""
    where_active = "WHERE u.is_active = TRUE" if active_only else ""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            f"""SELECT u.user_id, u.full_name, u.email, u.role, u.team_id, t.team_name,
                      u.is_active,
                      COALESCE(u.eligible_for_compensation, TRUE) AS eligible_for_compensation,
                      u.compensation_group,
                      u.hubspot_quota_usd, u.hubspot_quota_period
               FROM users u
               LEFT JOIN teams t ON u.team_id = t.team_id
               {where_active}
               ORDER BY u.user_id ASC"""
        )
        return cursor.fetchall() or []


# --- Outbound meetings (NAM / Western Europe per Q1 2026 policy) ---
def insert_outbound_meeting(
    user_id: int,
    region: str,
    meeting_date: date,
    incentive_amount: float,
    notes: Optional[str],
    created_by: int,
) -> int:
    sql = """
    INSERT INTO outbound_meetings (user_id, region, meeting_date, incentive_amount, notes, created_by)
    VALUES (%s, %s, %s, %s, %s, %s)
    RETURNING outbound_id
    """
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            sql,
            (user_id, region, meeting_date, incentive_amount, notes if notes else None, created_by),
        )
        row = cursor.fetchone()
        return int(row["outbound_id"])


def get_all_outbound_meetings() -> List[Dict[str, Any]]:
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """
            SELECT o.outbound_id, o.user_id, o.region, o.meeting_date, o.incentive_amount,
                   o.notes, o.created_at, u.full_name AS rep_name, u.email AS rep_email
            FROM outbound_meetings o
            JOIN users u ON o.user_id = u.user_id
            ORDER BY o.meeting_date DESC, o.outbound_id DESC
            """
        )
        return cursor.fetchall() or []


def delete_outbound_meeting(outbound_id: int) -> None:
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM outbound_meetings WHERE outbound_id = %s", (outbound_id,))


# --- Incentive Slabs ---
def get_active_slabs(slab_set: str = "DEFAULT") -> List[Dict[str, Any]]:
    """
    Active incentive slabs for a named set (DEFAULT = Enterprise / AM / legacy SMB).

    SMB Group A / B use slab_set SMB_A and SMB_B when defined in policy/commission_policy.json.
    """
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT slab_id, min_revenue, max_revenue, incentive_percentage, slab_set
               FROM incentive_slabs WHERE is_active = TRUE AND slab_set = %s
               ORDER BY min_revenue DESC""",
            (slab_set,),
        )
        rows = cursor.fetchall() or []
    if not rows and slab_set != "DEFAULT":
        return get_active_slabs("DEFAULT")
    return rows


# --- Excel Uploads ---
def create_upload(
    uploaded_by: int,
    file_name: str,
    records_processed: int = 0,
) -> int:
    """Create a new upload record. Returns upload_id."""
    sql = """
    INSERT INTO excel_uploads (uploaded_by, file_name, upload_status, records_processed)
    VALUES (%s, %s, 'DRAFT', %s)
    RETURNING upload_id
    """
    with db_cursor(commit=True) as cursor:
        cursor.execute(sql, (uploaded_by, file_name, records_processed))
        row = cursor.fetchone()
        return int(row["upload_id"])


def update_upload_status(
    upload_id: int,
    status: str,
    records_processed: Optional[int] = None,
) -> None:
    """Update upload status."""
    if status == "FINALIZED":
        sql = """UPDATE excel_uploads SET upload_status = %s, records_processed = COALESCE(%s, records_processed),
                 finalized_at = CURRENT_TIMESTAMP WHERE upload_id = %s"""
        with db_cursor(commit=True) as cursor:
            cursor.execute(sql, (status, records_processed, upload_id))
    else:
        sql = "UPDATE excel_uploads SET upload_status = %s"
        params = [status]
        if records_processed is not None:
            sql += ", records_processed = %s"
            params.append(records_processed)
        sql += " WHERE upload_id = %s"
        params.append(upload_id)
        with db_cursor(commit=True) as cursor:
            cursor.execute(sql, tuple(params))


def get_upload_by_id(upload_id: int) -> Optional[Dict[str, Any]]:
    """Get upload by ID."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT u.*, up.full_name as uploaded_by_name
               FROM excel_uploads u
               LEFT JOIN users up ON u.uploaded_by = up.user_id
               WHERE u.upload_id = %s""",
            (upload_id,),
        )
        return cursor.fetchone()


def get_uploads_for_user(user_id: int, admin: bool = False) -> List[Dict[str, Any]]:
    """Get uploads. Admins see all, others see only their own."""
    if admin:
        sql = """SELECT u.*, up.full_name as uploaded_by_name
                 FROM excel_uploads u
                 LEFT JOIN users up ON u.uploaded_by = up.user_id
                 ORDER BY u.uploaded_at DESC"""
        with db_cursor(commit=False) as cursor:
            cursor.execute(sql)
            return cursor.fetchall() or []
    sql = """SELECT u.*, up.full_name as uploaded_by_name
             FROM excel_uploads u
             LEFT JOIN users up ON u.uploaded_by = up.user_id
             WHERE u.uploaded_by = %s
             ORDER BY u.uploaded_at DESC"""
    with db_cursor(commit=False) as cursor:
        cursor.execute(sql, (user_id,))
        return cursor.fetchall() or []


# --- Deals ---
def _ensure_deals_extra_columns() -> None:
    """Ensure close_date, incentive_eligibility, license_resale_exclusion exist on deals."""
    for col_sql in [
        "ALTER TABLE deals ADD COLUMN close_date DATE NULL",
        "ALTER TABLE deals ADD COLUMN incentive_eligibility VARCHAR(50) NULL",
        "ALTER TABLE deals ADD COLUMN license_resale_exclusion BOOLEAN NOT NULL DEFAULT FALSE",
    ]:
        try:
            with db_cursor(commit=True) as cursor:
                cursor.execute(col_sql)
        except psycopg2.Error as e:
            if not _is_duplicate_column_error(e):
                raise


def insert_deals(deals: List[Tuple[Any, ...]]) -> None:
    """Bulk insert deals. Each tuple: (...9 legacy...) or 10 cols with license_resale_exclusion last."""
    if not deals:
        return
    _ensure_deals_extra_columns()
    rows: List[Tuple[Any, ...]] = []
    for t in deals:
        if len(t) >= 10:
            rows.append(tuple(t[:10]))
        elif len(t) == 9:
            rows.append(tuple(list(t) + [False]))
        else:
            raise ValueError(f"insert_deals: expected 9 or 10 values, got {len(t)}")
    sql = """
    INSERT INTO deals (deal_name, deal_owner_id, team_id, amount, payment_status, upload_id, paid_amount, close_date, incentive_eligibility, license_resale_exclusion)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    with db_cursor(commit=True) as cursor:
        cursor.executemany(sql, rows)


def get_deals_by_upload(upload_id: int) -> List[Dict[str, Any]]:
    """Get all deals for an upload."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT d.*, u.full_name as deal_owner_name, u.email AS deal_owner_email, t.team_name,
                      u.compensation_group AS owner_compensation_group,
                      u.hubspot_quota_usd AS owner_hubspot_quota_usd,
                      u.hubspot_quota_period AS owner_hubspot_quota_period
               FROM deals d
               LEFT JOIN users u ON d.deal_owner_id = u.user_id
               LEFT JOIN teams t ON d.team_id = t.team_id
               WHERE d.upload_id = %s
               ORDER BY d.deal_id""",
            (upload_id,),
        )
        return cursor.fetchall() or []


def sync_deal_paid_amounts_from_status(upload_id: int) -> None:
    """
    Set paid_amount from payment_status and amount (e.g. PAID → full amount when paid was still 0).

    Call before finalizing an upload so incentives and the deals table stay consistent.
    """
    from excel_service import effective_paid_amount_from_status

    deals = get_deals_by_upload(upload_id)
    if not deals:
        return
    with db_cursor(commit=True) as cursor:
        for d in deals:
            new_paid = effective_paid_amount_from_status(d)
            old_paid = float(d.get("paid_amount") or 0)
            if new_paid != old_paid:
                cursor.execute(
                    "UPDATE deals SET paid_amount = %s WHERE deal_id = %s",
                    (new_paid, d["deal_id"]),
                )


def get_deals_from_finalized_uploads() -> List[Dict[str, Any]]:
    """Get all deals from finalized uploads (for rep deal names in graphical view)."""
    with db_cursor(commit=False) as cursor:
        cursor.execute(
            """SELECT d.deal_id, d.deal_name, d.deal_owner_id, d.amount, d.paid_amount, d.payment_status, d.close_date,
                      u.full_name as deal_owner_name
               FROM deals d
               JOIN excel_uploads e ON d.upload_id = e.upload_id
               LEFT JOIN users u ON d.deal_owner_id = u.user_id
               WHERE e.upload_status = 'FINALIZED'
               ORDER BY u.full_name, d.deal_name""",
        )
        return cursor.fetchall() or []


def delete_deals_by_upload(upload_id: int) -> None:
    """Delete all deals for an upload (used when re-uploading or editing draft)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM deals WHERE upload_id = %s", (upload_id,))


def delete_upload(upload_id: int) -> None:
    """Delete an upload and all its deals. Incentives already generated (if finalized) are not removed."""
    delete_deals_by_upload(upload_id)
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM excel_uploads WHERE upload_id = %s", (upload_id,))


def delete_all_uploads() -> int:
    """Delete all deals then all uploads. Returns number of uploads deleted."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM deals")
        cursor.execute("DELETE FROM excel_uploads")
        return cursor.rowcount if cursor.rowcount is not None else 0


def update_deal(
    deal_id: int,
    deal_name: Optional[str] = None,
    deal_owner_id: Optional[int] = None,
    team_id: Optional[int] = None,
    amount: Optional[float] = None,
    payment_status: Optional[str] = None,
) -> None:
    """Update a single deal (for draft edits)."""
    updates = []
    params = []
    for col, val in [
        ("deal_name", deal_name),
        ("deal_owner_id", deal_owner_id),
        ("team_id", team_id),
        ("amount", amount),
        ("payment_status", payment_status),
    ]:
        if val is not None:
            updates.append(f"{col} = %s")
            params.append(val)
    if not updates:
        return
    params.append(deal_id)
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            f"UPDATE deals SET {', '.join(updates)} WHERE deal_id = %s",
            tuple(params),
        )


# --- Rep Incentives ---
def _ensure_rep_incentives_extra_columns() -> None:
    """Ensure close_date, incentive_eligibility, quota exist on rep_incentives."""
    for col_sql in [
        "ALTER TABLE rep_incentives ADD COLUMN close_date DATE NULL",
        "ALTER TABLE rep_incentives ADD COLUMN incentive_eligibility VARCHAR(50) NULL",
        "ALTER TABLE rep_incentives ADD COLUMN quota DECIMAL(18, 2) NULL",
    ]:
        try:
            with db_cursor(commit=True) as cursor:
                cursor.execute(col_sql)
        except psycopg2.Error as e:
            if not _is_duplicate_column_error(e):
                raise


def insert_rep_incentives(
    incentives: List[Tuple[Any, ...]],
) -> None:
    """Bulk insert rep incentives. Tuple length 12 (legacy) or 13 with ``quota`` last."""
    if not incentives:
        return
    _ensure_rep_incentives_extra_columns()
    rows: List[Tuple[Any, ...]] = []
    for t in incentives:
        if len(t) == 13:
            rows.append(t)
        elif len(t) == 12:
            rows.append(tuple(list(t) + [None]))
        else:
            raise ValueError(f"insert_rep_incentives: expected 12 or 13 values, got {len(t)}")
    sql = """
    INSERT INTO rep_incentives (user_id, team_id, total_deals_closed, total_revenue, total_paid_amount, slab_id, incentive_percentage, incentive_amount, payment_status, calculation_period, close_date, incentive_eligibility, quota)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    with db_cursor(commit=True) as cursor:
        cursor.executemany(sql, rows)


def get_rep_incentives(
    user_id: Optional[int] = None,
    team_id: Optional[int] = None,
    period: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Get rep incentives with optional filters."""
    sql = """
    SELECT ri.*, u.full_name, u.email, u.compensation_group AS owner_compensation_group,
           u.hubspot_quota_usd, u.hubspot_quota_period, t.team_name
    FROM rep_incentives ri
    JOIN users u ON ri.user_id = u.user_id
    JOIN teams t ON ri.team_id = t.team_id
    WHERE 1=1
    """
    params = []
    if user_id:
        sql += " AND ri.user_id = %s"
        params.append(user_id)
    if team_id:
        sql += " AND ri.team_id = %s"
        params.append(team_id)
    if period:
        sql += " AND ri.calculation_period = %s"
        params.append(period)
    sql += " ORDER BY ri.created_at DESC, u.full_name"
    with db_cursor(commit=False) as cursor:
        cursor.execute(sql, params or None)
        return cursor.fetchall() or []


def delete_rep_incentive(rep_incentive_id: int) -> None:
    """Delete a single rep incentive by ID."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM rep_incentives WHERE rep_incentive_id = %s", (rep_incentive_id,))


def delete_rep_incentives_by_period(calculation_period: str) -> None:
    """Delete all rep incentives for a calculation period (e.g. before re-finalizing)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM rep_incentives WHERE calculation_period = %s", (calculation_period,))


def delete_all_rep_incentives() -> int:
    """Delete all rep incentives. Returns number deleted."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM rep_incentives")
        return cursor.rowcount if cursor.rowcount is not None else 0


def update_rep_incentive_payment_status(rep_incentive_id: int, new_status: str) -> None:
    """Update payment status for a rep incentive (UNPAID, PAID, PARTIALLY_PAID)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            "UPDATE rep_incentives SET payment_status = %s WHERE rep_incentive_id = %s",
            (new_status, rep_incentive_id),
        )


# --- Team Incentives ---
def insert_team_incentives(
    incentives: List[Tuple[Any, ...]],
) -> None:
    """Bulk insert team incentives."""
    if not incentives:
        return
    sql = """
    INSERT INTO team_incentives (team_id, team_lead_id, total_team_deals, total_team_revenue, total_paid_amount, slab_id, incentive_percentage, incentive_amount, payment_status, calculation_period)
    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """
    with db_cursor(commit=True) as cursor:
        cursor.executemany(sql, incentives)


def get_team_incentives(
    team_id: Optional[int] = None,
    team_lead_id: Optional[int] = None,
    period: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Get team incentives with optional filters."""
    sql = """
    SELECT ti.*, t.team_name, u.full_name as team_lead_name, u.email as team_lead_email
    FROM team_incentives ti
    JOIN teams t ON ti.team_id = t.team_id
    JOIN users u ON ti.team_lead_id = u.user_id
    WHERE 1=1
    """
    params = []
    if team_id:
        sql += " AND ti.team_id = %s"
        params.append(team_id)
    if team_lead_id:
        sql += " AND ti.team_lead_id = %s"
        params.append(team_lead_id)
    if period:
        sql += " AND ti.calculation_period = %s"
        params.append(period)
    sql += " ORDER BY ti.created_at DESC, t.team_name"
    with db_cursor(commit=False) as cursor:
        cursor.execute(sql, params or None)
        return cursor.fetchall() or []


def delete_team_incentive(team_incentive_id: int) -> None:
    """Delete a single team incentive by ID."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM team_incentives WHERE team_incentive_id = %s", (team_incentive_id,))


def delete_team_incentives_by_period(calculation_period: str) -> None:
    """Delete all team incentives for a calculation period (e.g. before re-finalizing)."""
    with db_cursor(commit=True) as cursor:
        cursor.execute("DELETE FROM team_incentives WHERE calculation_period = %s", (calculation_period,))


# --- Audit Logs ---
def log_audit(
    action_type: str,
    entity_type: str,
    performed_by: Optional[int] = None,
    entity_id: Optional[str] = None,
    details: Optional[str] = None,
) -> None:
    """Write an audit log entry."""
    sql = """
    INSERT INTO audit_logs (action_type, performed_by, entity_type, entity_id, details)
    VALUES (%s, %s, %s, %s, %s)
    """
    with db_cursor(commit=True) as cursor:
        cursor.execute(
            sql,
            (action_type, performed_by, entity_type, entity_id, details),
        )


# --- Legacy compatibility (old Partner Compensation Tool) ---
def get_all_compensation_records() -> List[Dict[str, Any]]:
    """Legacy: returns empty list if partner_compensation table doesn't exist."""
    try:
        with db_cursor(commit=False) as cursor:
            cursor.execute(
                "SELECT * FROM partner_compensation ORDER BY created_at DESC"
            )
            return cursor.fetchall() or []
    except Exception:
        return []


def get_record_by_id(record_id: int) -> Optional[Dict[str, Any]]:
    """Legacy: get record by id from partner_compensation."""
    try:
        with db_cursor(commit=False) as cursor:
            cursor.execute(
                "SELECT * FROM partner_compensation WHERE id = %s",
                (record_id,),
            )
            return cursor.fetchone()
    except Exception:
        return None


def update_payment_status(record_id: int, new_status: str) -> None:
    """Legacy: update payment status in partner_compensation."""
    try:
        with db_cursor(commit=True) as cursor:
            cursor.execute(
                "UPDATE partner_compensation SET payment_status = %s WHERE id = %s",
                (new_status, record_id),
            )
    except Exception:
        pass


def insert_compensation_record(
    deal_id: str,
    sales_rep_name: str,
    region: str,
    business_unit: str,
    revenue: float,
    target: float,
    achievement_percent: float,
    payout_rate_percent: float,
    payout_amount: float,
    quarter: str,
    payment_status: str = "Pending",
) -> int:
    """Legacy: insert into partner_compensation. Creates table if needed."""
    try:
        initialize_schema()
    except Exception:
        pass
    create_legacy = """
    CREATE TABLE IF NOT EXISTS partner_compensation (
        id SERIAL PRIMARY KEY,
        deal_id VARCHAR(255) NOT NULL,
        sales_rep_name VARCHAR(255) NOT NULL,
        region VARCHAR(100) NOT NULL,
        business_unit VARCHAR(100) NOT NULL,
        revenue DECIMAL(18, 2) NOT NULL,
        target DECIMAL(18, 2) NOT NULL,
        achievement_percent DECIMAL(6, 2) NOT NULL,
        payout_rate_percent DECIMAL(5, 2) NOT NULL,
        payout_amount DECIMAL(18, 2) NOT NULL,
        quarter VARCHAR(20) NOT NULL,
        payment_status VARCHAR(50) NOT NULL DEFAULT 'Pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """
    with db_cursor(commit=True) as cursor:
        cursor.execute(create_legacy)
        cursor.execute(
            """INSERT INTO partner_compensation (
                deal_id, sales_rep_name, region, business_unit,
                revenue, target, achievement_percent, payout_rate_percent,
                payout_amount, quarter, payment_status
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id""",
            (
                deal_id, sales_rep_name, region, business_unit,
                revenue, target, achievement_percent, payout_rate_percent,
                payout_amount, quarter, payment_status,
            ),
        )
        row = cursor.fetchone()
        return int(row["id"])
