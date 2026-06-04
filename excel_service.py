"""
Excel upload and validation for CloudFuze Migrate Incentive Calculator.

Expected columns: Deal Name, Deal Owner, Amount, Payment Status, Team. Optional: Paid Amount, Close Date,
License resale exclusion (SMB: Microsoft/Google license resale — no commission).
"""

import io
from datetime import date, datetime
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd

from commission_policy import smb_rep_incentive_ineligible_display
from database import get_all_users_with_teams, get_team_by_name


def is_license_resale_excluded_deal(deal: Union[Dict[str, Any], Any]) -> bool:
    """
    True if deal is excluded from SMB commission (Microsoft/Google license resale per policy).

    Set via Excel column **License resale exclusion** or database ``deals.license_resale_exclusion``.
    """
    if hasattr(deal, "get"):
        v = deal.get("license_resale_exclusion")
        if v is True:
            return True
        if isinstance(v, str) and v.strip().lower() in ("yes", "true", "1", "y", "x"):
            return True
    return False


def effective_paid_amount_from_status(deal: Union[Dict[str, Any], Any]) -> float:
    """
    Amount to use as "paid" for incentives and display.

    - UNPAID → 0
    - PAID → full deal amount when stored paid_amount is missing/zero; otherwise stored paid_amount
      (so Excel can still set a partial paid amount).
    - PARTIALLY_PAID → stored paid_amount, or half of amount when paid is zero
    """
    if hasattr(deal, "get"):
        status = (deal.get("payment_status") or "").strip().upper()
        amount = float(deal.get("amount") or 0)
        paid = float(deal.get("paid_amount") or 0)
    else:
        status = (getattr(deal, "payment_status", None) or "").strip().upper()
        amount = float(getattr(deal, "amount", 0) or 0)
        paid = float(getattr(deal, "paid_amount", 0) or 0)
    if status in ("REFUNDED", "PARTIALLY_REFUNDED"):
        return 0.0
    if status == "UNPAID":
        return 0.0
    if status == "PAID":
        return amount if paid == 0 else paid
    if status == "PARTIALLY_PAID":
        return paid if paid > 0 else amount * 0.5
    return paid


@dataclass
class ValidationError:
    row: int
    column: str
    message: str


@dataclass
class ParsedDeal:
    deal_name: str
    deal_owner: str  # name or email
    amount: float
    paid_amount: float  # amount actually paid; used for incentive $ (slab % from Amount)
    payment_status: str
    team: str
    close_date: Optional[date] = None  # from uploaded file; optional
    incentive_eligibility: str = "Eligible"  # Eligible or below-minimum label (from file)
    license_resale_exclusion: bool = False  # SMB: exclude from commission (MS/Google license resale)


@dataclass
class ValidatedDeal:
    """Parsed deal with resolved user_id and team_id."""
    deal_name: str
    user_id: int
    team_id: int
    amount: float
    paid_amount: float
    payment_status: str
    close_date: Optional[date] = None
    incentive_eligibility: str = "Eligible"
    license_resale_exclusion: bool = False


@dataclass
class ParseResult:
    deals: List[ParsedDeal] = field(default_factory=list)
    errors: List[ValidationError] = field(default_factory=list)
    columns_found: List[str] = field(default_factory=list)


# Flexible column name mappings (case-insensitive)
COLUMN_ALIASES = {
    "deal name": ["deal name", "dealname", "deal"],
    "deal owner": ["deal owner", "dealowner", "owner", "sales rep"],
    "amount": ["amount", "revenue", "value", "deal value"],
    "paid amount": ["paid amount", "paidamount", "paid"],
    "payment status": ["payment status", "paymentstatus", "status"],
    "team": ["team", "team name", "teamname"],
    "close date": ["close date", "closedate", "date closed", "closed"],
    "incentive eligibility": ["incentive eligibility", "incentiveeligibility", "eligibility", "eligible", "commission eligibility"],
    "license resale exclusion": [
        "license resale exclusion",
        "licenseresaleexclusion",
        "ms google license",
        "microsoft google license",
        "exclude from commission",
        "license resale",
    ],
}


def _normalize_column(name: str) -> str:
    return str(name).strip().lower()


def _resolve_column(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    """Find DataFrame column that matches any alias."""
    cols = {_normalize_column(c): c for c in df.columns}
    for a in aliases:
        if a in cols:
            return cols[a]
    return None


def parse_excel(file_content: bytes, filename: str = "") -> ParseResult:
    """
    Parse Excel file and extract deals.

    Expected columns (flexible): Deal Name, Deal Owner, Amount, Payment Status, Team

    Returns:
        ParseResult with deals and validation errors.
    """
    result = ParseResult()
    try:
        df = pd.read_excel(io.BytesIO(file_content), engine="openpyxl")
    except Exception as e:
        result.errors.append(ValidationError(0, "file", f"Could not read Excel: {e}"))
        return result

    result.columns_found = list(df.columns)

    # Resolve columns
    deal_name_col = _resolve_column(df, COLUMN_ALIASES["deal name"])
    deal_owner_col = _resolve_column(df, COLUMN_ALIASES["deal owner"])
    amount_col = _resolve_column(df, COLUMN_ALIASES["amount"])
    paid_amount_col = _resolve_column(df, COLUMN_ALIASES["paid amount"])  # optional; default 0
    payment_col = _resolve_column(df, COLUMN_ALIASES["payment status"])
    team_col = _resolve_column(df, COLUMN_ALIASES["team"])
    close_date_col = _resolve_column(df, COLUMN_ALIASES["close date"])  # optional
    incentive_eligibility_col = _resolve_column(df, COLUMN_ALIASES["incentive eligibility"])  # optional
    license_resale_col = _resolve_column(df, COLUMN_ALIASES["license resale exclusion"])  # optional

    missing = []
    if not deal_name_col:
        missing.append("Deal Name")
    if not deal_owner_col:
        missing.append("Deal Owner")
    if not amount_col:
        missing.append("Amount")
    if not payment_col:
        missing.append("Payment Status")
    if not team_col:
        missing.append("Team")

    if missing:
        result.errors.append(
            ValidationError(0, "columns", f"Missing required columns: {', '.join(missing)}")
        )
        return result

    valid_statuses = {
        "paid",
        "unpaid",
        "partially_paid",
        "partially paid",
        "refunded",
        "partially_refunded",
        "partially refunded",
    }
    for idx, row in df.iterrows():
        row_num = int(idx) + 2  # 1-based, +1 for header
        deal_name = str(row.get(deal_name_col, "")).strip() if pd.notna(row.get(deal_name_col)) else ""
        deal_owner = str(row.get(deal_owner_col, "")).strip() if pd.notna(row.get(deal_owner_col)) else ""
        amount_val = row.get(amount_col)
        paid_amount_val = row.get(paid_amount_col) if paid_amount_col and pd.notna(row.get(paid_amount_col)) else 0
        payment_val = str(row.get(payment_col, "")).strip().lower().replace(" ", "_") if pd.notna(row.get(payment_col)) else ""
        team_val = str(row.get(team_col, "")).strip() if pd.notna(row.get(team_col)) else ""

        # Validate mandatory
        if not deal_name:
            result.errors.append(ValidationError(row_num, "Deal Name", "Required"))
            continue
        if not deal_owner:
            result.errors.append(ValidationError(row_num, "Deal Owner", "Required"))
            continue
        if not team_val:
            result.errors.append(ValidationError(row_num, "Team", "Required"))
            continue

        # Validate amount
        try:
            amount = float(amount_val)
            if amount < 0:
                result.errors.append(ValidationError(row_num, "Amount", "Must be non-negative"))
                continue
        except (TypeError, ValueError):
            result.errors.append(ValidationError(row_num, "Amount", "Must be a valid number"))
            continue

        # Paid amount (optional; default 0)
        try:
            paid_amount = float(paid_amount_val) if paid_amount_val not in (None, "") else 0.0
            if paid_amount < 0:
                result.errors.append(ValidationError(row_num, "Paid Amount", "Must be non-negative"))
                continue
            if paid_amount > amount:
                result.errors.append(ValidationError(row_num, "Paid Amount", "Cannot exceed Amount"))
                continue
        except (TypeError, ValueError):
            result.errors.append(ValidationError(row_num, "Paid Amount", "Must be a valid number"))
            continue

        # Normalize payment status
        if not payment_val or payment_val not in valid_statuses:
            payment_val = "unpaid"
        payment_val = payment_val.upper().replace(" ", "_")
        if "REFUND" in payment_val:
            payment_val = "PARTIALLY_REFUNDED" if "PARTIAL" in payment_val else "REFUNDED"
        elif "PARTIAL" in payment_val and "PAID" in payment_val:
            payment_val = "PARTIALLY_PAID"
        elif payment_val == "PAID":
            payment_val = "PAID"
        else:
            payment_val = "UNPAID"

        close_date_val = None
        if close_date_col and pd.notna(row.get(close_date_col)):
            try:
                raw = row.get(close_date_col)
                if hasattr(raw, "date"):
                    close_date_val = raw.date()
                else:
                    dt = pd.to_datetime(raw)
                    close_date_val = dt.date() if hasattr(dt, "date") else date(dt.year, dt.month, dt.day)
            except Exception:
                pass

        eligibility_val = "Eligible"
        if incentive_eligibility_col and pd.notna(row.get(incentive_eligibility_col)):
            raw_el = str(row.get(incentive_eligibility_col)).strip().lower()
            if raw_el in ("not eligible", "n/a", "no", "0", "na", "ineligible"):
                eligibility_val = smb_rep_incentive_ineligible_display()
            elif "below" in raw_el and "quota" in raw_el and "achievement" in raw_el:
                eligibility_val = smb_rep_incentive_ineligible_display()
            elif raw_el in ("eligible", "yes", "1"):
                eligibility_val = "Eligible"

        license_excl = False
        if license_resale_col and pd.notna(row.get(license_resale_col)):
            raw_lr = str(row.get(license_resale_col)).strip().lower()
            license_excl = raw_lr in ("yes", "true", "1", "y", "x", "exclude")

        result.deals.append(
            ParsedDeal(
                deal_name=deal_name,
                deal_owner=deal_owner,
                amount=amount,
                paid_amount=paid_amount,
                payment_status=payment_val,
                team=team_val,
                close_date=close_date_val,
                incentive_eligibility=eligibility_val,
                license_resale_exclusion=license_excl,
            )
        )

    return result


def validate_against_db(parsed: ParseResult) -> Tuple[List["ValidatedDeal"], List[ValidationError]]:
    """
    Validate that Deal Owner and Team exist in the database.

    Returns:
        (valid_deals with user_id/team_id, errors)
    """
    from database import get_all_teams

    users = get_all_users_with_teams(active_only=False)
    user_by_name = {u["full_name"].strip().lower(): u for u in users}
    user_by_email = {u["email"].strip().lower(): u for u in users}
    all_teams = get_all_teams()
    teams = {t["team_name"].strip().lower(): t for t in all_teams}

    valid: List[ValidatedDeal] = []
    errors = list(parsed.errors)

    for i, d in enumerate(parsed.deals):
        row_num = i + 2
        owner_key = d.deal_owner.strip().lower()
        owner = user_by_name.get(owner_key) or user_by_email.get(owner_key)
        if not owner:
            errors.append(ValidationError(row_num, "Deal Owner", f"User '{d.deal_owner}' not found"))
            continue

        team_key = d.team.strip().lower()
        team = teams.get(team_key)
        if not team:
            t = get_team_by_name(d.team)
            if t:
                team = t
                teams[team_key] = t
        if not team:
            errors.append(ValidationError(row_num, "Team", f"Team '{d.team}' not found"))
            continue

        valid.append(
            ValidatedDeal(
                deal_name=d.deal_name,
                user_id=owner["user_id"],
                team_id=team["team_id"],
                amount=d.amount,
                paid_amount=d.paid_amount,
                payment_status=d.payment_status,
                close_date=getattr(d, "close_date", None),
                incentive_eligibility=getattr(d, "incentive_eligibility", "Eligible"),
                license_resale_exclusion=bool(getattr(d, "license_resale_exclusion", False)),
            )
        )

    return valid, errors


def create_sample_excel() -> bytes:
    """Create a sample Excel file with correct column headers."""
    df = pd.DataFrame({
        "Deal Name": ["Sample Deal 1", "Sample Deal 2"],
        "Deal Owner": ["System Admin", "System Admin"],
        "Amount": [50000, 120000],
        "Paid Amount": [0, 120000],
        "Payment Status": ["UNPAID", "PAID"],
        "Team": ["SMB", "Enterprise"],
        "Close Date": ["2026-02-01", "2026-02-15"],
        "Incentive Eligibility": ["Eligible", smb_rep_incentive_ineligible_display()],
        "License resale exclusion": ["No", "No"],
    })
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()
