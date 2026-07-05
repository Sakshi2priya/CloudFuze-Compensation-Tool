"""
CloudFuze Migrate – Incentive Calculator Application

Admin: Upload Excel, validate, finalize, calculate incentives, export reports.
Member: Read-only view of own/team incentives.
"""

import contextlib
import hashlib
import html as html_module
import io
import math
import os
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import altair as alt
from dotenv import load_dotenv
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from auth import authenticate, ensure_admin_user, hash_password, validate_email, verify_password
from database import (
    create_upload, create_user, delete_all_rep_incentives, delete_all_uploads, delete_outbound_meeting,
    delete_rep_incentive, delete_team_incentive, delete_upload, ensure_outbound_team, get_all_outbound_meetings,
    get_all_teams, get_all_users_with_teams, get_deals_by_upload, get_deals_from_finalized_uploads,
    get_rep_incentives, get_team_incentives, get_uploads_for_user, get_user_by_email, initialize_schema,
    insert_deals, insert_outbound_meeting, log_audit, sync_deal_paid_amounts_from_status,
    update_team_goal, update_upload_status,
    update_user_password, update_user_profile,
)
from excel_service import (
    create_sample_excel,
    effective_paid_amount_from_status,
    parse_excel,
    validate_against_db,
    ParseResult,
    ParsedDeal,
    ValidationError,
)
from hubspot_service import (
    build_hubspot_goal_sync_plan,
    fetch_and_map_hubspot_deals,
    filter_hubspot_owner_labels_for_team,
    get_access_token,
)


@st.cache_data(ttl=600, show_spinner=False)
def _cached_hubspot_owners(cache_key: str, token: str):
    from hubspot_service import fetch_owners

    return fetch_owners(token)


def _hubspot_team_extra_tokens_from_users(team: str) -> list[str]:
    """
    Substrings to match HubSpot owner labels against User Management users on that team.
    Uses each user’s email (lower) and the part before @ so roster picks stay in sync with the DB.
    """
    if not team or str(team).strip() in ("", "Any"):
        return []
    want = str(team).strip()
    out: list[str] = []
    for u in get_all_users_with_teams(active_only=False):
        if (u.get("team_name") or "").strip() != want:
            continue
        em = (u.get("email") or "").strip().lower()
        if not em:
            continue
        out.append(em)
        if "@" in em:
            out.append(em.split("@", 1)[0])
    return out


@st.cache_data(ttl=600, show_spinner=False)
def _cached_hubspot_stages(cache_key: str, token: str):
    from hubspot_service import fetch_deal_pipeline_stages

    return fetch_deal_pipeline_stages(token)


@st.cache_data(ttl=600, show_spinner=False)
def _cached_hubspot_payment_options(cache_key: str, token: str):
    """Enumeration labels/values for the configured HubSpot payment status property."""
    from hubspot_service import _payment_status_property_name, fetch_deal_enumeration_options

    pn = _payment_status_property_name()
    if not pn:
        return []
    try:
        return fetch_deal_enumeration_options(token, pn)
    except Exception:
        return []


def _hubspot_cache_key(token: str) -> str:
    # Bump when HubSpot stage/payment shape changes so cache refreshes.
    return hashlib.sha256(f"{token}:hubspot_v3".encode()).hexdigest()[:40]


def _fmt_dollar(x):
    """Format numeric value as dollar string ($1,234.56)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    try:
        return f"${float(x):,.2f}"
    except (TypeError, ValueError):
        return x


def _format_df_dollars(df: pd.DataFrame, dollar_columns: list) -> pd.DataFrame:
    """Return a copy of df with specified columns formatted as dollar strings for display."""
    df = df.copy()
    for col in dollar_columns:
        if col in df.columns:
            df[col] = df[col].apply(_fmt_dollar)
    return df


# Distinct hues for categorical charts (cycles for many categories).
_CHART_VIVID_PALETTE: tuple[str, ...] = (
    "#1e88e5",
    "#00897b",
    "#8e24aa",
    "#fb8c00",
    "#43a047",
    "#e53935",
    "#5c6bc0",
    "#00acc1",
    "#d81b60",
    "#fdd835",
    "#6d4c41",
    "#3949ab",
)


def _chart_categorical_color(field: str, categories: list, *, legend) -> alt.Color:
    """Ordinal color encoding with vivid distinct hues (domain order preserved). ``legend`` = ``None`` to hide."""
    cats = [str(c) for c in categories]
    pal = _CHART_VIVID_PALETTE
    rng = [pal[i % len(pal)] for i in range(len(cats))]
    return alt.Color(field, scale=alt.Scale(domain=cats, range=rng), legend=legend)


def _enrich_rep_incentives_display(df: pd.DataFrame) -> pd.DataFrame:
    """
    For SMB rows: recompute ``quota`` with ``smb_individual_quota_usd_for_rep`` (policy list → overrides → HubSpot →
    Group A/B defaults) so displayed quota matches the engine. For Account Management when quota achievement is enabled,
    recompute with ``am_individual_quota_usd_for_rep``. Otherwise non-SMB: keep stored ``quota`` when set.
    ``quota_achievement_pct`` = total_revenue ÷ quota × 100.
    """
    from commission_policy import (
        ACCOUNT_MANAGEMENT_TEAM_NAME,
        AM_QUOTA_ACHIEVEMENT_ENABLED,
        SMB_TEAM_NAME,
        am_individual_quota_usd_for_rep,
        smb_individual_quota_usd_for_rep,
    )

    out = df.copy()
    if "total_revenue" not in out.columns:
        return out

    def _eff_quota(row):
        tn = (row.get("team_name") or "").strip()
        if tn == SMB_TEAM_NAME:
            eq = smb_individual_quota_usd_for_rep(
                row.get("owner_compensation_group"),
                row.get("hubspot_quota_usd"),
                full_name=row.get("full_name"),
                email=row.get("email"),
                calculation_period=row.get("calculation_period"),
            )
            return eq if eq > 0 else None
        if tn == ACCOUNT_MANAGEMENT_TEAM_NAME and AM_QUOTA_ACHIEVEMENT_ENABLED:
            eq = am_individual_quota_usd_for_rep(
                row.get("hubspot_quota_usd"),
                full_name=row.get("full_name"),
                email=row.get("email"),
                calculation_period=row.get("calculation_period"),
            )
            return eq if eq > 0 else None
        q = row.get("quota")
        try:
            if q is not None and not pd.isna(q) and float(q) > 0:
                return float(q)
        except (TypeError, ValueError):
            pass
        return None

    eff = out.apply(_eff_quota, axis=1)
    new_q = []
    for i in range(len(out)):
        row = out.iloc[i]
        q = row.get("quota")
        e = eff.iloc[i]
        is_smb = (row.get("team_name") or "").strip() == SMB_TEAM_NAME
        is_am_q = (row.get("team_name") or "").strip() == ACCOUNT_MANAGEMENT_TEAM_NAME and AM_QUOTA_ACHIEVEMENT_ENABLED
        if is_smb or is_am_q:
            try:
                if e is not None and not (isinstance(e, float) and pd.isna(e)) and float(e) > 0:
                    new_q.append(float(e))
                else:
                    try:
                        if q is not None and not pd.isna(q) and float(q) > 0:
                            new_q.append(float(q))
                        else:
                            new_q.append(None)
                    except (TypeError, ValueError):
                        new_q.append(None)
            except (TypeError, ValueError):
                new_q.append(None)
        else:
            try:
                if q is not None and not pd.isna(q) and float(q) > 0:
                    new_q.append(float(q))
                else:
                    new_q.append(e)
            except (TypeError, ValueError):
                new_q.append(e)
    out["quota"] = new_q
    out["quota_achievement_pct"] = [
        round(float(tr) / float(eq) * 100.0, 2)
        if eq is not None
        and not (isinstance(eq, float) and pd.isna(eq))
        and float(eq) > 0
        and tr is not None
        and not pd.isna(tr)
        else None
        for tr, eq in zip(out["total_revenue"], new_q)
    ]
    return out


def _team_goal_achievement_pct_value(row: dict, team_goal_by_id: dict):
    """Total team revenue ÷ database team goal × 100, or ``None`` if goal missing."""
    tid = row.get("team_id")
    tg = team_goal_by_id.get(tid) if tid is not None else None
    try:
        rev = float(row.get("total_team_revenue") or 0)
    except (TypeError, ValueError):
        rev = 0.0
    try:
        if tg is None or (isinstance(tg, float) and pd.isna(tg)):
            return None
        tg_f = float(tg)
    except (TypeError, ValueError):
        return None
    if tg_f <= 0:
        return None
    return round((rev / tg_f) * 100.0, 2)


def _manager_incentive_zero_reason(row: dict, team_goal_by_id: dict) -> str:
    """Short note when manager incentive is $0 — team-aware so ENT doesn't get the SMB 50% floor."""
    from commission_policy import (
        ACCOUNT_MANAGEMENT_TEAM_NAME,
        AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
        ENTERPRISE_TEAM_NAME,
        SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
        SMB_TEAM_NAME,
    )

    try:
        amt = float(row.get("incentive_amount") or 0)
    except (TypeError, ValueError):
        amt = 0.0
    if amt > 0:
        return "—"
    ach = _team_goal_achievement_pct_value(row, team_goal_by_id)
    if ach is None:
        return "No team goal set"

    # Pick the right minimum-achievement floor for this team.
    team_name = (row.get("team_name") or "").strip()
    if team_name == SMB_TEAM_NAME:
        team_min = float(SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT or 0)
    elif team_name == ACCOUNT_MANAGEMENT_TEAM_NAME:
        team_min = float(AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT or 0)
    elif team_name == ENTERPRISE_TEAM_NAME:
        team_min = 0.0  # ENT policy: minimum achievement is 0%, bands start at 0–75.99% → 7%.
    else:
        team_min = float(SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT or 0)

    if team_min > 0 and ach < team_min:
        return f"Below {team_min:.0f}% team goal achievement"
    try:
        pct = float(row.get("incentive_percentage") or 0)
    except (TypeError, ValueError):
        pct = 0.0
    if pct <= 0:
        return "No commission at this achievement tier (policy)"
    return "—"


def _enrich_team_incentives_display(df: pd.DataFrame) -> pd.DataFrame:
    """Add **Team goal achievement %**, recompute ENT incentive on-the-fly (using ENT bands),
    and add **Payout note**.

    The on-the-fly ENT recomputation is needed because rows already stored in the database
    were computed with the old shared thresholds; this corrects the display without requiring
    the user to re-finalize the upload.
    """
    if df.empty:
        return df
    from commission_policy import ENTERPRISE_TEAM_NAME, ent_team_commission_pct_from_achievement

    team_goal_by_id = {t["team_id"]: t.get("team_goal") for t in get_all_teams()}
    out = df.copy()
    out["Team goal achievement %"] = out.apply(
        lambda r: _team_goal_achievement_pct_value(r.to_dict(), team_goal_by_id),
        axis=1,
    )

    # Recompute ENT rows so the displayed incentive_percentage / incentive_amount match the
    # ENT policy bands (0–75.99 → 7%, 76–100 → 9%, 101–125 → 11%, 126%+ → 13%).
    def _recompute_ent(r):
        if (r.get("team_name") or "").strip() != ENTERPRISE_TEAM_NAME:
            return r
        ach = _team_goal_achievement_pct_value(r.to_dict(), team_goal_by_id) or 0.0
        new_pct = ent_team_commission_pct_from_achievement(float(ach))
        try:
            rev = float(r.get("total_team_revenue") or 0)
        except (TypeError, ValueError):
            rev = 0.0
        new_amt = round(rev * new_pct / 100.0, 2)
        r["incentive_percentage"] = new_pct
        r["incentive_amount"] = new_amt
        return r

    out = out.apply(_recompute_ent, axis=1)

    out["Payout note"] = out.apply(
        lambda r: _manager_incentive_zero_reason(r.to_dict(), team_goal_by_id),
        axis=1,
    )
    return out


def _manager_incentive_column_config():
    """Streamlit column config: teal-style progress for team achievement %."""
    return {
        "Team goal achievement %": st.column_config.ProgressColumn(
            "Team goal achievement %",
            help="Total team revenue ÷ team goal (from Team Goals above).",
            format="%.1f%%",
            min_value=0,
            max_value=200,
        ),
    }


def _format_pct_display(df: pd.DataFrame, pct_columns: list) -> pd.DataFrame:
    """Format numeric percentage columns as strings with a %% suffix."""
    df = df.copy()
    for col in pct_columns:
        if col not in df.columns:
            continue
        df[col] = df[col].apply(
            lambda x: f"{float(x):.2f}%" if x is not None and not (isinstance(x, float) and pd.isna(x)) else ""
        )
    return df


def _quarter_month_labels(year: int, quarter: int) -> list[str]:
    """Month labels like ``Jan 2026`` for ``calculation_period`` matching in rep_incentives."""
    months_map = {1: (1, 2, 3), 2: (4, 5, 6), 3: (7, 8, 9), 4: (10, 11, 12)}
    return [datetime(year, m, 1).strftime("%b %Y") for m in months_map[quarter]]


# Close-date preset list (mirrors the HubSpot Forecast date picker).
# Fiscal year is treated as calendar year (Jan–Dec). If your fiscal year is offset,
# adjust _resolve_close_date_range below.
_CLOSE_DATE_PRESETS: tuple[str, ...] = (
    "Today",
    "Yesterday",
    "Tomorrow",
    "This week",
    "This week so far",
    "Last week",
    "Next week",
    "This month",
    "This month so far",
    "Last month",
    "This quarter",
    "This fiscal year",
    "This quarter so far",
    "This fiscal quarter so far",
    "Last quarter",
    "Last fiscal quarter",
    "Next quarter",
    "Next fiscal quarter",
    "This year",
    "Last year",
    "Next year",
    "Last 7 days",
    "Last 14 days",
    "Last 30 days",
    "Last 60 days",
    "Last 90 days",
    "Last 180 days",
    "Custom date range",
)


def _quarter_of_month(month: int) -> int:
    return (int(month) - 1) // 3 + 1


def _quarter_start_end(year: int, quarter: int):
    from calendar import monthrange as _mrng
    start_m = (int(quarter) - 1) * 3 + 1
    end_m = start_m + 2
    last = _mrng(int(year), end_m)[1]
    return datetime(int(year), start_m, 1).date(), datetime(int(year), end_m, last).date()


def _iso_week_monday(d):
    return d - timedelta(days=d.weekday())


def _resolve_close_date_range(preset: str, today=None, custom_start=None, custom_end=None):
    """Map a preset label (or a custom range) to a (start_date, end_date) inclusive tuple.

    Returns ``(start, end)`` as ``date`` objects. ``today`` defaults to the current local date.
    """
    from calendar import monthrange as _mrng
    if today is None:
        today = datetime.now().date()
    p = (preset or "").strip()

    if p == "Today":
        return today, today
    if p == "Yesterday":
        d = today - timedelta(days=1)
        return d, d
    if p == "Tomorrow":
        d = today + timedelta(days=1)
        return d, d

    week_start = _iso_week_monday(today)
    week_end = week_start + timedelta(days=6)
    if p == "This week":
        return week_start, week_end
    if p == "This week so far":
        return week_start, today
    if p == "Last week":
        last_start = week_start - timedelta(days=7)
        return last_start, last_start + timedelta(days=6)
    if p == "Next week":
        nxt_start = week_start + timedelta(days=7)
        return nxt_start, nxt_start + timedelta(days=6)

    if p == "This month":
        first = today.replace(day=1)
        last = first.replace(day=_mrng(first.year, first.month)[1])
        return first, last
    if p == "This month so far":
        return today.replace(day=1), today
    if p == "Last month":
        first_this = today.replace(day=1)
        last_prev = first_this - timedelta(days=1)
        first_prev = last_prev.replace(day=1)
        return first_prev, last_prev

    cur_q = _quarter_of_month(today.month)
    if p in ("This quarter", "This fiscal quarter"):
        return _quarter_start_end(today.year, cur_q)
    if p in ("This quarter so far", "This fiscal quarter so far"):
        s, _e = _quarter_start_end(today.year, cur_q)
        return s, today
    if p in ("Last quarter", "Last fiscal quarter"):
        if cur_q == 1:
            return _quarter_start_end(today.year - 1, 4)
        return _quarter_start_end(today.year, cur_q - 1)
    if p in ("Next quarter", "Next fiscal quarter"):
        if cur_q == 4:
            return _quarter_start_end(today.year + 1, 1)
        return _quarter_start_end(today.year, cur_q + 1)

    if p in ("This year", "This fiscal year"):
        return datetime(today.year, 1, 1).date(), datetime(today.year, 12, 31).date()
    if p == "Last year":
        return datetime(today.year - 1, 1, 1).date(), datetime(today.year - 1, 12, 31).date()
    if p == "Next year":
        return datetime(today.year + 1, 1, 1).date(), datetime(today.year + 1, 12, 31).date()

    if p == "Last 7 days":
        return today - timedelta(days=7), today - timedelta(days=1)
    if p == "Last 14 days":
        return today - timedelta(days=14), today - timedelta(days=1)
    if p == "Last 30 days":
        return today - timedelta(days=30), today - timedelta(days=1)
    if p == "Last 60 days":
        return today - timedelta(days=60), today - timedelta(days=1)
    if p == "Last 90 days":
        return today - timedelta(days=90), today - timedelta(days=1)
    if p == "Last 180 days":
        return today - timedelta(days=180), today - timedelta(days=1)

    if p == "Custom date range":
        if custom_start and custom_end:
            return custom_start, custom_end
        return today, today

    # Fallback: today
    return today, today


def _quarter_for_date_range(start_d, end_d) -> tuple[int, int]:
    """Pick a single (year, quarter) to drive target/pinned lookups for an arbitrary range.

    Strategy: use the quarter that contains the *end* date of the range. This makes
    presets like 'Today', 'This month', 'Last 30 days' resolve to the current quarter
    most of the time, which is what users expect for dashboards.
    """
    return int(end_d.year), _quarter_of_month(end_d.month)


def _get_user_rep_name_matchers(user: dict | None) -> list[str] | None:
    """Return the list of lowercase rep-name substrings that match the current user.

    Used to filter pinned Q1 / April / May data so each user sees only their own row
    and deals (deal_owner, reps[], commission.reps[]).

    - ADMIN → returns None (no filter; sees everything)
    - Every other role (SALES_REP, SALES_MANAGER, or anything else) → returns a list
      of matchers derived from full_name + known aliases. Managers who match their own
      manager_name (Chitradip, Joy, Anthony) will still see their manager_incentive block;
      everyone else's data is hidden from them.
    - No/empty user or missing full_name → None (no filter — cautious default)
    """
    if not user:
        return None
    role = (user.get("role") or "").upper()
    if role == "ADMIN":
        return None  # Admins see everything
    full_name = (user.get("full_name") or "").strip().lower()
    if not full_name:
        return None
    matchers: list[str] = [full_name]
    first_name = (full_name.split() or [""])[0]
    if first_name and first_name != full_name:
        matchers.append(first_name)
    # Explicit alias mapping (short rep names used in reps[] JSON vs full user names).
    _aliases = {
        "yogesh vig": ["yogi", "yogesh"],
        "yogesh": ["yogi"],
        "vicky cariappa": ["vicky"],
        "kritika gupta": ["kritika"],
        "deepak r j": ["deepak"],
        "rutuja kawade": ["rutuja"],
        "joy prakash": ["joy"],
        "vivin joseph": ["vivin"],
        "lawrence lewis": ["lawrence", "larry"],
        "royston aden": ["royston"],
        "kartik kashyap": ["kartik"],
        "lennis brown": ["lennis"],
        "chitradip": ["chit"],
        "anthony raymond": ["anthony"],
        "arundhati sen": ["arundhati"],
    }
    for key, extras in _aliases.items():
        if key in full_name or full_name in key:
            matchers.extend(extras)
    # Deduplicate, preserve order
    seen: set[str] = set()
    out: list[str] = []
    for m in matchers:
        if m and m not in seen:
            seen.add(m)
            out.append(m)
    return out


def _rep_name_matches_user(rep_or_owner: str, matchers: list[str] | None) -> bool:
    """Check whether a rep name / deal_owner string matches any of the current user's matchers.

    Strips the "(Deactivated User)" suffix before comparing so deactivated deal owners
    still match their user name.
    """
    if not matchers:
        # No filter → matches everything
        return True
    low = (rep_or_owner or "").lower().replace(" (deactivated user)", "").strip()
    if not low:
        return False
    for m in matchers:
        if not m:
            continue
        if m == low or m in low or low in m:
            return True
        # Match on first word (e.g. "vicky" matches "vicky cariappa" and vice versa)
        first_low = low.split()[0] if low else ""
        first_m = m.split()[0] if m else ""
        if first_low and first_m and (first_low == first_m):
            return True
    return False


def _filter_pinned_for_user(pinned: dict, user: dict | None) -> dict:
    """Return a filtered copy of pinned data showing only the current user's rows/deals.

    If the user is ADMIN or SALES_MANAGER (matchers is None), returns pinned unchanged.
    Otherwise filters:
      - ``reps`` to only the user's rep entry
      - ``paid_deals`` / ``unpaid_deals_overrides`` to only the user's deals
      - ``commission.reps`` to only the user's row
      - ``manager_incentive`` blanked unless the user IS the manager
      - ``deals_summary`` recomputed from the filtered deal lists
    Team-level fields (``team_target_usd``, ``team_achievement_usd``, ``manager_incentive.team_*``)
    are left intact so the user can still see how their contribution rolls into team performance.
    """
    matchers = _get_user_rep_name_matchers(user)
    if not matchers or not pinned:
        return pinned

    filtered = dict(pinned)

    # Filter reps
    filtered["reps"] = [
        r for r in (pinned.get("reps") or [])
        if _rep_name_matches_user(r.get("name") or "", matchers)
    ]
    # Filter deals by deal_owner
    filtered["paid_deals"] = [
        d for d in (pinned.get("paid_deals") or [])
        if _rep_name_matches_user(d.get("deal_owner") or "", matchers)
    ]
    filtered["unpaid_deals_overrides"] = [
        d for d in (pinned.get("unpaid_deals_overrides") or [])
        if _rep_name_matches_user(d.get("deal_owner") or "", matchers)
    ]
    # Filter commission.reps
    if pinned.get("commission"):
        comm = dict(pinned["commission"])
        comm["reps"] = [
            c for c in (pinned["commission"].get("reps") or [])
            if _rep_name_matches_user(c.get("name") or "", matchers)
        ]
        filtered["commission"] = comm
    # Manager incentive: ALWAYS preserved intact (Chitradip for SMB, Joy for AM,
    # Anthony for ENT). Everyone can see the manager's commission calculation —
    # it's team-level context, not personal data being leaked.
    filtered["manager_incentive"] = pinned.get("manager_incentive") or {}
    # Recompute deals_summary from filtered deal lists
    _paid_n = len(filtered["paid_deals"])
    _unpaid_n = sum(
        1 for d in filtered["unpaid_deals_overrides"]
        if not d.get("is_partial_remainder") and not d.get("is_cancelled")
    )
    _cancelled_n = sum(1 for d in filtered["unpaid_deals_overrides"] if d.get("is_cancelled"))
    filtered["deals_summary"] = {
        "total_deals_closed": _paid_n + _unpaid_n,
        "deals_payment_received": _paid_n,
        "deals_payment_not_received": _unpaid_n,
        "deals_cancelled": _cancelled_n,
        "note": "Filtered to your personal view. Team-level totals still shown for context.",
    }
    return filtered


def _close_date_picker(key_prefix: str, default_preset: str = "This quarter"):
    """Render the Close-date picker as a Custom date range only (From / To).

    All preset options (This/Last quarter, This/Last month, YTD, etc.) have been
    removed — Q1 & monthly data are locked from pinned JSON files and the user only
    needs to select a custom From/To range that maps to the desired period.

    Returns: (start_date, end_date, label, year, quarter) — year/quarter are derived from the
    range using :func:`_quarter_for_date_range` and feed downstream policy lookups.
    """
    today = datetime.now().date()
    # Sensible default From/To based on the caller's requested default preset — so the
    # first render still lands on the intended quarter/month range without a preset UI.
    def_start, def_end = _resolve_close_date_range(default_preset, today=today)
    st.markdown("**Close date — Custom date range**")
    col_a, col_b = st.columns(2)
    with col_a:
        cs = st.date_input("From", value=def_start, key=f"{key_prefix}_close_date_from")
    with col_b:
        ce = st.date_input("To", value=def_end, key=f"{key_prefix}_close_date_to")
    start_d, end_d = _resolve_close_date_range(
        "Custom date range", today=today, custom_start=cs, custom_end=ce
    )
    label = f"{start_d:%b %d, %Y} – {end_d:%b %d, %Y}"
    yr, qt = _quarter_for_date_range(start_d, end_d)
    return start_d, end_d, label, yr, qt


def _sum_team_rep_revenue_for_periods(
    inv: list, user_id: int, period_labels: list[str], team_name: str
) -> float:
    """Sum ``total_revenue`` for a rep on a team; prefer rows whose ``calculation_period`` is in the quarter."""
    rows = [
        r
        for r in inv
        if r.get("user_id") == user_id and (r.get("team_name") or "").strip() == team_name
    ]
    if not rows:
        return 0.0
    in_q = [r for r in rows if (r.get("calculation_period") or "").strip() in period_labels]
    use = in_q if in_q else rows
    return float(sum(float(r.get("total_revenue") or 0.0) for r in use))


def _sum_smb_rep_revenue_for_periods(inv: list, user_id: int, period_labels: list[str]) -> float:
    """Sum ``total_revenue`` for SMB rep; prefer rows whose ``calculation_period`` is in the quarter."""
    from commission_policy import SMB_TEAM_NAME

    return _sum_team_rep_revenue_for_periods(inv, user_id, period_labels, SMB_TEAM_NAME)


def _initials_for_avatar(name: str) -> str:
    parts = (name or "").strip().split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    s = (name or "?").strip()
    return (s[:2] if len(s) >= 2 else s + "?").upper()[:2]


def _fmt_usd(x: float) -> str:
    return f"${x:,.2f}"


def _bullet_chart_row_html(
    title: str,
    subtitle: str,
    attained: float,
    target: float,
    *,
    initials_html: str = "",
) -> str:
    """
    Single-row bullet-style chart: qualitative bands vs goal, solid bar for attained, marker at 100% of goal.
    Bands: 0–50% of goal (darker), 50–80% (mid), 80–100% (lighter); remainder to scale max (pale).
    """
    tgt = max(float(target), 0.0)
    att = max(float(attained), 0.0)
    m = max(tgt * 1.22, att * 1.08, 1.0)
    if tgt <= 0:
        m = max(att * 1.15, 1.0)

    p50 = min(100.0, (0.5 * tgt / m * 100.0) if tgt > 0 else 0.0)
    p80 = min(100.0, (0.8 * tgt / m * 100.0) if tgt > 0 else 0.0)
    p100 = min(100.0, (tgt / m * 100.0) if tgt > 0 else 0.0)

    w1 = p50
    w2 = max(0.0, p80 - p50)
    w3 = max(0.0, p100 - p80)
    w4 = max(0.0, 100.0 - p100)
    bar_w = min(100.0, max(0.0, att / m * 100.0))
    mark_pct = p100
    pct_txt = round((att / tgt * 100.0) if tgt > 0 else 0.0)

    left_inner = ""
    if initials_html:
        left_inner = (
            f'<div class="ga-bullet-left ga-bullet-left-rep">'
            f'<div class="ga-rep-line">{initials_html}'
            f'<div><div class="ga-bullet-title">{html_module.escape(title)}</div>'
            f'<div class="ga-bullet-sub">{html_module.escape(subtitle)}</div></div></div></div>'
        )
    else:
        left_inner = (
            f'<div class="ga-bullet-left">'
            f'<div class="ga-bullet-title">{html_module.escape(title)}</div>'
            f'<div class="ga-bullet-sub">{html_module.escape(subtitle)}</div></div>'
        )

    return f"""
<div class="ga-bullet-row">
{left_inner}
<div class="ga-bullet-right">
<div class="ga-bullet-legend">
<span class="ga-leg ga-leg-a">0–50% of goal</span>
<span class="ga-leg ga-leg-b">50–80%</span>
<span class="ga-leg ga-leg-c">80–100%</span>
</div>
<div class="ga-bullet-track-wrap" role="img" aria-label="{html_module.escape(title)} attainment">
<div class="ga-bullet-zones">
<span class="ga-bz ga-bz-a" style="width:{w1:.4f}%"></span>
<span class="ga-bz ga-bz-b" style="width:{w2:.4f}%"></span>
<span class="ga-bz ga-bz-c" style="width:{w3:.4f}%"></span>
<span class="ga-bz ga-bz-d" style="width:{w4:.4f}%"></span>
</div>
<div class="ga-bullet-bar" style="width:{bar_w:.4f}%"></div>
<div class="ga-bullet-marker" style="left:{mark_pct:.4f}%"></div>
</div>
<div class="ga-bullet-footer">
<span class="ga-bullet-pct">{pct_txt}%</span>
<span class="ga-bullet-detail">{html_module.escape(_fmt_usd(att))} of {html_module.escape(_fmt_usd(tgt))}</span>
</div>
</div>
</div>
"""


def _load_pinned_quarter(team_prefix: str, year: int, quarter: int) -> dict | None:
    """Load pinned values for ``team_prefix`` (``"smb"`` or ``"am"``) for a given year/quarter.

    Looks for ``policy/{team_prefix}_q{quarter}_{year}_fixed.json`` (and a couple of legacy variants).
    Returns the parsed dict, or None if the file is missing / its year+quarter don't match.
    """
    import json as _json
    candidates = [
        Path(__file__).parent / "policy" / f"{team_prefix}_q{quarter}_{year}_fixed.json",
        Path(__file__).parent / "policy" / f"{team_prefix}_{year}_q{quarter}_fixed.json",
    ]
    for p in candidates:
        try:
            if p.exists():
                with p.open("r", encoding="utf-8") as f:
                    data = _json.load(f)
                if int(data.get("year", -1)) == int(year) and int(data.get("quarter", -1)) == int(quarter):
                    return data
        except (OSError, ValueError):
            continue
    return None


def _load_pinned_smb_quarter(year: int, quarter: int) -> dict | None:
    """Backward-compatible loader for SMB pinned quarter data."""
    return _load_pinned_quarter("smb", year, quarter)


def _load_pinned_am_quarter(year: int, quarter: int) -> dict | None:
    """Loader for AM pinned quarter data (e.g. policy/am_q1_2026_fixed.json)."""
    return _load_pinned_quarter("am", year, quarter)


_MONTH_ABBR = {1: "jan", 2: "feb", 3: "mar", 4: "apr", 5: "may", 6: "jun",
               7: "jul", 8: "aug", 9: "sep", 10: "oct", 11: "nov", 12: "dec"}
_MONTH_ABBR_LONG = {1: "january", 2: "february", 3: "march", 4: "april", 5: "may", 6: "june",
                    7: "july", 8: "august", 9: "september", 10: "october", 11: "november", 12: "december"}


def _load_pinned_monthly(team_prefix: str, year: int, month: int) -> dict | None:
    """Load a monthly pinned JSON for a team (``team_prefix`` is e.g. ``"smb"`` or ``"am"``).

    Looks for files like ``policy/{team_prefix}_{month_word}_{year}_fixed.json``
    (e.g. ``smb_april_2026_fixed.json``) and validates year/month if those fields exist.
    """
    import json as _json
    if not (1 <= month <= 12):
        return None
    month_abbr = _MONTH_ABBR.get(month)
    month_word = _MONTH_ABBR_LONG.get(month)
    candidates = [
        Path(__file__).parent / "policy" / f"{team_prefix}_{month_word}_{year}_fixed.json",
        Path(__file__).parent / "policy" / f"{team_prefix}_{month_abbr}_{year}_fixed.json",
        Path(__file__).parent / "policy" / f"{team_prefix}_{year}_{month:02d}_fixed.json",
    ]
    for p in candidates:
        try:
            if p.exists():
                with p.open("r", encoding="utf-8") as f:
                    data = _json.load(f)
                yr = int(data.get("year") or 0)
                mo = int(data.get("month") or 0)
                if yr == year and mo == month:
                    return data
        except (OSError, ValueError):
            continue
    return None


def _is_single_month_range(start_d, end_d) -> int:
    """If [start, end] falls within one calendar month, return the month number; else 0."""
    try:
        if start_d.year == end_d.year and start_d.month == end_d.month:
            return int(start_d.month)
    except Exception:
        return 0
    return 0


def _monthly_slab_pct(tiers: list[dict], achievement_pct: float) -> tuple[float, str]:
    """Pick the matching individual-tier commission % for an achievement %.

    Returns ``(commission_pct, label)`` where label is the human band like ``"61–75%"``.
    """
    x = float(achievement_pct or 0)
    for t in (tiers or []):
        lo = float(t.get("min_pct") or 0)
        hi = t.get("max_pct")
        try:
            hi_f = None if hi is None else float(hi)
        except (TypeError, ValueError):
            hi_f = None
        if hi_f is None:
            if x >= lo:
                return float(t.get("commission_pct") or 0), f"{int(lo)}%+"
        else:
            # Use the deck's "61–75%" interpretation: lo and hi inclusive.
            if lo == 0:
                if x < hi_f:
                    return float(t.get("commission_pct") or 0), f"< {int(hi_f)}%"
            else:
                if x >= lo and x <= hi_f:
                    return float(t.get("commission_pct") or 0), f"{int(lo)}–{int(hi_f)}%"
    return 0.0, "—"


def _fetch_monthly_actuals_from_hubspot(
    pinned: dict,
    team_prefix: str,
    extra_owner_ids: list[str] | None = None,
) -> tuple[bool, str]:
    """Pull HubSpot closed-won deals for the pinned month and update the JSON in place.

    Matches HubSpot owner emails to the User Management roster, then resolves each pinned rep
    (by first-name token, since pinned reps use short names like "Vicky", "Yogi") to a roster user
    and aggregates the booked + paid totals into the rep's ``achievement_usd`` and
    ``payment_received_usd`` fields. Writes the updated JSON back to disk.

    ``extra_owner_ids`` further restricts the fetch to a user-chosen subset of HubSpot owners.

    Returns ``(ok, message)``.
    """
    import json as _json
    import time as _time
    token = get_access_token()
    if not token:
        return False, "No HubSpot token. Set `HUBSPOT_ACCESS_TOKEN` in `.env` or `.streamlit/secrets.toml`."

    try:
        year = int(pinned.get("year") or 0)
        month = int(pinned.get("month") or 0)
    except (TypeError, ValueError):
        return False, "Pinned file is missing a numeric year/month."
    if not (year and 1 <= month <= 12):
        return False, "Pinned file year/month is invalid."

    # Quarter that contains this month
    quarter = (month - 1) // 3 + 1
    # Inclusive month window
    from calendar import monthrange as _mrng
    m_start = datetime(year, month, 1).date()
    m_end = datetime(year, month, _mrng(year, month)[1]).date()

    # Load HubSpot owners directly — this includes deactivated owners (so Lawrence's
    # historical deals still resolve even after his HubSpot account is shut off).
    try:
        from hubspot_service import fetch_owners as _fetch_owners
        owners_list = _fetch_owners(token) or []
    except Exception as e:
        return False, f"Couldn't load HubSpot owners (check token / scope `crm.objects.owners.read`): {e}"
    email_to_full_name: dict[str, str] = {}
    id_to_full_name: dict[str, str] = {}
    id_to_email: dict[str, str] = {}
    for o in owners_list:
        oid = str(o.get("id", "") or "").strip()
        em = (o.get("email") or "").strip().lower()
        fn = f"{(o.get('firstName') or '')} {(o.get('lastName') or '')}".strip()
        if oid:
            id_to_full_name[oid] = fn
            if em:
                id_to_email[oid] = em
        if em:
            email_to_full_name[em] = fn

    # Aliases for short-name reps (pinned files use first names like "Yogi", "Larry").
    _ALIASES = {
        "yogi": ["yogi", "yogesh"],
        "larry": ["larry", "lawrence"],
        "lawrence": ["lawrence", "larry"],
        "rutuja": ["rutuja"],
    }
    pinned_reps = pinned.get("reps") or []

    def _tokens_for(short: str) -> list[str]:
        first = (short.split(" ")[0] or "").strip().lower() if short else ""
        toks = _ALIASES.get(first, [first]) if first else []
        return [t for t in toks if t]

    # Retry the HubSpot call up to 3 times on transient connection / timeout errors.
    # We pass allowed_emails=None so the roster filter doesn't drop deactivated owners.
    mapped = None
    all_closed = None
    last_err: Exception | None = None
    for attempt in range(3):
        try:
            mapped, _stats, all_closed = fetch_and_map_hubspot_deals(
                token,
                allowed_emails=None,
                year=year,
                quarter=int(quarter),
                hubspot_owner_ids=(extra_owner_ids or None),
            )
            break
        except Exception as e:
            last_err = e
            err_text = str(e).lower()
            # Only retry on transient network errors (connection reset, timeout, chunked encoding)
            if any(tok in err_text for tok in ("connection reset", "connectionreset", "timed out", "timeout", "chunked", "max retries", "remote host", "10054")):
                if attempt < 2:
                    _time.sleep(1.5 * (attempt + 1))
                    continue
            break
    if all_closed is None:
        return False, (
            f"HubSpot fetch failed after retries. The network connection was reset. "
            f"This usually means: (1) HubSpot token has expired or been revoked — generate a new Private App token; "
            f"(2) the corporate proxy or firewall is blocking outbound HTTPS to api.hubapi.com; "
            f"(3) you're rate-limited — wait a minute and retry. Underlying error: {last_err}"
        )
    # Prefer the unfiltered all_closed list so deactivated/no-roster owners still come through.
    mapped = all_closed

    # Filter to deals whose close_date falls inside the target month.
    def _in_month(d) -> bool:
        try:
            cd = d.get("close_date")
            if hasattr(cd, "isoformat"):
                cd_d = cd if not hasattr(cd, "date") else (cd.date() if hasattr(cd, "date") else cd)
                # Could be datetime or date
                if hasattr(cd_d, "date"):
                    cd_d = cd_d.date()
            else:
                cd_d = datetime.fromisoformat(str(cd)).date()
            return m_start <= cd_d <= m_end
        except Exception:
            return False

    in_month = [d for d in mapped if _in_month(d)]

    # Per-deal: figure out owner name (works for deactivated owners via id_to_full_name),
    # then match to a pinned rep by first-name token (with aliases).
    def _owner_name_for_deal(d: dict) -> str:
        em = (d.get("deal_owner") or "").strip().lower()
        if em and em in email_to_full_name:
            return email_to_full_name[em].lower()
        # Fallback: use the owner_id carried on the original raw deal if present
        oid = str(d.get("hubspot_owner_id") or d.get("owner_id") or "").strip()
        if oid and oid in id_to_full_name:
            return id_to_full_name[oid].lower()
        return em or ""

    def _match_pinned(name_lower: str, tokens: list[str]) -> bool:
        if not name_lower or not tokens:
            return False
        parts = name_lower.replace(",", " ").split()
        for tok in tokens:
            if not tok:
                continue
            if tok in parts or any(p.startswith(tok) for p in parts):
                return True
            if tok in name_lower:
                return True
        return False

    booked_by_pinned: dict[str, float] = {}
    paid_by_pinned: dict[str, float] = {}
    matched_deals_by_pinned: dict[str, int] = {}
    unmatched_deals: list[str] = []
    deals_by_pinned: dict[str, list[dict]] = {}  # Per-rep deal details for drill-down UI
    for d in in_month:
        owner_name = _owner_name_for_deal(d)
        try:
            booked_amt = float(d.get("amount") or 0)
        except (TypeError, ValueError):
            booked_amt = 0.0
        try:
            paid_amt = float(d.get("paid_amount") or 0)
        except (TypeError, ValueError):
            paid_amt = 0.0

        matched_short = None
        for r in pinned_reps:
            short = (r.get("name") or "").strip()
            if not short:
                continue
            if _match_pinned(owner_name, _tokens_for(short)):
                matched_short = short
                break
        if matched_short:
            booked_by_pinned[matched_short] = booked_by_pinned.get(matched_short, 0.0) + booked_amt
            paid_by_pinned[matched_short] = paid_by_pinned.get(matched_short, 0.0) + paid_amt
            matched_deals_by_pinned[matched_short] = matched_deals_by_pinned.get(matched_short, 0) + 1
            # Save lean deal record for drill-down
            cd_v = d.get("close_date")
            try:
                cd_str = cd_v.isoformat() if hasattr(cd_v, "isoformat") else str(cd_v or "")
            except Exception:
                cd_str = ""
            deals_by_pinned.setdefault(matched_short, []).append(
                {
                    "deal_name": (d.get("deal_name") or "").strip() or "—",
                    "amount": booked_amt,
                    "paid_amount": paid_amt,
                    "close_date": cd_str,
                    "payment_status": (d.get("payment_status_label") or d.get("payment_status") or "—"),
                    "owner_name": owner_name,
                }
            )
        else:
            label = (d.get("deal_name") or "").strip() or "(deal)"
            unmatched_deals.append(f"{label} — owner: {owner_name or '(unknown)'} (${booked_amt:,.0f})")

    # Save deal drill-down details in session_state so the renderer can show them later.
    state_key = f"monthly_deals_{team_prefix}_{year}_{month:02d}"
    st.session_state[state_key] = deals_by_pinned

    # Write totals back into the pinned reps.
    fetched_summary: list[str] = []
    for r in pinned_reps:
        short = (r.get("name") or "").strip()
        booked = booked_by_pinned.get(short, 0.0)
        paid = paid_by_pinned.get(short, 0.0)
        r["achievement_usd"] = round(booked, 2)
        r["payment_received_usd"] = round(paid, 2)
        fetched_summary.append(f"{short}: ${booked:,.0f} booked / ${paid:,.0f} paid ({matched_deals_by_pinned.get(short, 0)} deal(s))")

    # Team aggregates
    pinned["team_achievement_usd"] = round(sum(float(r.get("achievement_usd") or 0) for r in pinned_reps), 2)

    # Persist
    fname = f"{team_prefix}_{_MONTH_ABBR_LONG[month]}_{year}_fixed.json"
    out_path = Path(__file__).parent / "policy" / fname
    try:
        out_path.write_text(_json.dumps(pinned, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError as e:
        return False, f"Wrote totals to memory but couldn't save file: {e}"

    msg = f"Fetched {len(in_month)} deals for {_MONTH_ABBR_LONG[month].title()} {year}. " + " · ".join(fetched_summary)
    if unmatched_deals:
        msg += (
            f"\n\n⚠️ {len(unmatched_deals)} deal(s) didn't match any pinned rep "
            "(owner name didn't match a first-name token): "
            + "; ".join(unmatched_deals[:8])
            + (f"; +{len(unmatched_deals) - 8} more" if len(unmatched_deals) > 8 else "")
        )
    return True, msg


def render_pinned_monthly_team_view(pinned: dict, team_label: str) -> None:
    """Render the Monthly Plan page for SMB / AM April 2026 (or any month with a pinned file).

    Shows: monthly quota table, eligibility/policy summary, per-rep Commission summary with
    Slab %, Achievement %, Eligibility reason, Base commission, Manage Deal (AM only), and
    Total Payout. Also renders a small manager-incentive summary applying the 60% team floor.

    User-level filtering: if the logged-in user is a SALES_REP, the pinned dict is filtered
    to show only their row/deals/commission. ADMIN and SALES_MANAGER see the full team.
    """
    # Filter pinned data to the current user's view (SALES_REPs see only their own row).
    _current_user = st.session_state.get("user") if hasattr(st, "session_state") else None
    _user_matchers = _get_user_rep_name_matchers(_current_user)
    if _user_matchers:
        pinned = _filter_pinned_for_user(pinned, _current_user)
        st.info(
            f"🔒 **Personal view** — showing only your data for {pinned.get('month_label') or 'this month'}. "
            "Team totals are shown for context so you can see how your contribution rolls up."
        )
    tiers = pinned.get("monthly_tiers") or {}
    ind_tiers = tiers.get("individual_tiers") or []
    mgr_tiers = tiers.get("manager_tiers") or []
    min_ach = float(tiers.get("min_achievement_pct_for_commission") or 60)
    mgr_min = float(tiers.get("manager_team_minimum_pct") or 60)
    md_pct = float(tiers.get("manage_deal_pct_am_only") or 5)
    is_am = (team_label or "").upper().startswith("AM") or "ACCOUNT" in (team_label or "").upper()
    comm = pinned.get("commission") or {}
    try:
        inr_rate = float(comm.get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0
    elig_text = (comm.get("eligibility_text") or "").strip()

    label = pinned.get("month_label") or f"{pinned.get('year')}-{pinned.get('month'):02d}"
    team_target = float(pinned.get("team_target_usd") or 0)
    reps = pinned.get("reps") or []

    # Any rep can carry an "exception" block (e.g. Joy's Washington Post server-split).
    # The share_usd value is added to BOTH team achievement and team payment received totals
    # used for slab selection and manager commission basis.
    total_exception_usd = 0.0
    exception_details: list[dict] = []
    for _r_exc in reps:
        _exc = _r_exc.get("exception")
        if _exc:
            try:
                _share = float(_exc.get("joy_share_usd") or _exc.get("share_usd") or 0)
            except (TypeError, ValueError):
                _share = 0.0
            if _share:
                total_exception_usd += _share
                exception_details.append(
                    {
                        "rep": (_r_exc.get("name") or "").strip(),
                        "deal_name": (_exc.get("deal_name") or "").strip(),
                        "deal_amount_usd": float(_exc.get("deal_amount_usd") or 0),
                        "share_usd": _share,
                        "note": (_exc.get("note") or "").strip(),
                    }
                )

    team_achievement_raw = float(pinned.get("team_achievement_usd") or sum(float(r.get("achievement_usd") or 0) for r in reps))
    team_achievement = team_achievement_raw + total_exception_usd
    team_pct = (team_achievement / team_target * 100.0) if team_target else 0.0

    # ---- Quota + summary ----
    st.markdown(f"### {label} — {team_label} Monthly Plan")

    # HubSpot fetch is disabled here because monthly commission cycles have already
    # been processed. Re-fetching from HubSpot would not reflect the correct
    # payment-received cutoff for the month. All numbers come from the pinned JSON
    # file. To re-enable the fetch UI for an open month, set
    # `COMP_TOOL_ENABLE_IMPORTS=1` in `.env`.
    _monthly_imports_enabled = (os.environ.get("COMP_TOOL_ENABLE_IMPORTS") or "").strip() in ("1", "true", "yes")
    _team_prefix = "am" if is_am else "smb"
    _month_lbl = _MONTH_ABBR_LONG[int(pinned.get('month') or 4)].title()

    if _monthly_imports_enabled:
        # Load HubSpot owners (cached) so the user can choose a subset for the fetch.
        _token_for_owners = get_access_token()
        _owner_choices: list[str] = []
        _owner_id_by_label: dict[str, str] = {}
        _default_owner_labels: list[str] = []
        if _token_for_owners:
            try:
                _ck = _hubspot_cache_key(_token_for_owners)
                _owners_cache = _cached_hubspot_owners(_ck, _token_for_owners)
                for _o in (_owners_cache or []):
                    _oid = str(_o.get("id", ""))
                    _em = (_o.get("email") or "").strip()
                    _fn = f"{(_o.get('firstName') or '')} {(_o.get('lastName') or '')}".strip()
                    _archived = bool(_o.get("_archived"))
                    _lbl_core = (f"{_fn} ({_em})" if _em else (_fn or _oid)).strip() or _oid
                    _lbl = f"{_lbl_core} — deactivated" if _archived else _lbl_core
                    _owner_choices.append(_lbl)
                    _owner_id_by_label[_lbl] = _oid
            except Exception:
                pass
            # Pre-select owners whose first-name matches a pinned rep token.
            _rep_tokens = {
                (r.get("name") or "").strip().split(" ")[0].lower()
                for r in (pinned.get("reps") or [])
                if (r.get("name") or "").strip()
            }
            _rep_tokens.discard("")
            # Alias map (short name → full HubSpot label substring)
            _alias = {"yogi": "yogesh", "larry": "lawrence", "lawrence": "lawrence"}
            for lbl in _owner_choices:
                low = lbl.lower()
                for tok in _rep_tokens:
                    target = _alias.get(tok, tok)
                    if target and target in low:
                        _default_owner_labels.append(lbl)
                        break

        st.markdown("**HubSpot fetch**")
        with st.container(border=True):
            if _owner_choices:
                _picked_owners = st.multiselect(
                    "Deal owners (HubSpot)",
                    options=_owner_choices,
                    default=_default_owner_labels,
                    key=f"monthly_owner_pick_{_team_prefix}_{pinned.get('year')}_{pinned.get('month')}",
                    help=(
                        "Restrict the fetch to specific HubSpot owners. Defaults to the owners whose first name matches "
                        "a pinned rep. Clear the list to fetch all owners matching the User Management roster."
                    ),
                )
                _picked_owner_ids = [_owner_id_by_label[l] for l in _picked_owners if l in _owner_id_by_label]
            else:
                _picked_owner_ids = []
                if not _token_for_owners:
                    st.caption("Set `HUBSPOT_ACCESS_TOKEN` in `.env` to enable the owner filter and fetch.")
                else:
                    st.caption("No HubSpot owners loaded (token may be invalid or lacks `crm.objects.owners.read`).")

            bt_col1, bt_col2 = st.columns([1, 3])
            with bt_col1:
                if st.button(
                    f"🔄 Fetch {_month_lbl} from HubSpot",
                    key=f"fetch_monthly_{_team_prefix}_{pinned.get('year')}_{pinned.get('month')}",
                    type="primary",
                    disabled=not _token_for_owners,
                ):
                    with st.spinner("Fetching deals from HubSpot…"):
                        ok, msg = _fetch_monthly_actuals_from_hubspot(
                            pinned,
                            _team_prefix,
                            extra_owner_ids=(_picked_owner_ids or None),
                        )
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
            with bt_col2:
                st.caption(
                    "Pulls closed-won deals from HubSpot whose **Close date** falls inside this month, "
                    "matches owners to the rep roster via User Management email, and updates the pinned JSON in place. "
                    "Up to 3 retries on transient connection errors."
                )

    # KPI cards: show the RAW HubSpot total in the "Achieved" card so the user can see
    # what came back from the fetch. The exception (e.g. Joy's Washington Post split) is
    # broken out in its own section below.
    raw_pct = (team_achievement_raw / team_target * 100.0) if team_target else 0.0
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Monthly target", f"${team_target:,.0f}")
    k2.metric("Achieved (HubSpot)", f"${team_achievement_raw:,.0f}", f"{raw_pct:.1f}%")
    k3.metric("Eligibility floor", f"{int(min_ach)}%")
    k4.metric("Manager team floor", f"{int(mgr_min)}%")
    if elig_text:
        st.caption(f"**Eligibility:** {elig_text}")

    # If any exceptions exist, render a breakdown block right under the KPIs.
    if total_exception_usd:
        rows_exc_html: list[str] = []
        for ex in exception_details:
            ex_rep = html_module.escape(ex.get("rep") or "—")
            ex_deal = html_module.escape(ex.get("deal_name") or "—")
            ex_share = float(ex.get("share_usd") or 0)
            ex_note = html_module.escape(ex.get("note") or "")
            rows_exc_html.append(
                f"<tr><td style='padding:8px 12px;border-bottom:1px solid #bfdbfe;'><strong>{ex_deal}</strong> "
                f"<span style='color:#1e3a8a;font-size:12px;'>→ {ex_rep}</span><br/>"
                f"<span style='font-size:12px;color:#1e3a8a;'>{ex_note}</span></td>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #bfdbfe;text-align:right;color:#1d4ed8;font-weight:500;'>+${ex_share:,.2f}</td></tr>"
            )
        eff_pct = (team_achievement / team_target * 100.0) if team_target else 0.0
        breakdown_html = (
            "<div style='width:100%;border:1px solid #bfdbfe;background:#eff6ff;border-radius:8px;padding:12px 14px;margin:14px 0;'>"
            f"<div style='font-size:14px;font-weight:600;color:#1e3a8a;margin-bottom:8px;'>Adjustments — added to team total</div>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            f"<tbody>{''.join(rows_exc_html)}"
            f"<tr style='background:#dbeafe;font-weight:600;'>"
            f"<td style='padding:10px 12px;'>HubSpot total ${team_achievement_raw:,.2f} + Exception ${total_exception_usd:,.2f}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#1e40af;'>= ${team_achievement:,.2f} ({eff_pct:.1f}%)</td>"
            f"</tr></tbody></table></div>"
        )
        st.markdown(breakdown_html, unsafe_allow_html=True)

    # ---- Deal-count strip + deal-owner / per-owner breakdown (Q1-style, for monthly) ----
    paid_deals_m = pinned.get("paid_deals") or []
    unpaid_deals_m = pinned.get("unpaid_deals_overrides") or []
    deals_summary_m = pinned.get("deals_summary") or {}
    if paid_deals_m or unpaid_deals_m:
        _paid_count_m = len(paid_deals_m)
        _unpaid_count_m = sum(1 for d in unpaid_deals_m if not d.get("is_partial_remainder") and not d.get("is_cancelled"))
        _cancelled_m = [d for d in unpaid_deals_m if d.get("is_cancelled")]
        _total_closed_m = int(deals_summary_m.get("total_deals_closed") or (_paid_count_m + _unpaid_count_m))
        _paid_sum_m = sum(float(d.get("amount_usd") or 0) for d in paid_deals_m)
        _unpaid_sum_m = sum(float(d.get("amount_usd") or 0) for d in unpaid_deals_m if not d.get("is_cancelled"))
        _combined_m = _paid_sum_m + _unpaid_sum_m

        st.markdown(
            "<div style='display:grid;grid-template-columns:repeat(3, 1fr);gap:12px;margin-top:14px;'>"
            "<div style='background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#3730a3;text-transform:uppercase;letter-spacing:0.5px;'>📦 Total deals closed</div>"
            f"<div style='font-size:32px;font-weight:700;color:#1e3a8a;margin-top:4px;'>{_total_closed_m}</div>"
            f"<div style='font-size:12px;color:#475569;margin-top:2px;'>Combined value: ${_combined_m:,.2f}</div>"
            "</div>"
            "<div style='background:#ecfdf5;border:1px solid #86efac;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#166534;text-transform:uppercase;letter-spacing:0.5px;'>✅ Payment received</div>"
            f"<div style='font-size:32px;font-weight:700;color:#14532d;margin-top:4px;'>{_paid_count_m} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
            f"<div style='font-size:12px;color:#166534;margin-top:2px;'>Total: ${_paid_sum_m:,.2f}</div>"
            "</div>"
            "<div style='background:#fef2f2;border:1px solid #fca5a5;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#991b1b;text-transform:uppercase;letter-spacing:0.5px;'>❌ Payment not received</div>"
            f"<div style='font-size:32px;font-weight:700;color:#7f1d1d;margin-top:4px;'>{_unpaid_count_m} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
            f"<div style='font-size:12px;color:#991b1b;margin-top:2px;'>Excluded value: ${_unpaid_sum_m:,.2f}</div>"
            "</div>"
            "</div>",
            unsafe_allow_html=True,
        )

        # ---- Deal owner breakdown table ----
        # Alias map: deal_owner full names → rep short names in reps[]
        _monthly_owner_alias = {
            # SMB reps
            "yogesh vig": "yogi",
            "kritika gupta": "kritika",
            "vicky cariappa": "vicky",
            "deepak r j": "deepak",
            "rutuja kawade": "rutuja",
            "lawrence lewis": "lawrence",
            "royston aden": "royston",
            "kartik kashyap": "kartik",
            # AM reps
            "vivin joseph": "vivin",
            "joy prakash": "joy",
            # Arundhati Sen matches the rep name directly (no alias needed)
        }
        def _monthly_normalize_owner(o: str) -> str:
            base = (o or "").replace(" (Deactivated User)", "").strip().lower()
            return _monthly_owner_alias.get(base, base)

        _rep_lookup_m = {(r.get("name") or "").strip().lower(): r for r in reps}
        m_owners: dict[str, dict] = {}
        for d in paid_deals_m:
            o = m_owners.setdefault(d.get("deal_owner") or "—", {"paid_n": 0, "unpaid_n": 0, "booked_d": 0.0, "paid_d": 0.0, "pending_d": 0.0})
            amt = float(d.get("amount_usd") or 0)
            o["paid_n"] += 1
            o["paid_d"] += amt
            o["booked_d"] += amt
        for d in unpaid_deals_m:
            if d.get("is_cancelled"):
                continue
            o = m_owners.setdefault(d.get("deal_owner") or "—", {"paid_n": 0, "unpaid_n": 0, "booked_d": 0.0, "paid_d": 0.0, "pending_d": 0.0})
            amt = float(d.get("amount_usd") or 0)
            o["unpaid_n"] += 1
            o["booked_d"] += amt
            o["pending_d"] += amt

        st.markdown("#### Deal owner breakdown")
        st.caption(f"Per-rep summary of {team_label} deals closed in {label} — Total deals, Payment received vs not received, Booked Amount, $ received and $ pending. Booked Amount and $ Payment received use the official rep totals from the JSON (rounded to whole dollars).")
        m_owner_rows: list[str] = []
        m_sum_total = m_sum_paid = m_sum_unpaid = 0
        m_sum_booked_off = m_sum_paid_off = 0.0
        m_sum_pending = 0.0
        for owner, o in sorted(m_owners.items(), key=lambda kv: (-(kv[1]["paid_n"] + kv[1]["unpaid_n"]), kv[0])):
            tot = o["paid_n"] + o["unpaid_n"]
            m_sum_total += tot
            m_sum_paid += o["paid_n"]
            m_sum_unpaid += o["unpaid_n"]
            m_sum_pending += o["pending_d"]
            rep_key = _monthly_normalize_owner(owner)
            rep_entry = _rep_lookup_m.get(rep_key)
            if rep_entry:
                booked_disp = round(float(rep_entry.get("achievement_usd") or 0))
                paid_disp = float(rep_entry.get("payment_received_usd") or 0)
            else:
                booked_disp = o["booked_d"]
                paid_disp = o["paid_d"]
            m_sum_booked_off += booked_disp
            m_sum_paid_off += paid_disp
            m_owner_rows.append(
                "<tr>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;'><strong>{html_module.escape(owner)}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{tot}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#166534;'><strong>{o['paid_n']}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#991b1b;'><strong>{o['unpaid_n']}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${booked_disp:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#166534;'>${paid_disp:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#991b1b;'>${o['pending_d']:,.2f}</td>"
                "</tr>"
            )
        m_official_booked = float(pinned.get("team_achievement_usd") or m_sum_booked_off)
        m_official_paid = m_sum_paid_off
        m_owner_rows.append(
            "<tr style='background:#fef9c3;font-weight:700;'>"
            f"<td style='padding:10px 12px;'>Total ({team_label} team)</td>"
            f"<td style='padding:10px 12px;text-align:right;'>{m_sum_total}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#166534;'>{m_sum_paid}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#991b1b;'>{m_sum_unpaid}</td>"
            f"<td style='padding:10px 12px;text-align:right;'>${m_official_booked:,.2f}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#14532d;'>${m_official_paid:,.2f}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#7f1d1d;'>${m_sum_pending:,.2f}</td>"
            "</tr>"
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-top:6px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            "<thead><tr style='background:#e0f2fe;color:#075985;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Deal Owner</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Total deals</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Payment received deals</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'># Payment not received</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Booked Amount</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>$ Payment received</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>$ Payment Pending</th>"
            f"</tr></thead><tbody>{''.join(m_owner_rows)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

        # ---- Per-owner expander with Excel export ----
        with st.expander(f"📋 See every {team_label} deal grouped by owner", expanded=False):
            m_own_paid: dict[str, list[dict]] = {}
            m_own_unpaid: dict[str, list[dict]] = {}
            for d in paid_deals_m:
                m_own_paid.setdefault(d.get("deal_owner") or "—", []).append(d)
            for d in unpaid_deals_m:
                m_own_unpaid.setdefault(d.get("deal_owner") or "—", []).append(d)

            import pandas as _pd
            _m_xls_rows: list[dict] = []
            for owner in sorted(set(list(m_own_paid.keys()) + list(m_own_unpaid.keys()))):
                for d in m_own_paid.get(owner, []):
                    _m_xls_rows.append({
                        "Deal Owner": owner,
                        "Deal": d.get("deal_name") or "",
                        "Close date": d.get("close_date") or "",
                        "Status": d.get("payment_status") or "Paid",
                        "Amount (USD)": float(d.get("amount_usd") or 0),
                        "Payment Receipt Amount (USD)": float(d.get("amount_usd") or 0),
                        "Pending (USD)": 0.0,
                    })
                for d in m_own_unpaid.get(owner, []):
                    _m_xls_rows.append({
                        "Deal Owner": owner,
                        "Deal": d.get("deal_name") or "",
                        "Close date": d.get("close_date") or "",
                        "Status": d.get("payment_status") or "Not Paid",
                        "Amount (USD)": float(d.get("amount_usd") or 0),
                        "Payment Receipt Amount (USD)": 0.0,
                        "Pending (USD)": float(d.get("amount_usd") or 0),
                    })
            try:
                _m_xls = _pinned_export_to_excel_bytes({f"All {team_label} deals by owner": _pd.DataFrame(_m_xls_rows)})
                _team_short_dl = (team_label or "team").lower().replace(" ", "_")
                _mo_int = int(pinned.get("month") or 4)
                _yr_int = int(pinned.get("year") or 2026)
                st.download_button(
                    label=f"📥 Export all {team_label} deals grouped by owner (Excel)",
                    data=_m_xls,
                    file_name=f"{_team_short_dl}_{_yr_int}_{_mo_int:02d}_deals_by_owner.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_monthly_deals_by_owner_{_team_short_dl}_{_yr_int}_{_mo_int:02d}",
                )
            except Exception:
                pass

            _grand_amt_m = 0.0
            _grand_paid_m = 0.0
            _grand_n_m = 0
            for owner in sorted(set(list(m_own_paid.keys()) + list(m_own_unpaid.keys()))):
                p_list = m_own_paid.get(owner, [])
                u_list = m_own_unpaid.get(owner, [])
                p_sum = sum(float(d.get("amount_usd") or 0) for d in p_list)
                u_sum = sum(float(d.get("amount_usd") or 0) for d in u_list if not d.get("is_cancelled"))
                deal_n = len(p_list) + sum(1 for d in u_list if not d.get("is_cancelled"))
                _grand_amt_m += p_sum + u_sum
                _grand_paid_m += p_sum
                _grand_n_m += deal_n
                _received_chip = f"<span style='color:#166534;font-size:13px;font-weight:500;'>· ${p_sum:,.2f} received</span>"
                _not_received_chip = (
                    f" <span style='color:#991b1b;font-size:13px;font-weight:500;'>· ${u_sum:,.2f} not received</span>"
                    if u_sum > 0 else ""
                )
                st.markdown(
                    f"<div style='margin-top:14px;margin-bottom:6px;font-size:18px;font-weight:700;color:#0f172a;'>"
                    f"{html_module.escape(owner)} — {deal_n} deal(s) {_received_chip}{_not_received_chip}"
                    "</div>",
                    unsafe_allow_html=True,
                )
                rows_h_m: list[str] = []
                for d in p_list:
                    dn = html_module.escape((d.get("deal_name") or "").strip())
                    amt = float(d.get("amount_usd") or 0)
                    cd = html_module.escape((d.get("close_date") or "").strip())
                    rows_h_m.append(
                        f"<tr style='background:#f0fdf4;'>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;'><strong>{dn}</strong></td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;'>{cd}</td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;color:#166534;'>Paid</td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;'><strong>${amt:,.2f}</strong></td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;color:#166534;'><strong>${amt:,.2f}</strong></td>"
                        "</tr>"
                    )
                for d in u_list:
                    dn = html_module.escape((d.get("deal_name") or "").strip())
                    amt = float(d.get("amount_usd") or 0)
                    cd = html_module.escape((d.get("close_date") or "").strip())
                    status = html_module.escape((d.get("payment_status") or "Not Paid").strip())
                    rows_h_m.append(
                        f"<tr style='background:#fef2f2;'>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;'><strong>{dn}</strong></td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;text-align:right;'>{cd}</td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;text-align:right;color:#b91c1c;'>{status}</td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;text-align:right;'><strong>${amt:,.2f}</strong></td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;text-align:right;color:#b91c1c;'><strong>$0.00</strong></td>"
                        "</tr>"
                    )
                rows_h_m.append(
                    "<tr style='background:#fef9c3;font-weight:700;'>"
                    f"<td style='padding:10px 12px;' colspan='2'>Total — {html_module.escape(owner)} ({deal_n} deal{'s' if deal_n != 1 else ''})</td>"
                    f"<td style='padding:10px 12px;text-align:right;'></td>"
                    f"<td style='padding:10px 12px;text-align:right;'>${(p_sum + u_sum):,.2f}</td>"
                    f"<td style='padding:10px 12px;text-align:right;color:#166534;'>${p_sum:,.2f}</td>"
                    "</tr>"
                )
                st.markdown(
                    "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:14px;'>"
                    "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
                    "<thead><tr style='background:#f1f5f9;color:#334155;'>"
                    "<th style='text-align:left;padding:8px 12px;font-weight:600;'>Deal</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Close date</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Status</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Amount</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Payment Receipt Amount</th>"
                    f"</tr></thead><tbody>{''.join(rows_h_m)}</tbody></table></div>",
                    unsafe_allow_html=True,
                )

            _grand_pending_m = _grand_amt_m - _grand_paid_m
            _official_booked_grand = float(pinned.get("team_achievement_usd") or _grand_amt_m)
            _official_paid_grand = m_sum_paid_off if m_sum_paid_off else _grand_paid_m
            st.markdown(
                "<div style='margin-top:6px;padding:14px 16px;background:#fef3c7;border:2px solid #f59e0b;border-radius:10px;'>"
                f"<div style='font-size:13px;font-weight:700;color:#78350f;text-transform:uppercase;letter-spacing:0.5px;'>🏷️ Grand total — {team_label} team (all owners)</div>"
                "<div style='display:grid;grid-template-columns:repeat(4, 1fr);gap:12px;margin-top:8px;'>"
                f"<div><div style='font-size:11px;color:#92400e;'>Total deals</div><div style='font-size:20px;font-weight:700;color:#78350f;'>{_grand_n_m}</div></div>"
                f"<div><div style='font-size:11px;color:#92400e;'>Total Amount</div><div style='font-size:20px;font-weight:700;color:#78350f;'>${_official_booked_grand:,.2f}</div></div>"
                f"<div><div style='font-size:11px;color:#92400e;'>Payment Receipt</div><div style='font-size:20px;font-weight:700;color:#166534;'>${_official_paid_grand:,.2f}</div></div>"
                f"<div><div style='font-size:11px;color:#92400e;'>Pending</div><div style='font-size:20px;font-weight:700;color:#991b1b;'>${_grand_pending_m:,.2f}</div></div>"
                "</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # ---- Commission summary ----
    st.markdown("#### Commission summary")
    st.caption("Commission is calculated on **Payment Received** (cash collected), not on booked revenue. Achievement % uses booked vs quota for tier selection.")
    money_cols = {"Quota", "Revenue Achieved", "Payment Received", "Base Commission", "Manage Deal", "Adjustments", "Total Payout"}
    inr_cols = {"Total Payout (INR)"}
    rows_html: list[str] = []

    base_total = 0.0
    md_total = 0.0
    adj_grand_total = 0.0
    payout_total = 0.0

    columns = [
        "Rep. Name",
        "Quota",
        "Revenue Achieved",
        "Payment Received",
        "Achieved %",
        "Slab %",
        "Eligibility",
        "Base Commission",
        "Manage Deal",
        "Adjustments",
        "Total Payout",
        "Exchange Rate",
        "Total Payout (INR)",
        "Eligibility reason",
    ]
    widths = {
        "Rep. Name": "150px",
        "Quota": "110px",
        "Revenue Achieved": "140px",
        "Payment Received": "140px",
        "Achieved %": "100px",
        "Slab %": "110px",
        "Eligibility": "110px",
        "Base Commission": "140px",
        "Manage Deal": "120px",
        "Adjustments": "120px",
        "Total Payout": "140px",
        "Exchange Rate": "120px",
        "Total Payout (INR)": "160px",
        "Eligibility reason": "440px",
    }
    head = "".join(
        f'<th style="text-align:left;padding:8px 12px;border-bottom:1px solid #e5e7eb;background:#ccfbf1;font-weight:500;color:#134e4a;min-width:{widths[c]};white-space:nowrap;">{html_module.escape(c)}</th>'
        for c in columns
    )

    for r in reps:
        # Skip managers from the rep commission summary (they appear in the manager
        # incentive section instead). They still count toward team bullet bars + totals.
        if bool(r.get("is_manager")):
            continue
        nm = (r.get("name") or "").strip() or "—"
        quota = float(r.get("target_usd") or 0)
        ach = float(r.get("achievement_usd") or 0)
        paid = float(r.get("payment_received_usd") or 0)
        manage_deal_usd = float(r.get("manage_deal_usd") or 0)
        manage_deal_count = float(r.get("manage_deal_count") or 0)
        left_org = bool(r.get("left_org"))
        pct = (ach / quota * 100.0) if quota else 0.0
        slab_pct, slab_label = _monthly_slab_pct(ind_tiers, pct)
        eligible = pct >= min_ach
        # left_org reps get no payout regardless of achievement.
        if left_org:
            eligible = False
            elig_status = "Left org"
            elig_color = "#6b7280"
            # Build a rep- + month-specific note, e.g. "Lawrence left the organization in April."
            _first_name = (nm.split() or [nm])[0]
            _month_word = ""
            try:
                _month_word = (pinned.get("month_label") or "").split()[0]
            except Exception:
                _month_word = ""
            if _month_word:
                reason = (
                    f"{_first_name} left the organization in {_month_word}. "
                    "No payout — their deals still count toward team total."
                )
            else:
                reason = f"{_first_name} left the organization. No payout — their deals still count toward team total."
        else:
            elig_status = "Eligible" if eligible else "Not eligible"
            elig_color = "#15803d" if eligible else "#b91c1c"
            if not eligible:
                reason = f"Below {int(min_ach)}% threshold — achievement {pct:.0f}%."
            else:
                reason = ""

        # Base commission is calculated on PAYMENT RECEIVED (not booked revenue).
        base_comm = round(paid * slab_pct / 100.0, 2) if eligible else 0.0
        if is_am and manage_deal_usd > 0:
            md_amount = round(manage_deal_usd * md_pct / 100.0, 2)
        elif is_am and manage_deal_count > 0:
            md_amount = round(manage_deal_count * 500, 2)
        else:
            md_amount = 0.0
        # Prior-period adjustments (e.g. Yogi's Q4 Yieldstreet) — always paid, regardless of April eligibility.
        adj_rows = r.get("adjustments") or []
        adj_total = round(sum(float(a.get("commission_usd") or 0) for a in adj_rows), 2)
        if adj_rows:
            extras = ", ".join(
                f"{(a.get('deal_name') or '').strip()} ({(a.get('period') or '').strip()}) +${float(a.get('commission_usd') or 0):,.0f}"
                for a in adj_rows
            )
            reason = (reason + (" · " if reason else "") + f"Adjustments: {extras}") if (reason or extras) else reason
        total_payout = round(base_comm + md_amount + adj_total, 2)
        # Per-rep total override (pinned exact figure, e.g. Yogi's $1,004 rounded total).
        _payout_override = r.get("total_payout_override_usd")
        if _payout_override is not None:
            try:
                total_payout = round(float(_payout_override), 2)
            except (TypeError, ValueError):
                pass
        payout_inr = round(total_payout * inr_rate, 2) if inr_rate else 0.0

        base_total += base_comm
        md_total += md_amount
        adj_grand_total += adj_total
        payout_total += total_payout

        cells = {
            "Rep. Name": nm,
            "Quota": quota,
            "Revenue Achieved": ach,
            "Payment Received": paid,
            "Achieved %": f"{pct:.0f}%",
            "Slab %": f"{slab_pct:.0f}%" if slab_pct else "0%",
            "Eligibility": f'<span style="color:{elig_color};font-weight:500;">{elig_status}</span>',
            "Base Commission": base_comm,
            "Manage Deal": md_amount,
            "Adjustments": adj_total,
            "Total Payout": total_payout,
            "Exchange Rate": f"₹{inr_rate:g}/USD" if inr_rate else "—",
            "Total Payout (INR)": payout_inr,
            "Eligibility reason": reason,
        }
        row_cells_html = []
        for c in columns:
            v = cells[c]
            if c in money_cols and isinstance(v, (int, float)):
                row_cells_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">${float(v):,.2f}</td>')
            elif c in inr_cols and isinstance(v, (int, float)):
                row_cells_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">₹{float(v):,.2f}</td>')
            elif c == "Eligibility":
                row_cells_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">{v}</td>')
            else:
                row_cells_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">{html_module.escape(str(v if v is not None else ""))}</td>')
        rows_html.append("<tr>" + "".join(row_cells_html) + "</tr>")

    table_html = (
        '<div style="width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;">'
        '<table style="border-collapse:collapse;font-size:13px;width:max-content;min-width:100%;">'
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)

    # ---- Footer totals ----
    payout_total_inr = round(payout_total * inr_rate, 2) if inr_rate else 0.0
    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Base commission (sum)", f"${base_total:,.2f}")
    f2.metric("Manage deal (sum)", f"${md_total:,.2f}")
    f3.metric("Total payout (USD)", f"${payout_total:,.2f}")
    if inr_rate:
        f4.metric(f"Total payout (INR @ ₹{inr_rate:g})", f"₹{payout_total_inr:,.2f}")

    # NB: Manager incentive lives on the dedicated Manager incentive tab — no duplicate
    # render at the bottom of the team view.

    if any(float(r.get("achievement_usd") or 0) == 0 for r in reps):
        st.caption(
            "⚠️ Some achievements are still at $0 — fill them in `policy/{prefix}_{month}_{year}_fixed.json` "
            "(or wire a HubSpot fetch) to compute real commissions.".format(
                prefix=("am" if is_am else "smb"),
                month=_MONTH_ABBR_LONG.get(int(pinned.get("month") or 0), "april"),
                year=int(pinned.get("year") or 2026),
            )
        )

    # ---- Per-rep deal drill-down (uses deals saved by the HubSpot fetch) ----
    state_key = f"monthly_deals_{_team_prefix}_{int(pinned.get('year') or 0)}_{int(pinned.get('month') or 0):02d}"
    deals_by_pinned: dict[str, list[dict]] = st.session_state.get(state_key) or {}
    st.markdown("---")
    st.markdown("#### Deal details")
    if not deals_by_pinned:
        st.caption("Click **🔄 Fetch from HubSpot** above to populate per-rep deal details for drill-down.")
    else:
        rep_options = [
            (r.get("name") or "").strip()
            for r in reps
            if (r.get("name") or "").strip() in deals_by_pinned
        ]
        if not rep_options:
            st.caption("No matched deals to drill into.")
        else:
            picked = st.selectbox(
                "View deals for…",
                options=["— select rep —"] + rep_options,
                key=f"deal_drilldown_pick_{_team_prefix}_{pinned.get('year')}_{pinned.get('month')}",
            )
            if picked and picked != "— select rep —":
                rep_deals = deals_by_pinned.get(picked, [])
                if not rep_deals:
                    st.info(f"No deals found for {picked} in this month.")
                else:
                    booked_sum = sum(float(d.get("amount") or 0) for d in rep_deals)
                    paid_sum = sum(float(d.get("paid_amount") or 0) for d in rep_deals)
                    d1, d2, d3 = st.columns(3)
                    d1.metric("Deal count", str(len(rep_deals)))
                    d2.metric("Total booked", f"${booked_sum:,.0f}")
                    d3.metric("Total paid", f"${paid_sum:,.0f}")
                    deals_df = pd.DataFrame(
                        [
                            {
                                "Deal Name": d.get("deal_name", "—"),
                                "Amount": float(d.get("amount") or 0),
                                "Payment Received": float(d.get("paid_amount") or 0),
                                "Payment Status": d.get("payment_status", "—"),
                                "Close Date": d.get("close_date", ""),
                            }
                            for d in rep_deals
                        ]
                    )
                    st.dataframe(
                        deals_df,
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            "Amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
                            "Payment Received": st.column_config.NumberColumn("Payment Received", format="$%.2f"),
                        },
                    )

    # Inline chat assistant for the monthly view (Claude-powered when API key set).
    # Session key differentiates SMB / AM by team_label so the two pages don't share history.
    _chat_prefix = "am_monthly" if (team_label or "").lower().startswith("a") else "smb_monthly"
    render_pinned_chat_assistant(pinned, session_prefix=_chat_prefix)


def render_pinned_am_quarterly_view(pinned: dict) -> None:
    """Render the AM Q1 2026 (or any pinned AM quarter) Sales Target + Commission summary view.

    User-level filtering: SALES_REPs see only their own row/deals/commission.
    ADMIN and SALES_MANAGER see the full team.
    """
    _current_user = st.session_state.get("user") if hasattr(st, "session_state") else None
    _user_matchers = _get_user_rep_name_matchers(_current_user)
    if _user_matchers:
        pinned = _filter_pinned_for_user(pinned, _current_user)
        st.info(
            f"🔒 **Personal view** — showing only your data for Q{pinned.get('quarter')} FY{pinned.get('year')}. "
            "Team totals are shown for context."
        )
    year = pinned.get("year")
    quarter = pinned.get("quarter")
    team_target = float(pinned.get("team_target_usd") or 0)
    reps_top = pinned.get("reps") or []
    team_achievement = float(pinned.get("team_achievement_usd") or sum(float(r.get("achievement_usd") or 0) for r in reps_top))
    team_pct = (team_achievement / team_target * 100.0) if team_target else 0.0
    comm = pinned.get("commission") or {}
    try:
        inr_rate = float(comm.get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0
    try:
        ded_pct = float(comm.get("team_deduction_pct") or 10)
    except (TypeError, ValueError):
        ded_pct = 10.0
    elig_text = (comm.get("eligibility_text") or "").strip()
    ded_reason = (comm.get("team_deduction_reason") or "").strip()

    # ---- KPI cards ----
    st.markdown(f"### Q{quarter} {year} — Account Management Sales Target")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Team target", f"${team_target:,.0f}")
    k2.metric("Achieved (booked)", f"${team_achievement:,.0f}", f"{team_pct:.1f}%")
    team_paid_sum = sum(float(r.get("payment_received_usd") or 0) for r in reps_top)
    k3.metric("Payment received", f"${team_paid_sum:,.0f}")
    k4.metric("Eligibility floor", f"{int(comm.get('eligibility_min_pct') or 50)}%")
    if elig_text:
        st.caption(f"**Eligibility:** {elig_text}")
    if ded_reason:
        st.caption(f"**Deduction reason:** {ded_reason}")

    st.markdown("---")

    # ---- Performance vs goal (bullet bars) — moved to top of page to match SMB layout ----
    perf_parts: list[str] = [
        '<div class="goal-attainment-wrap">',
        '<div class="ga-bullet-section">',
        '<p class="ga-bullet-section-title">Performance vs goal</p>',
        _bullet_chart_row_html(
            "Account Management team",
            f"Q{quarter} {year} total",
            team_achievement,
            team_target,
        ),
    ]
    for _r in reps_top:
        _nm = (_r.get("name") or "").strip()
        _ach = float(_r.get("achievement_usd") or 0)
        _tgt = float(_r.get("target_usd") or 0)
        _av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar(_nm))}</div>'
        perf_parts.append(_bullet_chart_row_html(_nm, "Individual quota", _ach, _tgt, initials_html=_av))
    perf_parts.extend(["</div>", "</div>"])
    st.markdown("".join(perf_parts), unsafe_allow_html=True)

    st.markdown("---")

    # ---- Deal-count strip + deal-owner / rep-level breakdowns ----
    paid_deals_am = pinned.get("paid_deals") or []
    unpaid_deals_am = pinned.get("unpaid_deals_overrides") or []
    deals_summary_am = pinned.get("deals_summary") or {}
    if paid_deals_am or unpaid_deals_am:
        # Build unique-deal sets (Artnet appears in both arrays as partial paid + partial remainder).
        partial_keys = {d.get("deal_name") for d in paid_deals_am if d.get("is_partial")}
        partial_remainder_keys = {d.get("deal_name") for d in unpaid_deals_am if d.get("is_partial_remainder")}
        # Count fully-paid + partial-paid (any payment received)
        any_payment_count = len(paid_deals_am)  # includes the partial-paid Artnet
        # Count fully-not-paid (excluding the partial remainder for Artnet)
        fully_unpaid_count = sum(1 for d in unpaid_deals_am if not d.get("is_partial_remainder"))
        # Total deals closed
        partial_paid_count = len(partial_keys)
        total_closed = (any_payment_count - partial_paid_count) + fully_unpaid_count + partial_paid_count
        total_closed = int(deals_summary_am.get("total_deals_closed") or total_closed)

        paid_sum_am = sum(float(d.get("amount_usd") or 0) for d in paid_deals_am)
        unpaid_sum_am = sum(float(d.get("amount_usd") or 0) for d in unpaid_deals_am)
        combined_value = paid_sum_am + unpaid_sum_am

        st.markdown(
            "<div style='display:grid;grid-template-columns:repeat(3, 1fr);gap:12px;margin-top:14px;'>"
            # Card 1 — Total closed
            "<div style='background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#3730a3;text-transform:uppercase;letter-spacing:0.5px;'>📦 Total deals closed</div>"
            f"<div style='font-size:32px;font-weight:700;color:#1e3a8a;margin-top:4px;'>{total_closed}</div>"
            f"<div style='font-size:12px;color:#475569;margin-top:2px;'>Combined value: ${combined_value:,.2f}</div>"
            "</div>"
            # Card 2 — Payment received
            "<div style='background:#ecfdf5;border:1px solid #86efac;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#166534;text-transform:uppercase;letter-spacing:0.5px;'>✅ Payment received</div>"
            f"<div style='font-size:32px;font-weight:700;color:#14532d;margin-top:4px;'>{any_payment_count} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
            f"<div style='font-size:12px;color:#166534;margin-top:2px;'>Total: ${paid_sum_am:,.2f}"
            + (f" · incl. {partial_paid_count} partially paid" if partial_paid_count else "")
            + "</div>"
            "</div>"
            # Card 3 — Payment NOT received
            "<div style='background:#fef2f2;border:1px solid #fca5a5;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#991b1b;text-transform:uppercase;letter-spacing:0.5px;'>❌ Payment not received</div>"
            f"<div style='font-size:32px;font-weight:700;color:#7f1d1d;margin-top:4px;'>{fully_unpaid_count} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
            f"<div style='font-size:12px;color:#991b1b;margin-top:2px;'>Excluded value: ${unpaid_sum_am:,.2f}</div>"
            "</div>"
            "</div>",
            unsafe_allow_html=True,
        )

        # ---- Deal owner breakdown ----
        # Aggregate by deal_owner. For each owner we compute:
        #   - paid_count: deals with payment received (incl. partial-paid)
        #   - partial_paid_count: count of partial-paid deals (annotation only)
        #   - unpaid_count: deals with NO payment received (excl. partial-remainder)
        #   - partial_unpaid_count: partial-remainder count (annotation only)
        #   - booked_amount: sum of total deal values (partial deals counted as full value)
        #   - paid_usd: payment received
        #   - pending_usd: payment pending
        owners: dict[str, dict] = {}
        def _bump_paid(owner: str, total_value: float, paid_value: float, is_partial: bool):
            o = owners.setdefault(owner or "—", {
                "paid_count": 0, "unpaid_count": 0,
                "partial_paid_count": 0, "partial_unpaid_count": 0,
                "booked_amount": 0.0, "paid_usd": 0.0, "pending_usd": 0.0,
            })
            o["paid_count"] += 1
            o["paid_usd"] += paid_value
            o["booked_amount"] += total_value
            if is_partial:
                o["partial_paid_count"] += 1
                # The unpaid portion of a partial paid deal contributes to pending
                o["pending_usd"] += max(total_value - paid_value, 0)

        def _bump_unpaid(owner: str, amount: float, is_partial_remainder: bool):
            o = owners.setdefault(owner or "—", {
                "paid_count": 0, "unpaid_count": 0,
                "partial_paid_count": 0, "partial_unpaid_count": 0,
                "booked_amount": 0.0, "paid_usd": 0.0, "pending_usd": 0.0,
            })
            if is_partial_remainder:
                # Don't double-count booked or pending — the partial-paid side already booked the full value
                # and added the pending portion. Just record the annotation count.
                o["partial_unpaid_count"] += 1
            else:
                o["unpaid_count"] += 1
                o["booked_amount"] += amount
                o["pending_usd"] += amount

        for d in paid_deals_am:
            is_partial = bool(d.get("is_partial"))
            total_v = float(d.get("total_deal_value_usd") or d.get("amount_usd") or 0) if is_partial else float(d.get("amount_usd") or 0)
            _bump_paid(d.get("deal_owner") or "", total_v, float(d.get("amount_usd") or 0), is_partial)
        for d in unpaid_deals_am:
            _bump_unpaid(d.get("deal_owner") or "", float(d.get("amount_usd") or 0), bool(d.get("is_partial_remainder")))

        st.markdown("#### Deal owner breakdown")
        st.caption("Per-rep summary of deals closed in Q1 — Total deals, Payment received vs not received, Booked Amount, $ received and $ pending.")
        owner_rows: list[str] = []
        sum_total = 0
        sum_paid = 0
        sum_unpaid = 0
        sum_partial_paid = 0
        sum_partial_unpaid = 0
        sum_booked = 0.0
        sum_paid_usd = 0.0
        sum_pending_usd = 0.0
        # Sort by total deal count desc, then by name
        for owner, o in sorted(owners.items(), key=lambda kv: (-(kv[1]["paid_count"] + kv[1]["unpaid_count"]), kv[0])):
            paid_ct = o["paid_count"]
            unpaid_ct = o["unpaid_count"]
            partial_paid_ct = o["partial_paid_count"]
            partial_unpaid_ct = o["partial_unpaid_count"]
            total_ct = paid_ct + unpaid_ct
            sum_total += total_ct
            sum_paid += paid_ct
            sum_unpaid += unpaid_ct
            sum_partial_paid += partial_paid_ct
            sum_partial_unpaid += partial_unpaid_ct
            sum_booked += o["booked_amount"]
            sum_paid_usd += o["paid_usd"]
            sum_pending_usd += o["pending_usd"]
            # Format the count cells with the partial annotation in line
            paid_cell = f"{paid_ct}" + (f" <span style='color:#ca8a04;font-size:11px;font-weight:500;'>({partial_paid_ct}-Partial Paid)</span>" if partial_paid_ct else "")
            unpaid_cell = f"{unpaid_ct}" + (f" <span style='color:#ca8a04;font-size:11px;font-weight:500;'>({partial_unpaid_ct}-Partial Payment not received)</span>" if partial_unpaid_ct else "")
            owner_rows.append(
                "<tr>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;'><strong>{html_module.escape(owner)}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{total_ct}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#166534;'><strong>{paid_cell}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#991b1b;'><strong>{unpaid_cell}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${o['booked_amount']:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#166534;'>${o['paid_usd']:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#991b1b;'>${o['pending_usd']:,.2f}</td>"
                "</tr>"
            )
        # Total row
        total_paid_cell = f"{sum_paid}" + (f" <span style='color:#ca8a04;font-size:11px;font-weight:500;'>({sum_partial_paid}-Partial Paid)</span>" if sum_partial_paid else "")
        total_unpaid_cell = f"{sum_unpaid}" + (f" <span style='color:#ca8a04;font-size:11px;font-weight:500;'>({sum_partial_unpaid}-Partial Payment not received)</span>" if sum_partial_unpaid else "")
        owner_rows.append(
            "<tr style='background:#fef9c3;font-weight:700;'>"
            f"<td style='padding:10px 12px;'>Total (AM team)</td>"
            f"<td style='padding:10px 12px;text-align:right;'>{sum_total}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#166534;'>{total_paid_cell}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#991b1b;'>{total_unpaid_cell}</td>"
            f"<td style='padding:10px 12px;text-align:right;'>${sum_booked:,.2f}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#14532d;'>${sum_paid_usd:,.2f}</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#7f1d1d;'>${sum_pending_usd:,.2f}</td>"
            "</tr>"
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-top:6px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            "<thead><tr style='background:#e0f2fe;color:#075985;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Deal Owner</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Total deals</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Payment received deals</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'># Payment not received</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Booked Amount</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>$ Payment received</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>$ Payment Pending</th>"
            f"</tr></thead><tbody>{''.join(owner_rows)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

        # ---- Per-owner deal lists (expander per rep) ----
        official_team_paid = float((pinned.get("manager_incentive") or {}).get("team_payment_received_usd") or paid_sum_am)

        # Build a *unified, deduplicated* per-owner deal list. Artnet appears in both
        # paid_deals (partial-paid portion) and unpaid_deals_overrides (partial remainder);
        # we merge those into a single row showing Amount (total deal value) and
        # Payment Receipt Amount (the paid portion).
        def _normalize_deal_key(name: str) -> str:
            # Strip the "(paid portion)" / "(unpaid remainder)" suffixes used for partial
            # deals so they collapse to the same key.
            s = (name or "").strip().lower()
            for tag in (" (paid portion)", " (unpaid remainder)", " (paid)", " (unpaid)"):
                if s.endswith(tag):
                    s = s[: -len(tag)].strip()
            return s

        unified_by_owner: dict[str, list[dict]] = {}
        # First pass — paid deals (gives us each deal's paid portion)
        for d in paid_deals_am:
            owner = d.get("deal_owner") or "—"
            key = _normalize_deal_key(d.get("deal_name") or "")
            is_partial = bool(d.get("is_partial"))
            amount = float(d.get("total_deal_value_usd") or d.get("amount_usd") or 0) if is_partial else float(d.get("amount_usd") or 0)
            paid = float(d.get("amount_usd") or 0)
            # Display name without the "(paid portion)" suffix when we're merging
            display_name = (d.get("deal_name") or "").replace(" (paid portion)", "").replace(" (paid)", "").strip()
            unified_by_owner.setdefault(owner, []).append({
                "_key": key,
                "deal_name": display_name,
                "close_date": d.get("close_date") or "",
                "status": d.get("payment_status") or "Paid",
                "amount_usd": amount,
                "paid_amount_usd": paid,
                "is_partial": is_partial,
            })
        # Second pass — unpaid deals. If the same normalized key already exists, merge;
        # otherwise add as a new row (Not Paid → amount=full value, paid=$0).
        for d in unpaid_deals_am:
            owner = d.get("deal_owner") or "—"
            key = _normalize_deal_key(d.get("deal_name") or "")
            is_partial_rem = bool(d.get("is_partial_remainder"))
            owner_rows = unified_by_owner.setdefault(owner, [])
            existing = next((r for r in owner_rows if r["_key"] == key), None)
            if existing and is_partial_rem:
                # Merge into the existing partial-paid row. Use the full deal value if known,
                # else paid + unpaid_remainder.
                full_value = float(d.get("total_deal_value_usd") or 0) or (existing["paid_amount_usd"] + float(d.get("amount_usd") or 0))
                existing["amount_usd"] = full_value
                existing["status"] = "Partially Paid"
            else:
                display_name = (d.get("deal_name") or "").replace(" (unpaid remainder)", "").replace(" (unpaid)", "").strip()
                owner_rows.append({
                    "_key": key,
                    "deal_name": display_name,
                    "close_date": d.get("close_date") or "",
                    "status": d.get("payment_status") or "Not Paid",
                    "amount_usd": float(d.get("amount_usd") or 0),
                    "paid_amount_usd": 0.0,
                    "is_partial": False,
                })

        with st.expander("📋 See every deal grouped by owner", expanded=False):
            # ---- Excel export (all owners, unified rows) ----
            import pandas as _pd
            _all_deal_rows: list[dict] = []
            for owner in sorted(unified_by_owner.keys()):
                for r in unified_by_owner[owner]:
                    _all_deal_rows.append({
                        "Deal Owner": owner,
                        "Deal": r["deal_name"],
                        "Close date": r["close_date"],
                        "Status": r["status"],
                        "Amount (USD)": r["amount_usd"],
                        "Payment Receipt Amount (USD)": r["paid_amount_usd"],
                        "Pending (USD)": max(r["amount_usd"] - r["paid_amount_usd"], 0),
                    })
            _df_all = _pd.DataFrame(_all_deal_rows)
            try:
                _xls_bytes = _pinned_export_to_excel_bytes({"All deals by owner": _df_all})
                st.download_button(
                    label="📥 Export all deals grouped by owner (Excel)",
                    data=_xls_bytes,
                    file_name=f"am_q1_{pinned.get('year') or 2026}_deals_by_owner.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_am_deals_by_owner_xlsx",
                    use_container_width=False,
                )
            except Exception:
                pass

            grand_amount_sum = 0.0
            grand_paid_sum = 0.0
            grand_unique = 0
            for owner in sorted(unified_by_owner.keys()):
                rows = unified_by_owner[owner]
                owner_amount = sum(r["amount_usd"] for r in rows)
                owner_paid = sum(r["paid_amount_usd"] for r in rows)
                owner_pending = owner_amount - owner_paid
                grand_amount_sum += owner_amount
                grand_paid_sum += owner_paid
                grand_unique += len(rows)
                # Owner header (HTML block to avoid mixed markdown/HTML rendering bug)
                received_chip = f"<span style='color:#166534;font-size:13px;font-weight:500;'>· ${owner_paid:,.2f} received</span>"
                not_received_chip = (
                    f" <span style='color:#991b1b;font-size:13px;font-weight:500;'>· ${owner_pending:,.2f} not received</span>"
                    if owner_pending > 0 else ""
                )
                st.markdown(
                    f"<div style='margin-top:14px;margin-bottom:6px;font-size:18px;font-weight:700;color:#0f172a;'>"
                    f"{html_module.escape(owner)} — {len(rows)} deal(s) "
                    f"{received_chip}{not_received_chip}"
                    "</div>",
                    unsafe_allow_html=True,
                )
                rows_h: list[str] = []
                for r in rows:
                    dn = html_module.escape(r["deal_name"])
                    cd = html_module.escape(r["close_date"])
                    status = r["status"]
                    amount = r["amount_usd"]
                    paid_amt = r["paid_amount_usd"]
                    is_partial = (status == "Partially Paid")
                    is_unpaid = (paid_amt == 0 and not is_partial)
                    if is_partial:
                        bg, border, badge_color = "#fef9c3", "#fde68a", "#854d0e"
                    elif is_unpaid:
                        bg, border, badge_color = "#fef2f2", "#fecaca", "#b91c1c"
                    else:
                        bg, border, badge_color = "#f0fdf4", "#bbf7d0", "#166534"
                    rows_h.append(
                        f"<tr style='background:{bg};'>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid {border};'><strong>{dn}</strong></td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;'>{cd}</td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;color:{badge_color};'>{html_module.escape(status)}</td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;'><strong>${amount:,.2f}</strong></td>"
                        f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;color:{badge_color};'><strong>${paid_amt:,.2f}</strong></td>"
                        f"</tr>"
                    )
                # Total row for this owner
                rows_h.append(
                    "<tr style='background:#fef9c3;font-weight:700;'>"
                    f"<td style='padding:10px 12px;' colspan='2'>Total — {html_module.escape(owner)} ({len(rows)} deal{'s' if len(rows) != 1 else ''})</td>"
                    f"<td style='padding:10px 12px;text-align:right;'></td>"
                    f"<td style='padding:10px 12px;text-align:right;'>${owner_amount:,.2f}</td>"
                    f"<td style='padding:10px 12px;text-align:right;color:#166534;'>${owner_paid:,.2f}</td>"
                    "</tr>"
                )
                st.markdown(
                    "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:14px;'>"
                    "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
                    "<thead><tr style='background:#f1f5f9;color:#334155;'>"
                    "<th style='text-align:left;padding:8px 12px;font-weight:600;'>Deal</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Close date</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Status</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Amount</th>"
                    "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Payment Receipt Amount</th>"
                    f"</tr></thead><tbody>{''.join(rows_h)}</tbody></table></div>",
                    unsafe_allow_html=True,
                )

            # Grand total row across all owners
            grand_pending = grand_amount_sum - grand_paid_sum
            st.markdown(
                "<div style='margin-top:6px;padding:14px 16px;background:#fef3c7;border:2px solid #f59e0b;border-radius:10px;'>"
                "<div style='font-size:13px;font-weight:700;color:#78350f;text-transform:uppercase;letter-spacing:0.5px;'>🏷️ Grand total — AM team (all owners)</div>"
                "<div style='display:grid;grid-template-columns:repeat(4, 1fr);gap:12px;margin-top:8px;'>"
                f"<div><div style='font-size:11px;color:#92400e;'>Total deals</div><div style='font-size:20px;font-weight:700;color:#78350f;'>{grand_unique}</div></div>"
                f"<div><div style='font-size:11px;color:#92400e;'>Total Amount</div><div style='font-size:20px;font-weight:700;color:#78350f;'>${grand_amount_sum:,.2f}</div></div>"
                f"<div><div style='font-size:11px;color:#92400e;'>Payment Receipt</div><div style='font-size:20px;font-weight:700;color:#166534;'>${grand_paid_sum:,.2f}</div></div>"
                f"<div><div style='font-size:11px;color:#92400e;'>Pending</div><div style='font-size:20px;font-weight:700;color:#991b1b;'>${grand_pending:,.2f}</div></div>"
                "</div>"
                "</div>",
                unsafe_allow_html=True,
            )

    # ---- Commission summary table (same column structure as SMB Q1) ----
    st.markdown("#### Commission summary")
    st.caption("Commission is calculated on **Payment Received**. Achievement % uses booked vs quota for tier selection.")
    comm_reps = comm.get("reps") or []
    columns = [
        "Rep. Name",
        "Group",
        "Quota",
        "Revenue Achieved",
        "Achieved %",
        "Slab %",
        "Base Compensation",
        f"{int(ded_pct)}% Deduction",
        "Manage Deal",
        "Manage Pay",
        "Total Payout",
        "Exchange Rate",
        "Total Payout (INR)",
        "Note",
    ]
    widths = {
        "Rep. Name": "160px",
        "Group": "70px",
        "Quota": "110px",
        "Revenue Achieved": "140px",
        "Achieved %": "100px",
        "Slab %": "100px",
        "Base Compensation": "150px",
        f"{int(ded_pct)}% Deduction": "130px",
        "Manage Deal": "110px",
        "Manage Pay": "110px",
        "Total Payout": "140px",
        "Exchange Rate": "120px",
        "Total Payout (INR)": "160px",
        "Note": "520px",
    }
    money_cols = {"Quota", "Revenue Achieved", "Base Compensation", f"{int(ded_pct)}% Deduction", "Manage Deal", "Manage Pay", "Total Payout"}
    head = "".join(
        f'<th style="text-align:left;padding:8px 12px;border-bottom:1px solid #e5e7eb;background:#e0f2f1;font-weight:500;color:#00695c;min-width:{widths[c]};white-space:nowrap;">{html_module.escape(c)}</th>'
        for c in columns
    )

    base_total = 0.0
    ded_total = 0.0
    md_total = 0.0
    payout_total = 0.0
    rows_html: list[str] = []

    for c_row in comm_reps:
        nm = (c_row.get("name") or "").strip() or "—"
        grp = (c_row.get("group") or "").strip()
        quota = float(c_row.get("quota_usd") or 0)
        rev = float(c_row.get("revenue_achieved_usd") or 0)
        try:
            achv_pct_v = float(c_row.get("eligible_pct") or 0)
        except (TypeError, ValueError):
            achv_pct_v = 0.0
        base = float(c_row.get("base_commission_usd") or 0)
        try:
            ded_pct_row = float(c_row.get("deduction_pct"))
        except (TypeError, ValueError):
            ded_pct_row = ded_pct
        explicit_ded = c_row.get("deduction_usd")
        if explicit_ded is not None:
            try:
                deduction = float(explicit_ded)
            except (TypeError, ValueError):
                deduction = round(base * (ded_pct_row / 100.0), 2)
        else:
            deduction = round(base * (ded_pct_row / 100.0), 2)
        md_v = float(c_row.get("manage_deal_usd") or 0)
        md_pay_v = float(c_row.get("manage_deal_paid_now_usd") or 0)
        payout = float(c_row.get("total_payout_usd") or 0)
        payout_inr = round(payout * inr_rate, 2) if inr_rate else 0.0
        slab_pct = (base / float(c_row.get("payment_received_usd") or 1) * 100.0) if base and c_row.get("payment_received_usd") else 0.0
        if achv_pct_v < float(comm.get("eligibility_min_pct") or 50):
            slab_display = "Not eligible"
        elif slab_pct <= 0:
            slab_display = "—"
        else:
            slab_display = (f"{slab_pct:.0f}%" if abs(slab_pct - round(slab_pct)) < 1e-6 else f"{slab_pct:.1f}%")

        base_total += base
        ded_total += deduction
        md_total += md_v
        payout_total += payout

        cells = {
            "Rep. Name": nm,
            "Group": grp,
            "Quota": quota,
            "Revenue Achieved": rev,
            "Achieved %": f"{achv_pct_v:.0f}%",
            "Slab %": slab_display,
            "Base Compensation": base,
            f"{int(ded_pct)}% Deduction": deduction,
            "Manage Deal": md_v,
            "Manage Pay": md_pay_v,
            "Total Payout": payout,
            "Exchange Rate": f"₹{inr_rate:g}/USD" if inr_rate else "—",
            "Total Payout (INR)": payout_inr,
            "Note": (c_row.get("note") or "").strip(),
        }
        row_html: list[str] = []
        for c in columns:
            v = cells.get(c, "")
            if c in money_cols and isinstance(v, (int, float)):
                row_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">${float(v):,.2f}</td>')
            elif c == "Total Payout (INR)" and isinstance(v, (int, float)):
                row_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">₹{float(v):,.2f}</td>')
            else:
                row_html.append(f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{widths[c]};">{html_module.escape(str(v if v is not None else ""))}</td>')
        rows_html.append("<tr>" + "".join(row_html) + "</tr>")

    table_html = (
        '<div style="width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;">'
        '<table style="border-collapse:collapse;font-size:13px;width:max-content;min-width:100%;">'
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)

    # Footer totals
    payout_total_inr = round(payout_total * inr_rate, 2) if inr_rate else 0.0
    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Base compensation (sum)", f"${base_total:,.2f}")
    f2.metric(f"{int(ded_pct)}% deduction (sum)", f"−${ded_total:,.2f}")
    f3.metric("Total payout (USD)", f"${payout_total:,.2f}")
    if inr_rate:
        f4.metric(f"Total payout (INR @ ₹{inr_rate:g})", f"₹{payout_total_inr:,.2f}")

    # Inline chat assistant for AM Q1 (Claude-powered when API key set)
    render_pinned_chat_assistant(pinned, session_prefix="am_q1")


def render_pinned_monthly_manager_view(pinned: dict, team_label: str) -> None:
    """Manager incentive view for a single calendar month (e.g. April 2026).

    Shows the eligibility check (≥60% team target), and if eligible the commission slab + USD/INR payout.
    Designed for the Manager incentive tab when the Sales Target close-date is set to a single month —
    so no quarterly numbers are shown in that case.
    """
    mgr = pinned.get("manager_incentive") or {}
    tiers = pinned.get("monthly_tiers") or {}
    mgr_tiers = tiers.get("manager_tiers") or []
    mgr_min = float(tiers.get("manager_team_minimum_pct") or mgr.get("manager_team_minimum_pct") or 60)
    mgr_name = (mgr.get("manager_name") or "Manager").strip()
    year = pinned.get("year")
    month = pinned.get("month")
    label = pinned.get("month_label") or f"{year}-{month:02d}"
    # Month name for section labels (e.g. "May" or "April"), always derived from the pinned
    # month integer so the manager incentive section titles/labels reflect the currently
    # selected period — never a stale "April" when May is active.
    month_name = (_MONTH_ABBR_LONG.get(int(month or 0), "").title() or label.split(" ")[0]).strip()

    reps = pinned.get("reps") or []
    team_target = float(pinned.get("team_target_usd") or 0)
    team_paid_raw = sum(float(r.get("payment_received_usd") or 0) for r in reps)
    team_booked_raw = sum(float(r.get("achievement_usd") or 0) for r in reps)

    # Apply per-rep exceptions (e.g. Joy's Washington Post share) to both totals.
    total_exception_usd = 0.0
    exception_rows: list[dict] = []
    for _r_exc in reps:
        _exc = _r_exc.get("exception")
        if _exc:
            try:
                _share = float(_exc.get("joy_share_usd") or _exc.get("share_usd") or 0)
            except (TypeError, ValueError):
                _share = 0.0
            if _share:
                total_exception_usd += _share
                exception_rows.append(
                    {
                        "rep": (_r_exc.get("name") or "").strip(),
                        "deal_name": (_exc.get("deal_name") or "").strip(),
                        "deal_amount_usd": float(_exc.get("deal_amount_usd") or 0),
                        "share_usd": _share,
                        "note": (_exc.get("note") or "").strip(),
                    }
                )

    team_paid = team_paid_raw + total_exception_usd
    team_booked = team_booked_raw + total_exception_usd
    team_pct = (team_booked / team_target * 100.0) if team_target else 0.0
    try:
        inr_rate = float((pinned.get("commission") or {}).get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0

    st.markdown(f"### {label} · {mgr_name} — {team_label} Manager Incentive")
    st.caption(f"Pinned monthly numbers from `policy/{('am' if team_label.lower().startswith('am') or 'account' in team_label.lower() else 'smb')}_{_MONTH_ABBR_LONG.get(int(month or 0), 'april')}_{year}_fixed.json`.")

    # ---- Top KPI cards ----
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Monthly target", f"${team_target:,.0f}")
    k2.metric("Achieved (booked)", f"${team_booked:,.0f}", f"{team_pct:.1f}%")
    k3.metric("Payment received", f"${team_paid:,.0f}")
    k4.metric("Eligibility floor", f"{int(mgr_min)}%")

    st.markdown("---")

    # ---- Eligibility check (for current-month base only — pending commissions are NOT subject to this floor) ----
    is_eligible = team_pct >= mgr_min
    if is_eligible:
        mgr_slab_pct, mgr_band = _monthly_slab_pct(mgr_tiers, team_pct)
        mgr_base_amt = round(team_paid * mgr_slab_pct / 100.0, 2)
        st.markdown(
            f"<div style='background:#dcfce7;border:1px solid #4ade80;border-radius:8px;padding:14px 18px;'>"
            f"<div style='font-size:18px;font-weight:600;color:#14532d;'>✅ {mgr_name} is eligible for {month_name} base commission</div>"
            f"<div style='margin-top:6px;color:#14532d;'>Team achievement <strong>{team_pct:.1f}%</strong> ≥ floor <strong>{int(mgr_min)}%</strong>. Slab band: <strong>{mgr_band}</strong>.</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    else:
        mgr_slab_pct, mgr_band = 0.0, "—"
        mgr_base_amt = 0.0
        st.markdown(
            f"<div style='background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:14px 18px;'>"
            f"<div style='font-size:18px;font-weight:600;color:#7f1d1d;'>❌ {mgr_name} is NOT eligible for {month_name} base commission</div>"
            f"<div style='margin-top:6px;color:#7f1d1d;'>Team achievement is <strong>{team_pct:.1f}%</strong> of the ${team_target:,.0f} target, which is below the <strong>{int(mgr_min)}%</strong> floor. "
            f"Pending commissions from prior periods (below) are still paid.</div>"
            f"</div>",
            unsafe_allow_html=True,
        )

    # ---- Base calculation table ----
    st.markdown(f"#### {mgr_name} — {month_name} base calculation")
    st.markdown(
        f"<table style='border-collapse:collapse;font-size:13px;width:100%;border:1px solid #e5e7eb;border-radius:8px;'>"
        f"<thead><tr style='background:#ccfbf1;'>"
        f"<th style='text-align:left;padding:8px 12px;font-weight:500;color:#134e4a;'>Metric</th>"
        f"<th style='text-align:right;padding:8px 12px;font-weight:500;color:#134e4a;'>Value</th>"
        f"</tr></thead><tbody>"
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>Team achievement (booked)</td><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${team_booked:,.0f}</td></tr>"
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>Team achievement %</td><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{team_pct:.1f}%</td></tr>"
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>Team payment received (basis)</td><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${team_paid:,.0f}</td></tr>"
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>Slab band</td><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{mgr_band}</td></tr>"
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>Commission rate</td><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{mgr_slab_pct:.0f}%</td></tr>"
        f"<tr style='background:#bbf7d0;font-weight:600;'><td style='padding:10px 12px;'>{month_name} base payout (USD)</td><td style='padding:10px 12px;text-align:right;color:#14532d;'>${mgr_base_amt:,.2f}</td></tr>"
        + "</tbody></table>",
        unsafe_allow_html=True,
    )
    if is_eligible:
        st.caption(f"Calculation: ${team_paid:,.0f} × {mgr_slab_pct:.0f}% = ${mgr_base_amt:,.2f}")
    else:
        st.caption(f"Base = $0 (team below {int(mgr_min)}% floor).")

    # ---- Pending commissions (carry-over from prior periods, always paid) ----
    pending = mgr.get("pending_commissions") or []
    try:
        pending_total = float(mgr.get("total_pending_commission_usd") or sum(float(p.get("commission_usd") or 0) for p in pending))
    except (TypeError, ValueError):
        pending_total = sum(float(p.get("commission_usd") or 0) for p in pending)

    if pending:
        st.markdown("#### Pending commissions (prior periods)")
        pending_note = (mgr.get("pending_commission_note") or "").strip()
        if pending_note:
            st.caption(pending_note)
        pend_rows_html: list[str] = []
        for p in pending:
            dn = html_module.escape((p.get("deal_name") or "").strip())
            per = html_module.escape((p.get("period") or "").strip())
            amt = float(p.get("deal_amount_usd") or 0)
            rate = float(p.get("rate_pct") or 0)
            comm = float(p.get("commission_usd") or 0)
            note = html_module.escape((p.get("note") or "").strip())
            pend_rows_html.append(
                f"<tr style='background:#ecfdf5;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;'><strong>{dn}</strong><br/>"
                f"<span style='font-size:12px;color:#14532d;'>{note}</span></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;'>{per}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;'>${amt:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;'>{rate:g}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;color:#14532d;'><strong>+${comm:,.2f}</strong></td>"
                f"</tr>"
            )
        pend_rows_html.append(
            f"<tr style='background:#bbf7d0;font-weight:600;'>"
            f"<td colspan='4' style='padding:10px 12px;'>Total pending commission</td>"
            f"<td style='padding:10px 12px;text-align:right;font-variant-numeric:tabular-nums;color:#14532d;'>+${pending_total:,.2f}</td>"
            f"</tr>"
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #86efac;border-radius:8px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;min-width:680px;'>"
            "<thead><tr style='background:#a7f3d0;color:#064e3b;'>"
            "<th style='text-align:left;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Deal</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Period</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Deal amount</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Slab rate</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Commission</th>"
            f"</tr></thead><tbody>{''.join(pend_rows_html)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Clawbacks (carry-forward / prior-period deductions) ----
    clawbacks = mgr.get("clawbacks") or []
    try:
        total_cb = float(mgr.get("total_clawback_usd") or sum(float(c.get("deduction_usd") or 0) for c in clawbacks))
    except (TypeError, ValueError):
        total_cb = sum(float(c.get("deduction_usd") or 0) for c in clawbacks)

    if clawbacks:
        st.markdown("#### Clawbacks / carry-forward deductions")
        st.caption("Negative adjustments — typically carry-forward balances or prior-period slab corrections.")
        cb_rows_html: list[str] = []
        for cb in clawbacks:
            lbl = html_module.escape((cb.get("label") or "").strip())
            per = html_module.escape((cb.get("period") or "").strip())
            cb_note = html_module.escape((cb.get("note") or "").strip())
            try:
                cb_amt = float(cb.get("deal_amount_usd") or 0)
            except (TypeError, ValueError):
                cb_amt = 0.0
            try:
                cb_rate = float(cb.get("rate_pct") or 0)
            except (TypeError, ValueError):
                cb_rate = 0.0
            try:
                cb_ded = float(cb.get("deduction_usd") or 0)
            except (TypeError, ValueError):
                cb_ded = 0.0
            cb_rows_html.append(
                f"<tr style='background:#fff1f2;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;'><strong>{lbl}</strong>"
                f"<br/><span style='font-size:12px;color:#7f1d1d;'>{cb_note}</span></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;'>{per}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;font-variant-numeric:tabular-nums;'>" + (f"${cb_amt:,.2f}" if cb_amt else "—") + "</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;font-variant-numeric:tabular-nums;'>" + (f"{cb_rate:g}%" if cb_rate else "—") + "</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;font-variant-numeric:tabular-nums;color:#b91c1c;'><strong>−${cb_ded:,.2f}</strong></td>"
                f"</tr>"
            )
        cb_rows_html.append(
            f"<tr style='background:#fee2e2;font-weight:600;'>"
            f"<td colspan='4' style='padding:10px 12px;'>Total clawback</td>"
            f"<td style='padding:10px 12px;text-align:right;font-variant-numeric:tabular-nums;color:#991b1b;'>−${total_cb:,.2f}</td>"
            f"</tr>"
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:2px solid #fecaca;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;min-width:680px;'>"
            "<thead><tr style='background:#fecaca;color:#7f1d1d;'>"
            "<th style='text-align:left;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;'>Reason</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;'>Period</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;'>Received Amount</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;'>Slab rate</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;'>Deduction</th>"
            f"</tr></thead><tbody>{''.join(cb_rows_html)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Exceptions (e.g. Joy's Washington Post server split) — added to team totals ----
    if exception_rows:
        st.markdown("#### Exceptions (added to team total)")
        st.caption("Special allocations beyond captured HubSpot deals (e.g. server splits, SOW carve-outs). Each share is added to the team total used for the manager commission.")
        exc_rows_html: list[str] = []
        for ex in exception_rows:
            rep_nm = html_module.escape(ex.get("rep") or "—")
            deal_nm = html_module.escape(ex.get("deal_name") or "—")
            note = html_module.escape(ex.get("note") or "")
            deal_amt = float(ex.get("deal_amount_usd") or 0)
            share = float(ex.get("share_usd") or 0)
            exc_rows_html.append(
                f"<tr style='background:#eff6ff;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #bfdbfe;'><strong>{deal_nm}</strong>"
                f"<br/><span style='font-size:12px;color:#1e3a8a;'>{note}</span></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #bfdbfe;text-align:right;'>{rep_nm}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #bfdbfe;text-align:right;font-variant-numeric:tabular-nums;'>${deal_amt:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #bfdbfe;text-align:right;font-variant-numeric:tabular-nums;color:#1d4ed8;'><strong>+${share:,.2f}</strong></td>"
                f"</tr>"
            )
        exc_rows_html.append(
            f"<tr style='background:#dbeafe;font-weight:600;'>"
            f"<td colspan='3' style='padding:10px 12px;'>Total exceptions added</td>"
            f"<td style='padding:10px 12px;text-align:right;font-variant-numeric:tabular-nums;color:#1e40af;'>+${total_exception_usd:,.2f}</td>"
            f"</tr>"
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:2px solid #bfdbfe;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;min-width:680px;'>"
            "<thead><tr style='background:#bfdbfe;color:#1e3a8a;'>"
            "<th style='text-align:left;padding:10px 12px;border-bottom:2px solid #93c5fd;font-weight:600;'>Deal</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #93c5fd;font-weight:600;'>Attributed to</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #93c5fd;font-weight:600;'>Deal amount</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #93c5fd;font-weight:600;'>Share added</th>"
            f"</tr></thead><tbody>{''.join(exc_rows_html)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Optional plain-text exception note from JSON (legacy field) ----
    exc_note = (mgr.get("exception_note") or "").strip()
    if exc_note and not exception_rows:
        st.info(f"**Exception applied:** {exc_note}")

    # ---- Final payout ----
    explicit_final = mgr.get("final_payout_usd")
    if explicit_final is not None:
        final_amt = float(explicit_final)
    else:
        final_amt = round(mgr_base_amt + pending_total - total_cb, 2)
    final_inr = round(final_amt * inr_rate, 2) if inr_rate else 0.0

    st.markdown(f"#### {mgr_name} — {label} total payout")
    inr_suffix = (
        f"<div style='font-size:14px;font-weight:500;color:#166534;margin-top:4px;'>≈ ₹{final_inr:,.0f} (INR @ ₹{inr_rate:g}/USD)</div>"
        if inr_rate
        else ""
    )
    if final_amt < 0:
        st.markdown(
            f"<div style='background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:14px 18px;color:#7f1d1d;'>"
            f"<div style='font-size:18px;font-weight:600;'>Total {label} Payout: $0.00</div>"
            f"<div style='font-size:14px;font-weight:500;color:#991b1b;margin-top:4px;'>Net balance: <strong>${final_amt:,.2f}</strong> — carried forward to next compensation cycle.</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"<div style='background:#bbf7d0;border:1px solid #4ade80;border-radius:8px;padding:14px 18px;color:#14532d;'>"
            f"<div style='font-size:18px;font-weight:600;'>Total {label} Payout: ${final_amt:,.2f}</div>"
            f"{inr_suffix}"
            f"</div>",
            unsafe_allow_html=True,
        )

    # ---- Step-by-step calculation summary ----
    explicit_calc = (mgr.get("calculation_note") or "").strip()
    if explicit_calc:
        st.caption("Calculation: " + explicit_calc.replace("$", "\\$"))
    else:
        parts_calc = [f"Base \\${mgr_base_amt:,.2f}"]
        if pending_total:
            parts_calc.append(f"+ Pending \\${pending_total:,.2f}")
        if total_cb:
            parts_calc.append(f"− Clawback \\${total_cb:,.2f}")
        parts_calc.append(f"= \\${final_amt:,.2f}")
        st.caption("Calculation: " + " ".join(parts_calc))


def _gather_compensation_chat_context() -> dict:
    """Gather a compact dict from every data source so the AI assistant can answer broadly.

    Includes pinned SMB Q1 2026 data, live rep_incentives aggregates per team, user roster,
    outbound meetings, policy summary (commission tiers / outbound payouts / fixed targets),
    and any HubSpot session cache the user has fetched in this session.
    """
    from commission_policy import (
        ACCOUNT_MANAGEMENT_TEAM_NAME,
        AM_QUOTA_ACHIEVEMENT_ENABLED,
        ENTERPRISE_TEAM_NAME,
        OUTBOUND_ELIGIBLE_REGIONS,
        OUTBOUND_MEETING_PAYOUT_NOTE,
        OUTBOUND_MEETING_PAYOUT_ROWS,
        OUTBOUND_MEETING_PAYOUT_TITLE,
        OUTBOUND_POLICY_LABEL,
        REP_SLAB_ROWS_FOR_DB,
        SMB_ACHIEVEMENT_TIERS,
        SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
        SMB_QUOTA_ACHIEVEMENT_ENABLED,
        SMB_TEAM_NAME,
        TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT,
        TEAM_QUARTERLY_TARGETS_USD,
    )

    # Pinned SMB
    pinned = _load_pinned_smb_quarter(2026, 1) or {}

    # Database aggregates
    incentives = []
    try:
        incentives = get_rep_incentives() or []
    except Exception:
        incentives = []
    by_team: dict[str, list[dict]] = {}
    for r in incentives:
        tn = (r.get("team_name") or "").strip() or "Unknown"
        by_team.setdefault(tn, []).append(r)

    def _team_summary(rows: list[dict]) -> dict:
        if not rows:
            return {"rep_count": 0, "total_revenue_usd": 0, "total_paid_usd": 0, "total_incentive_usd": 0}
        return {
            "rep_count": len({r.get("user_id") for r in rows if r.get("user_id") is not None}),
            "total_revenue_usd": round(sum(float(r.get("total_revenue") or 0) for r in rows), 2),
            "total_paid_usd": round(sum(float(r.get("total_paid_amount") or 0) for r in rows), 2),
            "total_incentive_usd": round(sum(float(r.get("incentive_amount") or 0) for r in rows), 2),
            "periods": sorted({(r.get("calculation_period") or "").strip() for r in rows if r.get("calculation_period")}),
        }

    teams = {tn: _team_summary(rows) for tn, rows in by_team.items()}

    # User roster (lean)
    users = []
    try:
        for u in (get_all_users_with_teams(active_only=False) or []):
            users.append(
                {
                    "user_id": u.get("user_id"),
                    "full_name": u.get("full_name"),
                    "email": u.get("email"),
                    "team": u.get("team_name"),
                    "role": u.get("role"),
                    "compensation_group": u.get("compensation_group"),
                    "hubspot_quota_usd": u.get("hubspot_quota_usd"),
                    "active": u.get("is_active", True),
                }
            )
    except Exception:
        users = []

    # Outbound meetings
    outbound = []
    try:
        for r in (get_all_outbound_meetings() or [])[:200]:
            outbound.append(
                {
                    "rep_name": r.get("rep_name"),
                    "rep_email": r.get("rep_email"),
                    "region": r.get("region"),
                    "meeting_date": str(r.get("meeting_date") or ""),
                    "incentive_usd": float(r.get("incentive_amount") or 0),
                }
            )
    except Exception:
        outbound = []

    # Policy summary
    policy = {
        "team_quarterly_targets_usd": dict(TEAM_QUARTERLY_TARGETS_USD or {}),
        "smb": {
            "team_name": SMB_TEAM_NAME,
            "min_achievement_pct_for_commission": SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
            "quota_achievement_enabled": bool(SMB_QUOTA_ACHIEVEMENT_ENABLED),
            "achievement_tiers": SMB_ACHIEVEMENT_TIERS,
            "rep_slab_rows": [list(r) for r in REP_SLAB_ROWS_FOR_DB],
        },
        "am": {
            "team_name": ACCOUNT_MANAGEMENT_TEAM_NAME,
            "quota_achievement_enabled": bool(AM_QUOTA_ACHIEVEMENT_ENABLED),
        },
        "ent": {"team_name": ENTERPRISE_TEAM_NAME},
        "manager_pool_tiers_pct": [list(t) for t in TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT],
        "outbound": {
            "policy_label": OUTBOUND_POLICY_LABEL,
            "eligible_regions": [list(t) for t in OUTBOUND_ELIGIBLE_REGIONS],
            "payout_title": OUTBOUND_MEETING_PAYOUT_TITLE,
            "payout_rows": OUTBOUND_MEETING_PAYOUT_ROWS,
            "payout_note": OUTBOUND_MEETING_PAYOUT_NOTE,
        },
    }

    # Optional HubSpot session cache (only what's already been fetched in this session).
    hubspot_session = {}
    for ctx in ("smb", "am", "ent"):
        sd = st.session_state.get(f"hubspot_last_fetch_summary_{ctx}")
        if sd:
            hubspot_session[ctx] = {"summary": sd}

    return {
        "pinned_smb_q1_2026": pinned,
        "team_aggregates_from_db": teams,
        "user_roster": users,
        "outbound_meetings": outbound,
        "policy": policy,
        "hubspot_session_cache": hubspot_session,
    }


def _detect_pinned_period_for_question(question: str) -> dict | None:
    """When no API key, route a question to the most relevant pinned file based on keywords.

    Returns the pinned dict if a match is found, else None. Detects the team (SMB vs AM) from
    rep / manager names and the period (April / quarterly) from month words or date patterns.
    """
    q = (question or "").lower()
    if not q.strip():
        return None

    # Period detection
    april_indicators = ("april", "apr ", "apr-", "apr/", "2026/04", "2026-04", "04/2026", "04-2026", "04/01", "04/30", "4/2026")
    quarterly_indicators = ("q1", "q2", "q3", "q4", "quarter", "fy2026", "fy26")
    is_april = any(tok in q for tok in april_indicators)
    is_quarterly = any(tok in q for tok in quarterly_indicators) and not is_april

    # Team detection by rep / manager first names
    am_indicators = ("joy", "vivin", "arundhati", "arundhathi", "account management", "am team", " am ", " am.", " am,")
    smb_indicators = ("chitradip", "chit ", "lawrence", "larry", "yogi", "yogesh", "vicky", "kritika", "deepak", "kartik", "royston", "rutuja", "smb")
    is_am = any(tok in q for tok in am_indicators)
    is_smb = any(tok in q for tok in smb_indicators)

    # Resolve which pinned file
    if is_april and is_am:
        return _load_pinned_monthly("am", 2026, 4)
    if is_april and is_smb:
        return _load_pinned_monthly("smb", 2026, 4)
    if is_april:
        # No specific team — default AM April since that's where Joy/most exception data lives.
        return _load_pinned_monthly("am", 2026, 4) or _load_pinned_monthly("smb", 2026, 4)
    if is_am:
        return _load_pinned_am_quarter(2026, 1)
    if is_smb or is_quarterly:
        return _load_pinned_smb_quarter(2026, 1)
    return None


def _answer_pinned_period(pinned: dict, question: str) -> str:
    """Lightweight rule-based answer over any pinned period (monthly or quarterly).

    Uses the pre-computed summary so totals, eligibility, and rep status are correct
    without doing any math by hand. Covers the most common questions; otherwise asks
    the user to set ANTHROPIC_API_KEY for full AI answers.
    """
    if not pinned:
        return (
            "I couldn't find pinned data for that period. Set `ANTHROPIC_API_KEY` in `.env` "
            "for full AI-powered answers across all periods, or rephrase your question."
        )

    q = (question or "").lower().strip()
    summary = _compute_pinned_summary(pinned)
    period = summary.get("period") or {}
    yr = period.get("year")
    qt = period.get("quarter")
    mo = period.get("month")
    if qt:
        period_lbl = f"Q{qt} FY{yr}"
    elif mo:
        period_lbl = period.get("month_label") or f"{yr}-{mo:02d}"
    else:
        period_lbl = "this period"

    inr_rate = float(summary.get("inr_rate_per_usd") or 0)

    # Escape "\$" so Streamlit Markdown doesn't treat it as LaTeX math-mode.
    def _usd(x):
        return f"\\${float(x):,.2f}"

    def _inr(x):
        return f"₹{float(x) * inr_rate:,.0f}" if inr_rate else ""

    def _usd_with_inr(x):
        """USD with optional INR appendix — used for inline mentions like "(≈ ₹XYZ)"."""
        u = _usd(x)
        if inr_rate:
            return f"{u} (≈ {_inr(x)})"
        return u

    def _esc_md(s):
        """Escape $ in JSON-derived text so it doesn't trigger Markdown math-mode."""
        if not s:
            return ""
        return str(s).replace("\\$", "$").replace("$", "\\$")

    # Look up specific rep / manager
    rep_status_list = summary.get("rep_status") or []
    rep_payouts_list = summary.get("rep_payouts") or []
    mgr_summary = summary.get("manager_incentive_summary") or {}
    mgr_name_low = (mgr_summary.get("manager_name") or "").lower()
    # Define team_summary early so the rep-narrative branch below can read the eligibility floor.
    team_summary = summary.get("team") or {}

    def _cross_period_note_for_rep(first_name: str) -> str:
        """Look across other pinned files to see if this rep has notable status (e.g. left_org)
        in another period — useful when answering questions on Q1 about someone who later left.
        """
        if not first_name:
            return ""
        candidates = (
            _load_pinned_monthly("smb", 2026, 4),
            _load_pinned_monthly("am", 2026, 4),
            _load_pinned_smb_quarter(2026, 1),
            _load_pinned_am_quarter(2026, 1),
        )
        current_yr = (period or {}).get("year")
        current_qt = (period or {}).get("quarter")
        current_mo = (period or {}).get("month")
        for cand in candidates:
            if not cand:
                continue
            # Skip the current period itself
            if int(cand.get("year") or 0) == int(current_yr or 0) and (
                int(cand.get("quarter") or 0) == int(current_qt or 0)
                and int(cand.get("month") or 0) == int(current_mo or 0)
            ):
                continue
            for rr in (cand.get("reps") or []):
                rname = (rr.get("name") or "").strip().lower()
                if rname.startswith(first_name) and rr.get("left_org"):
                    mo_word = ""
                    try:
                        mo_word = (cand.get("month_label") or "").split()[0]
                    except Exception:
                        mo_word = ""
                    if mo_word:
                        return f"Additionally, {first_name.title()} left the organization in {mo_word}, so no further payouts were processed afterward."
                    return f"Additionally, {first_name.title()} later left the organization, so no further payouts were processed afterward."
        return ""

    def _format_manager_answer() -> str:
        return (
            f"**{mgr_summary.get('manager_name')} — {period_lbl} manager incentive**\n"
            f"- Eligible for base this period: {'✅ yes' if mgr_summary.get('is_eligible_for_base') else '❌ no'}\n"
            f"- Base commission: {_usd(mgr_summary.get('base_usd') or 0)}\n"
            f"- Pending commissions (carry-over from prior periods): +{_usd(mgr_summary.get('pending_total_usd') or 0)}\n"
            f"- Clawbacks: −{_usd(mgr_summary.get('clawback_total_usd') or 0)}\n"
            f"- **Final payout: {_usd_with_inr(mgr_summary.get('final_payout_usd') or 0)}**"
            + (f"\n\n_{_esc_md(mgr_summary.get('calculation_note'))}_" if mgr_summary.get("calculation_note") else "")
        )

    # ------------------------------------------------------------------
    # ENT-specific intent handlers (Q1 ENT page suggested prompts)
    # ------------------------------------------------------------------
    # These take priority over the generic per-rep / manager narrative because
    # the Enterprise team uses different language (slab bands, Payment Received
    # vs booked revenue, Not-Paid deal exclusion, Q4 pending + Cambrex clawback).
    team_label_low = ((pinned.get("manager_incentive") or {}).get("team_label") or "").lower()
    unpaid_overrides = pinned.get("unpaid_deals_overrides") or []
    paid_deals_block = pinned.get("paid_deals") or []
    mgr_pending_block = (pinned.get("manager_incentive") or {}).get("pending_commissions") or []
    mgr_clawback_block = (pinned.get("manager_incentive") or {}).get("clawbacks") or []
    # ENT-specific gate: must be tied to the Enterprise team (or Anthony as sole rep).
    # The AM page now also has paid_deals / unpaid_deals_overrides, so we can't use those
    # alone to decide — that would mis-route AM questions to the ENT handlers.
    is_ent_period = ("enterprise" in team_label_low) or (
        bool(summary.get("manager_is_also_a_rep")) and (
            mgr_name_low.startswith("anthony") or "anthony" in mgr_name_low
        )
    )

    def _ent_unpaid_deals_answer() -> str:
        """Answer questions about 'Not Paid' Q1 deals and how they're excluded from Payment Received."""
        if not unpaid_overrides:
            return f"No 'Not Paid' deal exclusions are recorded for {period_lbl}."
        booked_total = float(pinned.get("team_achievement_usd") or summary.get("team", {}).get("achievement_raw_usd") or 0)
        paid_total = float(pinned.get("team_payment_received_usd") or summary.get("team", {}).get("payment_received_effective_usd") or 0)
        unpaid_total = sum(float(u.get("amount_usd") or 0) for u in unpaid_overrides)
        paid_deals_total = sum(float(p.get("amount_usd") or 0) for p in paid_deals_block)
        _paid_count_q = len(paid_deals_block)
        _unpaid_count_q = len(unpaid_overrides)
        _closed_count_q = _paid_count_q + _unpaid_count_q
        out: list[str] = [
            f"### {period_lbl} — Q1 deals marked **Not Paid** (excluded from Payment Received)",
            "",
            f"**Deal counts:** {_closed_count_q} deals closed in total — "
            f"**{_paid_count_q}** with payment received ({_usd(paid_deals_total)}), "
            f"**{_unpaid_count_q}** with payment not yet received ({_usd(unpaid_total)}).",
            "",
            f"These {_unpaid_count_q} Q1 deals were marked **Not Paid** in HubSpot at the time the Q1 commission was calculated. "
            "Per ENT policy, commission is calculated on **Payment Received**, not booked revenue — so these deals are "
            "excluded from the commission base until payment arrives.",
            "",
            "| Deal | Close date | Amount |",
            "| --- | --- | --- |",
        ]
        for u in unpaid_overrides:
            dn = _esc_md((u.get("deal_name") or "").strip() or "(unnamed deal)")
            cd = _esc_md((u.get("close_date") or "").strip() or "—")
            amt = float(u.get("amount_usd") or 0)
            out.append(f"| **{dn}** | {cd} | {_usd(amt)} |")
        out.append("")
        out.append(f"**Total unpaid: {_usd(unpaid_total)}**")
        out.append("")
        # If we have the paid-deal detail, show it for symmetry.
        if paid_deals_block:
            out.append(f"### Paid Q1 deals (included in Payment Received)")
            out.append("")
            out.append("| Deal | Close date | Amount |")
            out.append("| --- | --- | --- |")
            for p in paid_deals_block:
                dn = _esc_md((p.get("deal_name") or "").strip() or "(unnamed deal)")
                cd = _esc_md((p.get("close_date") or "").strip() or "—")
                amt = float(p.get("amount_usd") or 0)
                out.append(f"| **{dn}** | {cd} | {_usd(amt)} |")
            out.append("")
            out.append(f"**Total paid: {_usd(paid_deals_total)}**")
            out.append("")
        out.append("**How the exclusion works**")
        out.append(f"- Booked revenue (HubSpot total): {_usd(booked_total)}")
        if paid_deals_block:
            out.append(f"- = Paid Q1 deals (Payment Received): {_usd(paid_deals_total)}")
        out.append(f"- Minus Not-Paid deals: −{_usd(unpaid_total)}")
        out.append(f"- **= Payment Received: {_usd(paid_total)}**")
        out.append("")
        # Find the ENT rep / slab rate from the commission block
        ent_reps = ((pinned.get("commission") or {}).get("reps") or [])
        if ent_reps:
            r0 = ent_reps[0]
            slab_pct = float(r0.get("slab_pct") or 0)
            base = float(r0.get("base_commission_usd") or 0)
            person = (r0.get("name") or "Anthony").split()[0]
            out.append(
                f"Anthony's base commission is therefore **{_usd(paid_total)} × {slab_pct:g}% = {_usd(base)}** — "
                f"not {_usd(booked_total)} × {slab_pct:g}%. Once payment arrives for any of the Not-Paid deals, the "
                f"corresponding commission will be released in a future cycle."
            )
        else:
            out.append(
                "Anthony's Q1 base commission is calculated on the Payment Received amount above at the ENT slab rate."
            )
        return "\n".join(out)

    def _ent_refund_clawback_answer() -> str:
        """Answer questions about the Cambrex refund / clawback."""
        if not mgr_clawback_block:
            return f"No refunds or clawbacks are recorded for {period_lbl}."
        out: list[str] = [
            f"### {period_lbl} — Refund / clawback breakdown",
            "",
        ]
        cb_total = 0.0
        for cb in mgr_clawback_block:
            lbl = _esc_md((cb.get("label") or cb.get("deal_name") or "").strip() or "(unnamed deal)")
            per = _esc_md((cb.get("period") or "").strip())
            refund_amt = float(cb.get("deal_amount_usd") or 0)
            rate = float(cb.get("rate_pct") or 0)
            ded = float(cb.get("deduction_usd") or 0)
            note = _esc_md((cb.get("note") or "").strip())
            cb_total += ded
            out.append(f"**{lbl}** ({per})")
            out.append(f"- Original deal amount refunded: {_usd(refund_amt)}")
            out.append(f"- Commission previously paid at the Q4 rate: {refund_amt:,.2f} × {rate:g}% = **{_usd(ded)}**")
            out.append(f"- This amount is **clawed back** (deducted from {period_lbl}'s payout).")
            if note:
                out.append(f"  _{note}_")
            out.append("")
        if len(mgr_clawback_block) > 1:
            out.append(f"**Total clawback: −{_usd(cb_total)}**")
            out.append("")
        # Show how it flows into Anthony's total
        ent_reps = ((pinned.get("commission") or {}).get("reps") or [])
        if ent_reps:
            r0 = ent_reps[0]
            base = float(r0.get("base_commission_usd") or 0)
            pending_total = float((pinned.get("manager_incentive") or {}).get("total_pending_commission_usd") or 0)
            final = float((pinned.get("manager_incentive") or {}).get("final_payout_usd") or float(r0.get("total_payout_usd") or 0))
            person = (r0.get("name") or "Anthony").split()[0]
            out.append("**How it flows into the final payout**")
            out.append(
                f"- {person}'s Q1 base: {_usd(base)} + Q4 pending: +{_usd(pending_total)} − clawback: −{_usd(cb_total)} "
                f"= **{_usd(final)}**"
            )
        return "\n".join(out)

    def _ent_anthony_total_answer() -> str:
        """Answer 'What is Anthony's total Q1 payout, including Q4 pending deals?' — full breakdown."""
        ent_reps = ((pinned.get("commission") or {}).get("reps") or [])
        if not ent_reps:
            return f"No commission entries are recorded for {period_lbl}."
        r0 = ent_reps[0]
        nm = (r0.get("name") or "Anthony Raymond").strip()
        quota = float(r0.get("quota_usd") or 0)
        rev = float(r0.get("revenue_achieved_usd") or 0)
        paid = float(r0.get("payment_received_usd") or 0)
        ach_pct = float(r0.get("eligible_pct") or 0)
        slab_pct = float(r0.get("slab_pct") or 0)
        base = float(r0.get("base_commission_usd") or 0)
        unpaid_total = sum(float(u.get("amount_usd") or 0) for u in unpaid_overrides)
        pending_total = float((pinned.get("manager_incentive") or {}).get("total_pending_commission_usd") or 0)
        clawback_total = float((pinned.get("manager_incentive") or {}).get("total_clawback_usd") or 0)
        final = float((pinned.get("manager_incentive") or {}).get("final_payout_usd") or float(r0.get("total_payout_usd") or 0))

        _paid_count_a = len(paid_deals_block)
        _unpaid_count_a = len(unpaid_overrides)
        _closed_count_a = _paid_count_a + _unpaid_count_a
        out: list[str] = [
            f"### {nm} — {period_lbl} total payout: **{_usd(final)}**",
            "",
            "**Performance**",
            f"- Quota: {_usd(quota)}",
            f"- Revenue achieved (booked): {_usd(rev)} ({ach_pct:.1f}% of quota)",
        ]
        if _closed_count_a > 0:
            out.append(
                f"- Deals closed: **{_closed_count_a}** total — "
                f"**{_paid_count_a}** paid, **{_unpaid_count_a}** not paid"
            )
        out.extend([
            f"- Payment received (after excluding {_unpaid_count_a} Not-Paid deals totaling {_usd(unpaid_total)}): **{_usd(paid)}**",
            "",
            "**Base commission**",
            f"- Achievement of {ach_pct:.1f}% lands in the **0–75.99% ENT slab band → {slab_pct:g}%** commission rate.",
            f"- Commission is calculated on **Payment Received**, not booked revenue:",
            f"  - {_usd(paid)} × {slab_pct:g}% = **{_usd(base)}**",
            "",
        ])
        if mgr_pending_block:
            out.append("**Q4 pending commissions (paid in Q1)**")
            for p in mgr_pending_block:
                dn = _esc_md((p.get("deal_name") or "").strip())
                per = _esc_md((p.get("period") or "").strip())
                deal_amt = float(p.get("deal_amount_usd") or 0)
                rate = float(p.get("rate_pct") or 0)
                comm = float(p.get("commission_usd") or 0)
                out.append(f"- **{dn}** ({per}) — {_usd(deal_amt)} × {rate:g}% = **+{_usd(comm)}**")
            out.append(f"  **Subtotal: +{_usd(pending_total)}**")
            out.append("")
        if mgr_clawback_block:
            out.append("**Refunds / clawbacks**")
            for cb in mgr_clawback_block:
                lbl = _esc_md((cb.get("label") or cb.get("deal_name") or "").strip())
                per = _esc_md((cb.get("period") or "").strip())
                refund_amt = float(cb.get("deal_amount_usd") or 0)
                rate = float(cb.get("rate_pct") or 0)
                ded = float(cb.get("deduction_usd") or 0)
                out.append(f"- **{lbl}** ({per}) — {_usd(refund_amt)} × {rate:g}% = **−{_usd(ded)}**")
            out.append(f"  **Subtotal: −{_usd(clawback_total)}**")
            out.append("")
        out.append("**Final math**")
        out.append(
            f"- Q1 base {_usd(base)} + Q4 pending +{_usd(pending_total)} − clawback −{_usd(clawback_total)} = **{_usd(final)}**"
        )
        return "\n".join(out)

    def _ent_step_by_step_answer() -> str:
        """Numbered, policy-aware walkthrough of the full ENT commission calculation.

        Used for prompts like "Explain the ENT commission calculation step by step",
        "Walk me through the calculation", "How was the commission calculated", etc.
        """
        ent_reps = ((pinned.get("commission") or {}).get("reps") or [])
        if not ent_reps:
            return f"No ENT commission entries are recorded for {period_lbl}."
        r0 = ent_reps[0]
        nm = (r0.get("name") or "Anthony Raymond").strip()
        quota = float(r0.get("quota_usd") or 0)
        booked = float(r0.get("revenue_achieved_usd") or 0)
        paid = float(r0.get("payment_received_usd") or 0)
        ach_pct = float(r0.get("eligible_pct") or 0)
        slab_pct = float(r0.get("slab_pct") or 0)
        base = float(r0.get("base_commission_usd") or 0)
        unpaid_total = sum(float(u.get("amount_usd") or 0) for u in unpaid_overrides)
        pending_total = float((pinned.get("manager_incentive") or {}).get("total_pending_commission_usd") or 0)
        clawback_total = float((pinned.get("manager_incentive") or {}).get("total_clawback_usd") or 0)
        final = float(
            (pinned.get("manager_incentive") or {}).get("final_payout_usd")
            or float(r0.get("total_payout_usd") or 0)
        )

        _paid_count_s = len(paid_deals_block)
        _unpaid_count_s = len(unpaid_overrides)
        _closed_count_s = _paid_count_s + _unpaid_count_s
        out: list[str] = [
            f"## {period_lbl} — ENT commission calculation, step by step",
            "",
            f"This is exactly how {nm}'s **{_usd(final)}** total payout is built from the policy and the deal data.",
            "",
            "---",
            "",
            "### Step 1 — Pull the team's Q1 performance from HubSpot",
            "",
            f"- **Quota** (Q1 ENT target): {_usd(quota)}",
            f"- **Revenue achieved (booked)**: {_usd(booked)}  →  **{ach_pct:.1f}% of quota**",
        ]
        if _closed_count_s > 0:
            out.append(
                f"- **Deal counts**: {_closed_count_s} deals closed total — "
                f"**{_paid_count_s}** with payment received, **{_unpaid_count_s}** with payment not received."
            )
        out.extend([
            "",
            "These are the raw deal totals closed in the quarter. They aren't the commission base yet — "
            "we still need to subtract any deals that weren't actually paid by the close of the commission cycle.",
            "",
            "---",
            "",
            "### Step 2 — Exclude Q1 deals that were Not Paid",
            "",
            "ENT policy: **commission is calculated on Payment Received, not booked revenue.** "
            "Any deal still marked Not Paid in HubSpot at the time the commission is run is excluded from the base.",
            "",
        ])
        if unpaid_overrides:
            out.append("**Not-Paid deals (excluded):**")
            out.append("")
            out.append("| Deal | Close date | Amount |")
            out.append("| --- | --- | --- |")
            for u in unpaid_overrides:
                dn = _esc_md((u.get("deal_name") or "").strip() or "(unnamed deal)")
                cd = _esc_md((u.get("close_date") or "").strip() or "—")
                amt = float(u.get("amount_usd") or 0)
                out.append(f"| **{dn}** | {cd} | {_usd(amt)} |")
            out.append("")
            out.append(f"**Total Not-Paid: {_usd(unpaid_total)}**")
            out.append("")
            if paid_deals_block:
                paid_deals_total = sum(float(p.get("amount_usd") or 0) for p in paid_deals_block)
                out.append("**Paid Q1 deals (included in Payment Received):**")
                out.append("")
                out.append("| Deal | Close date | Amount |")
                out.append("| --- | --- | --- |")
                for p in paid_deals_block:
                    dn = _esc_md((p.get("deal_name") or "").strip() or "(unnamed deal)")
                    cd = _esc_md((p.get("close_date") or "").strip() or "—")
                    amt = float(p.get("amount_usd") or 0)
                    out.append(f"| **{dn}** | {cd} | {_usd(amt)} |")
                out.append("")
                out.append(f"**Total Paid: {_usd(paid_deals_total)}**")
                out.append("")
            out.append("Math:")
            out.append("```")
            out.append(f"Booked          {booked:>14,.2f}")
            out.append(f"− Not Paid      {unpaid_total:>14,.2f}")
            out.append(f"────────────────────────────────")
            out.append(f"Payment Received{paid:>14,.2f}")
            out.append("```")
        else:
            out.append(f"No Not-Paid exclusions for {period_lbl}. Payment Received = Booked = {_usd(paid)}.")
        out.append("")
        out.append("---")
        out.append("")
        out.append("### Step 3 — Determine the slab band from achievement %")
        out.append("")
        out.append("ENT slab bands (achievement vs quota → commission rate):")
        out.append("")
        out.append("| Achievement | Commission rate |")
        out.append("| --- | --- |")
        out.append("| 0 – 75.99% | **7%** |")
        out.append("| 76 – 100% | 9% |")
        out.append("| 101 – 125% | 11% |")
        out.append("| 126%+ | 13% |")
        out.append("")
        out.append(
            f"At **{ach_pct:.1f}% of quota**, {nm.split()[0]} lands in the **0 – 75.99%** band → "
            f"**{slab_pct:g}%** commission rate."
        )
        out.append("")
        out.append(
            "Note: ENT has a **0% eligibility floor** — any positive revenue qualifies, "
            "so there's no \"below the floor → $0 commission\" outcome like SMB/AM."
        )
        out.append("")
        out.append("---")
        out.append("")
        out.append("### Step 4 — Calculate Q1 base commission")
        out.append("")
        out.append("```")
        out.append(f"Base = Payment Received × Slab %")
        out.append(f"     = {paid:,.2f} × {slab_pct:g}%")
        out.append(f"     = {base:,.2f}")
        out.append("```")
        out.append("")
        out.append(
            f"**Q1 base commission: {_usd(base)}**  "
            f"(Note: we use Payment Received {_usd(paid)}, NOT booked revenue {_usd(booked)}.)"
        )
        out.append("")
        out.append("---")
        out.append("")
        # Step 5 — Q4 pending
        if mgr_pending_block:
            out.append("### Step 5 — Add Q4 pending commissions (paid in Q1 at the Q4 rate)")
            out.append("")
            out.append(
                "These are deals that closed in Q4 FY2025 but whose payment was actually received during Q1. "
                "Per policy, the commission is paid at the **Q4 rate (10%)**, not the current Q1 rate."
            )
            out.append("")
            out.append("| Deal | Period | Paid amount | Q4 rate | Commission |")
            out.append("| --- | --- | --- | --- | --- |")
            for p in mgr_pending_block:
                dn = _esc_md((p.get("deal_name") or "").strip())
                per = _esc_md((p.get("period") or "").strip())
                deal_amt = float(p.get("deal_amount_usd") or 0)
                rate = float(p.get("rate_pct") or 0)
                comm = float(p.get("commission_usd") or 0)
                out.append(f"| **{dn}** | {per} | {_usd(deal_amt)} | {rate:g}% | **+{_usd(comm)}** |")
            out.append("")
            out.append(f"**Subtotal pending: +{_usd(pending_total)}**")
            out.append("")
            out.append("---")
            out.append("")
        # Step 6 — Clawbacks
        if mgr_clawback_block:
            out.append("### Step 6 — Deduct Q4 refunds / clawbacks")
            out.append("")
            out.append(
                "If a Q4 deal that previously paid out commission is later cancelled and refunded, "
                "that earlier commission is **clawed back** from the current quarter's payout — "
                "again at the **original Q4 rate (10%)**."
            )
            out.append("")
            out.append("| Refund / cancelled deal | Period | Refund amount | Q4 rate | Clawback |")
            out.append("| --- | --- | --- | --- | --- |")
            for cb in mgr_clawback_block:
                lbl = _esc_md((cb.get("label") or cb.get("deal_name") or "").strip())
                per = _esc_md((cb.get("period") or "").strip())
                refund_amt = float(cb.get("deal_amount_usd") or 0)
                rate = float(cb.get("rate_pct") or 0)
                ded = float(cb.get("deduction_usd") or 0)
                out.append(f"| **{lbl}** | {per} | {_usd(refund_amt)} | {rate:g}% | **−{_usd(ded)}** |")
            out.append("")
            out.append(f"**Subtotal clawback: −{_usd(clawback_total)}**")
            out.append("")
            out.append("---")
            out.append("")
        # Step 7 — Final
        step_n = 7 if (mgr_pending_block and mgr_clawback_block) else (6 if (mgr_pending_block or mgr_clawback_block) else 5)
        out.append(f"### Step {step_n} — Add it all up")
        out.append("")
        out.append("```")
        out.append(f"Q1 base                {base:>12,.2f}")
        if mgr_pending_block:
            out.append(f"+ Q4 pending           {pending_total:>12,.2f}")
        if mgr_clawback_block:
            out.append(f"− Q4 clawback          {clawback_total:>12,.2f}")
        out.append(f"───────────────────────────────────")
        out.append(f"Final {period_lbl} payout {final:>12,.2f}")
        out.append("```")
        out.append("")
        out.append(f"### ✅ Final Q1 FY2026 payout: **{_usd(final)}**")
        out.append("")
        out.append(
            "_Note: ENT is paid in **USD only** — no INR conversion is applied._"
        )
        return "\n".join(out)

    if is_ent_period:
        # Step-by-step calculation walkthrough — highest priority because it overlaps
        # with other ENT triggers (it also contains the word "commission").
        if any(
            t in q for t in (
                "step by step", "step-by-step", "step by-step",
                "walk me through", "walkthrough", "walk through",
                "explain the calculation", "explain calculation",
                "explain the ent", "explain ent",
                "explain the commission", "explain commission",
                "how is the commission calculated", "how was the commission calculated",
                "how is the ent commission", "how the commission",
                "in detail", "detailed breakdown", "detailed explanation",
                "show me the math", "show the math", "show the calculation",
            )
        ):
            return _ent_step_by_step_answer()
        # Q1: "Anthony's total Q1 FY2026 payout, including Q4 pending deals" — full breakdown.
        if ("anthony" in q) and any(
            t in q for t in ("total", "payout", "commission", "earn", "pending", "including", "summariz", "summary", "breakdown", "bullet")
        ):
            return _ent_anthony_total_answer()
        # Q2: "Which Q1 deals were Not Paid? How are they excluded from Payment Received?"
        # Also catches "paid deals" / "deals we received payment for" — the unpaid-deals
        # answer now shows both paid and unpaid for symmetry.
        if any(
            t in q for t in (
                "not paid", "not-paid", "unpaid", "non-paid", "non paid",
                "excluded from payment", "exclude from payment", "exclusion",
                "payment received", "received the payment", "received payment",
                "paid deals", "paid q1 deals", "paid deal list",
                "estée lauder", "estee lauder", "experian", "cloudsoft",
                "ezcater", "bunsow", "nfl - slack", "deel - spo", "tracelink",
                "marketcast", "convex insurance", "cumulus global",
            )
        ):
            return _ent_unpaid_deals_answer()
        # Q3: "What is the Cambrex refund and how is it deducted?"
        if any(
            t in q for t in (
                "cambrex", "refund", "refunded", "cancelled deal", "canceled deal",
                "cancellation", "cancelled", "canceled", "clawed back",
            )
        ):
            return _ent_refund_clawback_answer()

    # ------------------------------------------------------------------
    # SMB-specific intent handlers (Q1 SMB page suggested prompts)
    # ------------------------------------------------------------------
    # The SMB page has 7 commissioned reps + Chitradip as manager, 44 deals (with 1
    # Keytel cancellation), and 3 Q4 clawbacks. Scoped strictly to SMB.
    is_smb_period = ("smb" in team_label_low or team_label_low.startswith("smb")) and not is_ent_period and not (
        "account management" in team_label_low or "account mgmt" in team_label_low
    )

    def _smb_deal_counts_answer() -> str:
        """SMB Q1 deal counts: total / paid / not paid / cancelled with breakdown."""
        if not (paid_deals_block or unpaid_overrides):
            return f"No SMB deal-level data recorded for {period_lbl}."
        cancelled = [d for d in unpaid_overrides if d.get("is_cancelled")]
        non_cancelled_unpaid = [d for d in unpaid_overrides if not d.get("is_cancelled") and not d.get("is_partial_remainder")]
        paid_n = len(paid_deals_block)
        unpaid_n = len(non_cancelled_unpaid)
        cancelled_n = len(cancelled)
        total = paid_n + unpaid_n  # effective (cancelled deals don't count)
        paid_sum = sum(float(d.get("amount_usd") or 0) for d in paid_deals_block)
        unpaid_sum = sum(float(d.get("amount_usd") or 0) for d in non_cancelled_unpaid)
        cancelled_sum = sum(float(d.get("amount_usd") or 0) for d in cancelled)
        official_team_paid = float((pinned.get("manager_incentive") or {}).get("team_payment_received_usd") or paid_sum)
        official_team_booked = float(pinned.get("team_achievement_usd") or (paid_sum + unpaid_sum))

        out = [
            f"### {period_lbl} — SMB team deal counts",
            "",
            f"- 📦 **Total effective deals closed**: **{total}** (paid + not received; cancellations excluded)",
            f"- ✅ **Payment received**: **{paid_n} deal(s)** — total **{_usd(official_team_paid)}** (official)",
            f"- ❌ **Payment not received**: **{unpaid_n} deal(s)** — excluded value **{_usd(unpaid_sum)}**",
        ]
        if cancelled_n:
            out.append(f"- ⚠️ **Cancelled (was Closed Won, later cancelled)**: **{cancelled_n} deal(s)** — {_usd(cancelled_sum)} removed from team achievement")
        out.append("")
        out.append("**Team-level summary**")
        out.append(f"- Booked Amount (after cancellations): **{_usd(official_team_booked)}**")
        out.append(f"- Payment Received: **{_usd(official_team_paid)}**")
        out.append(f"- Pending: **{_usd(unpaid_sum)}**")
        if cancelled_n:
            out.append("")
            out.append("**Cancelled deals**")
            for d in cancelled:
                dn = _esc_md((d.get("deal_name") or "").strip())
                owner = _esc_md((d.get("deal_owner") or "").strip())
                amt = float(d.get("amount_usd") or 0)
                out.append(f"- **{dn}** — Owner: {owner} · {_usd(amt)}")
        return "\n".join(out)

    def _smb_deal_owner_answer() -> str:
        """SMB Q1 deal-owner breakdown using official rep figures for Booked and Received."""
        if not (paid_deals_block or unpaid_overrides):
            return f"No SMB deal-owner data recorded for {period_lbl} AM."
        # Build rep lookup
        rep_lookup = {(r.get("name") or "").strip().lower(): r for r in (pinned.get("reps") or [])}
        owner_alias = {"yogesh vig": "yogi"}

        owners: dict[str, dict] = {}
        for d in paid_deals_block:
            o = owners.setdefault(d.get("deal_owner") or "—", {"paid_n": 0, "unpaid_n": 0, "cancelled_n": 0, "booked_deal": 0.0, "paid_deal": 0.0, "pending_deal": 0.0, "cancelled_sum": 0.0})
            o["paid_n"] += 1
            o["paid_deal"] += float(d.get("amount_usd") or 0)
            o["booked_deal"] += float(d.get("amount_usd") or 0)
        for d in unpaid_overrides:
            o = owners.setdefault(d.get("deal_owner") or "—", {"paid_n": 0, "unpaid_n": 0, "cancelled_n": 0, "booked_deal": 0.0, "paid_deal": 0.0, "pending_deal": 0.0, "cancelled_sum": 0.0})
            amt = float(d.get("amount_usd") or 0)
            if d.get("is_cancelled"):
                o["cancelled_n"] += 1
                o["cancelled_sum"] += amt
            else:
                o["unpaid_n"] += 1
                o["booked_deal"] += amt
                o["pending_deal"] += amt

        out = [
            f"### {period_lbl} — SMB deals by owner",
            "",
            "_Booked Amount and $ Payment received use the **official rep totals** from the JSON (rounded to whole dollars). $ Pending sums from the deal-level data._",
            "",
            "| Deal Owner | Total | Payment received | # Payment not received | Booked Amount | $ Payment received | $ Payment Pending |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
        sum_total = sum_paid = sum_unpaid = 0
        sum_pending = 0.0
        for owner, o in sorted(owners.items(), key=lambda kv: (-(kv[1]["paid_n"] + kv[1]["unpaid_n"]), kv[0])):
            tot = o["paid_n"] + o["unpaid_n"]
            # Look up rep by normalized name (strip "(Deactivated User)" and apply aliases)
            base_name = (owner or "").replace(" (Deactivated User)", "").strip().lower()
            base_name = owner_alias.get(base_name, base_name)
            rep = rep_lookup.get(base_name)
            if rep:
                booked_disp = round(float(rep.get("achievement_usd") or 0))
                paid_disp = float(rep.get("payment_received_usd") or 0)
            else:
                booked_disp = o["booked_deal"]
                paid_disp = o["paid_deal"]
            sum_total += tot
            sum_paid += o["paid_n"]
            sum_unpaid += o["unpaid_n"]
            sum_pending += o["pending_deal"]
            out.append(
                f"| **{_esc_md(owner)}** | {tot} | {o['paid_n']} | {o['unpaid_n']} | "
                f"{_usd(booked_disp)} | {_usd(paid_disp)} | {_usd(o['pending_deal'])} |"
            )
        official_total_booked = float(pinned.get("team_achievement_usd") or 0)
        official_total_paid = float((pinned.get("manager_incentive") or {}).get("team_payment_received_usd") or 0)
        out.append(
            f"| **Total (SMB team)** | **{sum_total}** | **{sum_paid}** | **{sum_unpaid}** | "
            f"**{_usd(official_total_booked)}** | **{_usd(official_total_paid)}** | **{_usd(sum_pending)}** |"
        )
        return "\n".join(out)

    def _smb_keytel_cancellation_answer() -> str:
        """Explain the Keytel Systems cancellation and its impact on Yogi & team."""
        # Find the cancelled deal (if any)
        cancelled = [d for d in unpaid_overrides if d.get("is_cancelled")]
        if not cancelled:
            return f"No cancelled deals recorded for {period_lbl} SMB."
        out = [f"### {period_lbl} — Cancelled deals and their impact", ""]
        team_booked = float(pinned.get("team_achievement_usd") or 0)
        for d in cancelled:
            dn = _esc_md((d.get("deal_name") or "").strip())
            owner_raw = (d.get("deal_owner") or "").strip()
            owner = _esc_md(owner_raw)
            amt = float(d.get("amount_usd") or 0)
            cd = _esc_md((d.get("close_date") or "").strip())
            status = _esc_md((d.get("payment_status") or "Cancelled").strip())
            reason = _esc_md((d.get("reason") or "").strip())
            pre_cancel_team = team_booked + amt
            out.append(f"**{dn}**")
            out.append(f"- **Owner**: {owner}")
            out.append(f"- **Amount**: {_usd(amt)}")
            out.append(f"- **Status**: {status} (Closed Won on {cd})")
            if reason:
                out.append(f"- _{reason}_")
            out.append("")
            out.append("**Impact on the team**")
            out.append(f"- Original Q1 Booked Amount (with the deal): {_usd(pre_cancel_team)}")
            out.append(f"- After cancellation: **{_usd(team_booked)}** (= {_usd(pre_cancel_team)} − {_usd(amt)})")
            # Identify the affected rep
            base_owner = owner_raw.replace(" (Deactivated User)", "").strip().lower()
            owner_alias = {"yogesh vig": "yogi"}
            base_owner = owner_alias.get(base_owner, base_owner)
            rep_lookup = {(r.get("name") or "").strip().lower(): r for r in (pinned.get("reps") or [])}
            rep = rep_lookup.get(base_owner)
            if rep:
                rep_name = (rep.get("name") or "").strip()
                rep_ach = float(rep.get("achievement_usd") or 0)
                rep_pre = rep_ach + amt
                quota = float(rep.get("target_usd") or 0)
                pre_pct = (rep_pre / quota * 100.0) if quota else 0.0
                post_pct = (rep_ach / quota * 100.0) if quota else 0.0
                out.append(f"- **{rep_name}'s individual achievement** drops from {_usd(rep_pre)} ({pre_pct:.1f}% of quota) to **{_usd(rep_ach)}** ({post_pct:.1f}% of quota).")
            out.append("")
        out.append(
            "_Cancellations remove the deal value from the team Booked Amount, the rep's individual achievement, "
            "and the bullet chart percentage. Payment Received is unaffected because no money was ever collected._"
        )
        return "\n".join(out)

    def _smb_chitradip_manager_stepbystep_answer() -> str:
        """Detailed step-by-step walkthrough of Chitradip's SMB manager commission.

        Period-agnostic — works for Q1, April, May, or any other pinned SMB period.
        Handles base commission, pending commissions from prior periods, clawbacks,
        and the 60% team floor for monthly plans.
        """
        mgr = pinned.get("manager_incentive") or {}
        team_target = float(mgr.get("team_target_usd") or pinned.get("team_target_usd") or 0)
        team_ach = float(mgr.get("team_achievement_usd") or pinned.get("team_achievement_usd") or 0)
        team_ach_pct = float(mgr.get("team_achievement_pct") or 0) or (
            (team_ach / team_target * 100.0) if team_target else 0.0
        )
        team_paid = float(mgr.get("team_payment_received_usd") or 0)
        rate_pct = float(mgr.get("commission_rate_pct") or 1.5)
        mgr_base = float(mgr.get("manager_q1_amount_usd") or 0) or round(team_paid * rate_pct / 100.0, 2)
        pending_items = mgr.get("pending_commissions") or []
        pending_total = float(mgr.get("total_pending_commission_usd") or sum(float(p.get("commission_usd") or 0) for p in pending_items))
        clawback_items = mgr.get("clawbacks") or []
        clawback_total = float(mgr.get("total_clawback_usd") or sum(float(c.get("deduction_usd") or 0) for c in clawback_items))
        final = float(mgr.get("final_payout_usd")) if mgr.get("final_payout_usd") is not None else (mgr_base + pending_total - clawback_total)
        # For monthly plans the manager floor is 60% (team_minimum); for quarterly the base rule differs
        is_monthly = bool(pinned.get("month"))
        mgr_min = int((pinned.get("monthly_tiers") or {}).get("manager_team_minimum_pct") or mgr.get("manager_team_minimum_pct") or (60 if is_monthly else 0))
        is_eligible = team_ach_pct >= mgr_min if mgr_min else True

        step = 1
        out = [
            f"## {period_lbl} — Chitradip's SMB Manager Commission, Step by Step",
            "",
            f"Chitradip is the SMB manager. His **{period_lbl}** commission is calculated on the team's **Payment Received** at a flat "
            f"**{rate_pct:g}%** manager rate."
            + (f" A **{mgr_min}% team-achievement floor** applies for monthly plans." if is_monthly and mgr_min else "")
            + (" Pending commissions from prior periods and any clawbacks are then applied to reach the final payout." if (pending_items or clawback_items) else ""),
            "",
            "---",
            "",
            f"### Step {step} — Pull the team's {period_lbl} performance",
            "",
            f"- **Team target**: {_usd(team_target)}",
            f"- **Team achievement (booked)**: {_usd(team_ach)} → **{team_ach_pct:.1f}% of quota**",
            f"- **Team payment received** (commission base): **{_usd(team_paid)}**",
            "",
            "---",
            "",
        ]
        step += 1
        # Step: eligibility
        out.append(f"### Step {step} — Eligibility check")
        out.append("")
        if is_monthly and mgr_min:
            if is_eligible:
                out.append(
                    f"Monthly plan applies a **{mgr_min}% team-achievement floor** for the manager base. "
                    f"Team is at **{team_ach_pct:.1f}%**, which is **≥ {mgr_min}%** → ✅ Chitradip is **eligible** for base commission."
                )
            else:
                out.append(
                    f"Monthly plan applies a **{mgr_min}% team-achievement floor** for the manager base. "
                    f"Team is at **{team_ach_pct:.1f}%**, which is **below {mgr_min}%** → ❌ Chitradip earns **$0 base commission** for {period_lbl}. "
                    "Prior-period pending commissions (if any) are still paid — they aren't subject to the current-period floor."
                )
        else:
            out.append(f"Quarterly plan — Chitradip is eligible for base commission on team payment received at **{rate_pct:g}%**.")
        out.append("")
        out.append("---")
        out.append("")
        step += 1
        # Step: base
        out.append(f"### Step {step} — Calculate {period_lbl} base manager commission")
        out.append("")
        out.append("```")
        if is_monthly and mgr_min and not is_eligible:
            out.append(f"Base = $0 (team below {mgr_min}% floor)")
        else:
            out.append(f"Base = Team Payment Received × {rate_pct:g}%")
            out.append(f"     = {team_paid:,.2f} × {rate_pct:g}%")
            out.append(f"     = {mgr_base:,.2f}")
        out.append("```")
        out.append("")
        _effective_base = mgr_base if is_eligible else 0.0
        out.append(f"**{period_lbl} base commission: {_usd(_effective_base)}**")
        out.append("")
        out.append("---")
        out.append("")
        step += 1
        # Step: pending commissions
        if pending_items:
            out.append(f"### Step {step} — Add prior-period pending commissions")
            out.append("")
            out.append(
                "Pending commissions are prior-period deals whose payment was received in this cycle. "
                "They're paid at the ORIGINAL period's rate (Q4 = 1%, Q1 = 1.5%, etc.) — NOT the current rate. "
                "**They are NOT subject to the current-period team floor.**"
            )
            out.append("")
            out.append("| Deal | Prior period | Deal amount | Rate | Commission |")
            out.append("| --- | --- | ---: | ---: | ---: |")
            for p in pending_items:
                lbl = _esc_md((p.get("deal_name") or "").strip())
                per = _esc_md((p.get("period") or "").strip())
                deal_amt = float(p.get("deal_amount_usd") or 0)
                rate = float(p.get("rate_pct") or 0)
                comm = float(p.get("commission_usd") or 0)
                out.append(f"| **{lbl}** | {per} | {_usd(deal_amt)} | {rate:g}% | **+{_usd(comm)}** |")
            out.append("")
            out.append(f"**Subtotal pending: +{_usd(pending_total)}**")
            out.append("")
            out.append("---")
            out.append("")
            step += 1
        # Step: clawbacks
        if clawback_items:
            out.append(f"### Step {step} — Apply prior-period clawback adjustments")
            out.append("")
            out.append(
                "Clawbacks are prior-period corrections that reduce Chitradip's current commission — "
                "deals that were counted but never collected, or slab over-payments that need correction."
            )
            out.append("")
            out.append("| Clawback | Deal amount | Rate | Deduction |")
            out.append("| --- | ---: | ---: | ---: |")
            for cb in clawback_items:
                lbl = _esc_md((cb.get("label") or cb.get("deal_name") or "").strip())
                deal_amt = float(cb.get("deal_amount_usd") or 0)
                rate = float(cb.get("rate_pct") or 0)
                ded = float(cb.get("deduction_usd") or 0)
                note = _esc_md((cb.get("note") or "").strip())
                out.append(f"| **{lbl}** | {_usd(deal_amt)} | {rate:g}% | **−{_usd(ded)}** |")
                if note:
                    out.append(f"| _{note}_ |  |  |  |")
            out.append("")
            out.append(f"**Subtotal clawback: −{_usd(clawback_total)}**")
            out.append("")
            out.append("---")
            out.append("")
            step += 1
        # Final add-up
        out.append(f"### Step {step} — Add it all up")
        out.append("")
        out.append("```")
        out.append(f"{period_lbl} base            {_effective_base:>12,.2f}")
        if pending_items:
            out.append(f"+ Prior-period pending   {pending_total:>12,.2f}")
        if clawback_items:
            out.append(f"− Prior-period clawbacks {clawback_total:>12,.2f}")
        out.append(f"───────────────────────────────────────────")
        out.append(f"Final {period_lbl} payout      {final:>12,.2f}")
        out.append("```")
        out.append("")
        _inr_txt = f" (≈ {_inr(final)})" if inr_rate else ""
        icon = "✅" if final > 0 else ("⚠️" if final == 0 else "❌")
        out.append(f"### {icon} Final {period_lbl} payout: **{_usd(final)}**{_inr_txt}")
        return "\n".join(out)

    def _smb_rep_detail_answer(rep_first_name: str) -> str:
        """Detailed per-rep SMB commission answer for the current period (Q1/April/May).

        Covers performance vs quota, eligibility check, base commission math,
        manage-deal incentive (if any), prior-period adjustments (Yogi's Yieldstreet
        in April, Lawrence's 2 manage deals in Q1, etc.), left_org note, and any
        total_payout_override_usd from the pinned JSON.
        """
        comm_reps = (pinned.get("commission") or {}).get("reps") or []
        rep_top_list = pinned.get("reps") or []
        elig_floor = int((pinned.get("commission") or {}).get("eligibility_min_pct") or 50)
        is_monthly_p = bool(pinned.get("month"))

        rep = next(
            (r for r in comm_reps if (r.get("name") or "").lower().startswith(rep_first_name.lower())),
            None,
        )
        rep_top = next(
            (r for r in rep_top_list if (r.get("name") or "").lower().startswith(rep_first_name.lower())),
            None,
        )
        if not rep and not rep_top:
            return f"No data found for {rep_first_name.title()} in {period_lbl} SMB."

        # Pull figures from whichever source has them (rep_top for monthly, rep for quarterly)
        rep_name = ((rep or {}).get("name") or (rep_top or {}).get("name") or rep_first_name.title()).strip()
        quota = float((rep or {}).get("quota_usd") or (rep_top or {}).get("target_usd") or 0)
        booked = float((rep or {}).get("revenue_achieved_usd") or (rep_top or {}).get("achievement_usd") or 0)
        paid_recv = float((rep or {}).get("payment_received_usd") or (rep_top or {}).get("payment_received_usd") or 0)
        ach_pct = (booked / quota * 100.0) if quota else 0.0
        if rep and rep.get("eligible_pct") is not None:
            ach_pct = float(rep.get("eligible_pct") or 0)
        base = float((rep or {}).get("base_commission_usd") or 0)
        ded_pct = float((rep or {}).get("deduction_pct") or 0)
        ded_amt = float((rep or {}).get("deduction_usd") or 0)
        md_total = float((rep or {}).get("manage_deal_usd") or (rep_top or {}).get("manage_deal_usd") or 0)
        md_paid_now = float((rep or {}).get("manage_deal_paid_now_usd") or (rep_top or {}).get("manage_deal_paid_now_usd") or 0)
        md_pending = max(md_total - md_paid_now, 0)
        payout_override = (rep or {}).get("total_payout_override_usd") if rep else (rep_top or {}).get("total_payout_override_usd")
        payout = float(payout_override) if payout_override is not None else float((rep or {}).get("total_payout_usd") or 0)
        left_org = bool((rep_top or {}).get("left_org") or (rep or {}).get("left_org"))
        adjustments = (rep_top or {}).get("adjustments") or (rep or {}).get("adjustments") or []
        clawback_usd = float((rep_top or {}).get("clawback_usd") or (rep or {}).get("clawback_usd") or 0)
        clawback_note = ((rep_top or {}).get("clawback_note") or (rep or {}).get("clawback_note") or "").strip()
        note_v = ((rep_top or {}).get("note") or (rep or {}).get("note") or "").strip()
        eligible = ach_pct >= elig_floor

        rep_paid_deals = [d for d in paid_deals_block if _rep_name_in_owner((d.get("deal_owner") or ""), rep_first_name)]
        rep_unpaid_deals = [d for d in unpaid_overrides if _rep_name_in_owner((d.get("deal_owner") or ""), rep_first_name)]

        out = [
            f"## {rep_name} — {period_lbl} commission breakdown",
            "",
            f"This is the full step-by-step calculation for **{rep_name}** in {period_lbl}. All numbers are from the pinned JSON.",
            "",
            "---",
            "",
            "### Step 1 — Performance vs quota",
            "",
            f"- **Individual quota**: {_usd(quota)}",
            f"- **Revenue achieved (booked)**: {_usd(booked)} → **{ach_pct:.1f}% of quota**",
            f"- **Payment received**: {_usd(paid_recv)}",
            "",
        ]
        if left_org:
            out.append(f"_Note: {rep_name} left the organization — his deals still count toward team total, but no personal payout is processed._")
            out.append("")
        out.append("---")
        out.append("")
        out.append("### Step 2 — Eligibility check")
        out.append("")
        _floor_note = ""
        if is_monthly_p and elig_floor == 60:
            _floor_note = " The monthly plan applies a **strict 60% floor** — anything below (even 59.9%) earns 0% base commission."
        if eligible:
            out.append(
                f"{rep_name} is at **{ach_pct:.1f}%** of quota → meets the **{elig_floor}% eligibility floor** → ✅ **eligible** for base commission."
                + _floor_note
            )
        else:
            out.append(
                f"Achievement is **{ach_pct:.1f}%** — **below the {elig_floor}% eligibility floor** → ❌ **not eligible** for base commission. Base = **$0**."
                + _floor_note
            )
        out.append("")
        out.append("---")
        out.append("")
        # Step: base
        out.append("### Step 3 — Base commission")
        out.append("")
        if eligible:
            _tier = "monthly slab" if is_monthly_p else "SMB slab"
            slab_pct = float((rep or {}).get("slab_pct") or 0) or (round(base / paid_recv * 100.0, 2) if paid_recv else 0.0)
            out.append("```")
            out.append(f"Base = Payment Received × {_tier} rate")
            out.append(f"     = {paid_recv:,.2f} × {slab_pct:g}%")
            out.append(f"     = {base:,.2f}")
            out.append("```")
        else:
            out.append("Not eligible — base = **$0**.")
        out.append("")
        if ded_amt > 0:
            out.append("---")
            out.append("")
            out.append(f"### Step 4 — Apply the {int(ded_pct)}% deduction (no manage deal closed)")
            out.append("")
            out.append("SMB policy: reps who didn't close a Q1 manage deal incur a **10% deduction** on their achieved commission.")
            out.append("")
            out.append("```")
            out.append(f"Deduction = Base × {int(ded_pct)}%")
            out.append(f"          = {base:,.2f} × {int(ded_pct)}%")
            out.append(f"          = {ded_amt:,.2f}")
            out.append("```")
            out.append("")
        if md_total > 0:
            out.append("---")
            out.append("")
            out.append("### Manage deal incentive")
            out.append("")
            if md_paid_now > 0 and md_paid_now < md_total:
                deals_count = int(round(md_total / 500))
                paid_count = int(round(md_paid_now / 500))
                out.append(
                    f"{rep_name} closed **{deals_count} manage deal(s)** at $500 each → **{_usd(md_total)} earned total**."
                )
                out.append(
                    f"At the time {period_lbl} commissions were calculated, payment was received for **{paid_count}** of those deals → **{_usd(md_paid_now)}** paid this cycle."
                )
                out.append(f"The remaining **{_usd(md_pending)}** is **pending next cycle**.")
            else:
                out.append(f"Manage deal incentive: **{_usd(md_total)}**.")
            out.append("")
            out.append("_Manage deal incentive is paid regardless of overall quota achievement — separate from the base-commission eligibility floor._")
            out.append("")
        if adjustments:
            out.append("---")
            out.append("")
            out.append("### Prior-period adjustments")
            out.append("")
            out.append("| Deal | Prior period | Paid amount | Rate | Commission |")
            out.append("| --- | --- | ---: | ---: | ---: |")
            for a in adjustments:
                dn = _esc_md((a.get("deal_name") or "").strip())
                per = _esc_md((a.get("period") or "").strip())
                pa = float(a.get("paid_amount_usd") or a.get("deal_amount_usd") or 0)
                rate = float(a.get("rate_pct") or 0)
                comm = float(a.get("commission_usd") or 0)
                sign = "+" if comm >= 0 else "−"
                out.append(f"| **{dn}** | {per} | {_usd(pa)} | {rate:g}% | **{sign}{_usd(abs(comm))}** |")
            out.append("")
        if clawback_usd > 0:
            out.append("---")
            out.append("")
            out.append(f"### Clawback adjustment: −{_usd(clawback_usd)}")
            out.append("")
            if clawback_note:
                out.append(f"_{_esc_md(clawback_note)}_")
                out.append("")
        # Final
        out.append("---")
        out.append("")
        out.append("### Final payout math")
        out.append("")
        out.append("```")
        _base_eff = base if eligible else 0.0
        out.append(f"Base commission           {_base_eff:>12,.2f}")
        if ded_amt > 0:
            out.append(f"− {int(ded_pct)}% deduction            {ded_amt:>12,.2f}")
        if md_paid_now > 0:
            out.append(f"+ Manage deal paid now    {md_paid_now:>12,.2f}")
        if adjustments:
            adj_sum = sum(float(a.get("commission_usd") or 0) for a in adjustments)
            _sign = "+" if adj_sum >= 0 else "−"
            out.append(f"{_sign} Prior-period adjustments {abs(adj_sum):>12,.2f}")
        if clawback_usd > 0:
            out.append(f"− Clawback adjustment     {clawback_usd:>12,.2f}")
        out.append(f"───────────────────────────────────────────")
        out.append(f"Final {period_lbl} payout       {payout:>12,.2f}")
        out.append("```")
        out.append("")
        icon = "✅" if payout > 0 else ("⚠️" if payout == 0 else "❌")
        _inr_txt = f" (≈ {_inr(payout)})" if inr_rate else ""
        out.append(f"### {icon} Final {period_lbl} payout: **{_usd(payout)}**{_inr_txt}")
        if note_v:
            out.append("")
            out.append(f"_{_esc_md(note_v)}_")
        # Rep's deals for context
        if rep_paid_deals or rep_unpaid_deals:
            out.append("")
            out.append("---")
            out.append("")
            out.append(f"### {rep_name}'s {period_lbl} deals")
            out.append("")
            out.append("| Deal | Close date | Status | Amount |")
            out.append("| --- | --- | --- | ---: |")
            for d in rep_paid_deals + rep_unpaid_deals:
                dn = _esc_md((d.get("deal_name") or "").strip())
                cd = _esc_md((d.get("close_date") or "").strip())
                status = _esc_md((d.get("payment_status") or "").strip())
                amt = float(d.get("amount_usd") or 0)
                out.append(f"| **{dn}** | {cd} | {status} | {_usd(amt)} |")
        return "\n".join(out)

    def _rep_name_in_owner(owner: str, first_name: str) -> bool:
        low = (owner or "").lower().replace(" (deactivated user)", "").strip()
        fn = (first_name or "").lower()
        if not fn:
            return False
        return fn in low or low.startswith(fn) or fn == low.split(" ")[0]

    def _smb_eligibility_answer() -> str:
        """Who is / isn't eligible for SMB commission this period."""
        comm_reps = (pinned.get("commission") or {}).get("reps") or []
        rep_top_list = pinned.get("reps") or []
        elig_floor = int((pinned.get("commission") or {}).get("eligibility_min_pct") or 50)
        is_monthly_p = bool(pinned.get("month"))
        out = [
            f"### {period_lbl} — SMB rep eligibility check",
            "",
        ]
        if is_monthly_p and elig_floor == 60:
            out.append(
                f"**Rule:** Monthly plan applies a **strict {elig_floor}% floor** — "
                f"anything below {elig_floor}% (even 59.9%) earns **0%** base commission."
            )
        else:
            out.append(f"**Rule:** A minimum **{elig_floor}%** of the individual quota is required for base commission.")
        out.append("")
        eligible_rows: list[str] = []
        ineligible_rows: list[str] = []
        # Prefer rep_top_list (has quota + achievement); fall back to comm_reps.
        source = rep_top_list if rep_top_list else [
            {"name": c.get("name"), "target_usd": c.get("quota_usd"), "achievement_usd": c.get("revenue_achieved_usd")}
            for c in comm_reps
        ]
        for r in source:
            nm = (r.get("name") or "").strip()
            if not nm:
                continue
            tgt = float(r.get("target_usd") or r.get("quota_usd") or 0)
            ach = float(r.get("achievement_usd") or r.get("revenue_achieved_usd") or 0)
            pct = (ach / tgt * 100.0) if tgt else 0.0
            left_org = bool(r.get("left_org"))
            if left_org:
                ineligible_rows.append(f"- **{nm}** — Left the organization → **not eligible**")
            elif pct >= elig_floor:
                eligible_rows.append(f"- **{nm}** — {pct:.1f}% of {_usd(tgt)} → ✅ **eligible**")
            else:
                ineligible_rows.append(f"- **{nm}** — {pct:.1f}% of {_usd(tgt)} → ❌ below {elig_floor}% floor")
        out.append("**✅ Eligible reps:**")
        out.append("")
        out.extend(eligible_rows or ["- (none)"])
        out.append("")
        out.append("**❌ Not eligible:**")
        out.append("")
        out.extend(ineligible_rows or ["- (none)"])
        return "\n".join(out)

    def _smb_period_summary_answer() -> str:
        """3-bullet summary of the SMB period."""
        mgr = pinned.get("manager_incentive") or {}
        team_target = float(pinned.get("team_target_usd") or 0)
        team_ach = float(pinned.get("team_achievement_usd") or 0)
        team_pct = (team_ach / team_target * 100.0) if team_target else 0.0
        team_paid = float(mgr.get("team_payment_received_usd") or 0)
        mgr_final = float(mgr.get("final_payout_usd") or 0)
        deals_summary = pinned.get("deals_summary") or {}
        total_deals = int(deals_summary.get("total_deals_closed") or 0)
        received_deals = int(deals_summary.get("deals_payment_received") or 0)
        not_received_deals = int(deals_summary.get("deals_payment_not_received") or 0)
        return "\n".join([
            f"### {period_lbl} — SMB in 3 bullets",
            "",
            f"1. **Team performance**: Booked {_usd(team_ach)} against target {_usd(team_target)} = **{team_pct:.1f}%**. "
            f"Payment received (commission base): {_usd(team_paid)}. "
            f"{total_deals} deals closed ({received_deals} paid, {not_received_deals} not paid).",
            "",
            f"2. **Chitradip's manager payout**: **{_usd(mgr_final)}**"
            + (f" ({mgr.get('calculation_note') or ''})" if mgr.get('calculation_note') else "")
            + ".",
            "",
            f"3. **Rep commissions**: See the Commission summary table on this page for each rep's slab, base, adjustments, and final payout. "
            f"Ask \"Explain [rep name]'s {period_lbl} commission\" for a step-by-step breakdown.",
        ])

    if is_smb_period:
        # SMB-0: Chitradip manager step-by-step
        if (
            ("chitradip" in q and any(t in q for t in (
                "step by step", "step-by-step", "walk through", "walkthrough",
                "manager commission", "manager's commission", "manager incentive",
                "explain", "calculation", "in detail", "clawback", "clawbacks",
            )))
            or "noventiq" in q
            or "synergy employment" in q
            or "ceg solutions" in q
            or "yieldstreet" in q
            or "willow wealth" in q
            or "legal soft" in q
            or ("3 clawback" in q or "three clawback" in q)
        ):
            return _smb_chitradip_manager_stepbystep_answer()
        # SMB-0.5: individual rep detail — routes BEFORE deal-owner breakdown so single-rep
        # questions get a personal answer rather than the all-reps table.
        _smb_rep_names_in_q = {
            n for n in ("vicky", "yogi", "yogesh", "kritika", "lawrence", "larry", "royston", "kartik", "deepak", "rutuja", "lennis")
            if n in q
        }
        # Trigger single-rep answer when exactly ONE rep is mentioned AND question is
        # about their commission (not just a general "vicky's deals" query).
        if len(_smb_rep_names_in_q) == 1 and any(
            t in q for t in (
                "commission", "payout", "eligibility", "eligible", "explain",
                "calculation", "why", "how did", "how does", "step by step",
                "step-by-step", "leave the org", "left the org", "left the organi",
                "manage deal", "2 manage", "two manage",
            )
        ):
            _first_name = next(iter(_smb_rep_names_in_q))
            # Map "larry" → "lawrence", "yogesh" → "yogi" for rep lookup
            _first_name = {"larry": "lawrence", "yogesh": "yogi"}.get(_first_name, _first_name)
            return _smb_rep_detail_answer(_first_name)
        # SMB-A: deal counts
        if any(
            t in q for t in (
                "how many deals", "how many deal", "deal count", "deal counts",
                "number of deals", "number of deal", "total deals closed",
                "deals closed", "deal closed", "closed and how many",
            )
        ):
            return _smb_deal_counts_answer()
        # SMB-B: deal owner breakdown
        if any(
            t in q for t in (
                "deal owner", "by owner", "per owner", "owner breakdown",
                "all owners", "every owner", "each owner",
                "break down deals", "breakdown deals", "deals by",
            )
        ) or len(_smb_rep_names_in_q) > 1:
            return _smb_deal_owner_answer()
        # SMB-C: Keytel cancellation (Q1-only)
        if any(t in q for t in ("keytel", "keytel systems")):
            return _smb_keytel_cancellation_answer()
        # SMB-D: eligibility check
        if any(t in q for t in ("who is eligible", "who's eligible", "not eligible", "eligibility", "60% floor", "eligible for smb", "eligible for commission")):
            return _smb_eligibility_answer()
        # SMB-E: 3-bullet summary
        if any(t in q for t in ("summarize", "summary", "3 bullets", "three bullets", "key insights")):
            return _smb_period_summary_answer()

    # ------------------------------------------------------------------
    # AM-specific intent handlers (Q1 AM page suggested prompts)
    # ------------------------------------------------------------------
    # The AM page now has paid_deals / unpaid_deals_overrides arrays, plus a Joy-as-
    # both-rep-and-manager scenario with Wipro clawback, the Artnet partial payment,
    # and now 37 total deals across 4 deal owners. Strictly scoped to AM.
    is_am_period = ("account management" in team_label_low or "account mgmt" in team_label_low) and not is_ent_period
    deals_summary_pinned = pinned.get("deals_summary") or {}

    def _am_deal_counts_answer() -> str:
        """Total / paid / not paid deal counts for AM Q1."""
        if not (paid_deals_block or unpaid_overrides):
            return f"No deal-level data is recorded for {period_lbl} AM."
        partial_paid = [d for d in paid_deals_block if d.get("is_partial")]
        partial_remainder = [d for d in unpaid_overrides if d.get("is_partial_remainder")]
        any_payment_count = len(paid_deals_block)  # includes partial-paid
        fully_unpaid_count = sum(1 for d in unpaid_overrides if not d.get("is_partial_remainder"))
        total_closed = int(deals_summary_pinned.get("total_deals_closed") or (
            any_payment_count - len(partial_paid) + fully_unpaid_count + len(partial_paid)
        ))
        paid_sum = sum(float(d.get("amount_usd") or 0) for d in paid_deals_block)
        unpaid_sum = sum(float(d.get("amount_usd") or 0) for d in unpaid_overrides)
        return (
            f"### {period_lbl} — AM team deal counts\n\n"
            f"- 📦 **Total deals closed**: **{total_closed}**\n"
            f"- ✅ **Payment received**: **{any_payment_count} deal(s)** — total {_usd(paid_sum)}"
            + (f" _(includes {len(partial_paid)} partially-paid deal — Artnet)_" if partial_paid else "")
            + "\n"
            f"- ❌ **Payment not received**: **{fully_unpaid_count} deal(s)** — excluded value {_usd(unpaid_sum)}"
            + (f"\n  _(the Artnet partial remainder of ${(partial_remainder[0].get('amount_usd') or 0):,.2f} is also pending — see Artnet for details)_" if partial_remainder else "")
            + "\n\n"
            f"Per ENT-style policy on Payment Received, only deals with payment received contribute to the commission base. "
            f"The {fully_unpaid_count} Not-Paid deal(s) plus the partial Artnet remainder are excluded until payment arrives."
        )

    def _am_deal_owner_answer() -> str:
        """Deal-owner breakdown for AM Q1 — same columns as the page table:
        Deal Owner | Total | Payment received deals | # Payment not received | Booked Amount | $ Received | $ Pending."""
        if not (paid_deals_block or unpaid_overrides):
            return f"No deal-owner data recorded for {period_lbl} AM."
        owners: dict[str, dict] = {}
        for d in paid_deals_block:
            o = owners.setdefault(d.get("deal_owner") or "—", {
                "paid_n": 0, "unpaid_n": 0, "partial_paid_n": 0, "partial_unpaid_n": 0,
                "booked": 0.0, "paid_usd": 0.0, "pending_usd": 0.0,
            })
            o["paid_n"] += 1
            paid_amt = float(d.get("amount_usd") or 0)
            o["paid_usd"] += paid_amt
            is_partial = bool(d.get("is_partial"))
            total_v = float(d.get("total_deal_value_usd") or 0) if is_partial else paid_amt
            o["booked"] += total_v if is_partial else paid_amt
            if is_partial:
                o["partial_paid_n"] += 1
                o["pending_usd"] += max(total_v - paid_amt, 0)
        for d in unpaid_overrides:
            o = owners.setdefault(d.get("deal_owner") or "—", {
                "paid_n": 0, "unpaid_n": 0, "partial_paid_n": 0, "partial_unpaid_n": 0,
                "booked": 0.0, "paid_usd": 0.0, "pending_usd": 0.0,
            })
            amt = float(d.get("amount_usd") or 0)
            if d.get("is_partial_remainder"):
                # Don't double-count booked/pending — already added on the paid side
                o["partial_unpaid_n"] += 1
            else:
                o["unpaid_n"] += 1
                o["booked"] += amt
                o["pending_usd"] += amt

        out: list[str] = [
            f"### {period_lbl} — AM deals by owner",
            "",
            "| Deal Owner | Total | Payment received deals | # Payment not received | Booked Amount | $ Payment received | $ Payment Pending |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
        sum_total = sum_paid = sum_unpaid = 0
        sum_partial_paid = sum_partial_unpaid = 0
        sum_booked = sum_paid_usd = sum_pending_usd = 0.0
        for owner, o in sorted(owners.items(), key=lambda kv: -(kv[1]["paid_n"] + kv[1]["unpaid_n"])):
            tot = o["paid_n"] + o["unpaid_n"]
            paid_cell = f"{o['paid_n']}" + (f" ({o['partial_paid_n']}-Partial Paid)" if o["partial_paid_n"] else "")
            unpaid_cell = f"{o['unpaid_n']}" + (f" ({o['partial_unpaid_n']}-Partial Payment not received)" if o["partial_unpaid_n"] else "")
            out.append(
                f"| **{_esc_md(owner)}** | {tot} | {paid_cell} | {unpaid_cell} | "
                f"{_usd(o['booked'])} | {_usd(o['paid_usd'])} | {_usd(o['pending_usd'])} |"
            )
            sum_total += tot
            sum_paid += o["paid_n"]
            sum_unpaid += o["unpaid_n"]
            sum_partial_paid += o["partial_paid_n"]
            sum_partial_unpaid += o["partial_unpaid_n"]
            sum_booked += o["booked"]
            sum_paid_usd += o["paid_usd"]
            sum_pending_usd += o["pending_usd"]
        total_paid_cell = f"**{sum_paid}**" + (f" ({sum_partial_paid}-Partial Paid)" if sum_partial_paid else "")
        total_unpaid_cell = f"**{sum_unpaid}**" + (f" ({sum_partial_unpaid}-Partial Payment not received)" if sum_partial_unpaid else "")
        out.append(
            f"| **Total (AM team)** | **{sum_total}** | {total_paid_cell} | {total_unpaid_cell} | "
            f"**{_usd(sum_booked)}** | **{_usd(sum_paid_usd)}** | **{_usd(sum_pending_usd)}** |"
        )
        out.append("")
        out.append(
            "_Note: each deal is counted once. Artnet is shown on the paid side as Partial Paid (paid portion in $ Received), "
            "with the unpaid remainder rolled into $ Pending._"
        )
        return "\n".join(out)

    def _am_rep_level_answer() -> str:
        """Rep-level pending vs received money, broken down per owner."""
        if not (paid_deals_block or unpaid_overrides):
            return f"No rep-level deal data recorded for {period_lbl} AM."
        owners: dict[str, dict] = {}
        for d in paid_deals_block:
            o = owners.setdefault(d.get("deal_owner") or "—", {"received": 0.0, "pending": 0.0, "paid_deal_count": 0, "unpaid_deal_count": 0})
            o["received"] += float(d.get("amount_usd") or 0)
            o["paid_deal_count"] += 1
        for d in unpaid_overrides:
            o = owners.setdefault(d.get("deal_owner") or "—", {"received": 0.0, "pending": 0.0, "paid_deal_count": 0, "unpaid_deal_count": 0})
            o["pending"] += float(d.get("amount_usd") or 0)
            if not d.get("is_partial_remainder"):
                o["unpaid_deal_count"] += 1
        out: list[str] = [
            f"### {period_lbl} — AM rep-level: received vs pending",
            "",
            "| Rep | $ Payment received | $ Pending (not received) | # Paid deals | # Not paid deals |",
            "| --- | ---: | ---: | ---: | ---: |",
        ]
        total_received = total_pending = 0.0
        total_paid_n = total_unpaid_n = 0
        for owner, o in sorted(owners.items(), key=lambda kv: -kv[1]["received"]):
            out.append(
                f"| **{_esc_md(owner)}** | {_usd(o['received'])} | {_usd(o['pending'])} | "
                f"{o['paid_deal_count']} | {o['unpaid_deal_count']} |"
            )
            total_received += o["received"]
            total_pending += o["pending"]
            total_paid_n += o["paid_deal_count"]
            total_unpaid_n += o["unpaid_deal_count"]
        out.append(
            f"| **Total (AM team)** | **{_usd(total_received)}** | **{_usd(total_pending)}** | "
            f"**{total_paid_n}** | **{total_unpaid_n}** |"
        )
        out.append("")
        out.append(
            "_Note: $ Pending includes any Artnet partial-remainder amount, even though that's the same deal counted on the paid side._"
        )
        return "\n".join(out)

    def _am_unpaid_deals_answer() -> str:
        """List Not Paid AM deals with the partial-payment detail for Artnet."""
        if not unpaid_overrides:
            return f"No Not-Paid AM deals recorded for {period_lbl}."
        out: list[str] = [
            f"### {period_lbl} — AM deals with payment not received",
            "",
            "These deals are excluded from the Payment Received commission base until payment arrives:",
            "",
            "| Deal | Owner | Close date | Status | Amount |",
            "| --- | --- | --- | --- | ---: |",
        ]
        unpaid_total = 0.0
        for d in unpaid_overrides:
            dn = _esc_md((d.get("deal_name") or "").strip())
            owner = _esc_md((d.get("deal_owner") or "").strip())
            cd = _esc_md((d.get("close_date") or "").strip())
            status = _esc_md((d.get("payment_status") or "Not Paid").strip())
            amt = float(d.get("amount_usd") or 0)
            unpaid_total += amt
            out.append(f"| **{dn}** | {owner} | {cd} | {status} | {_usd(amt)} |")
        out.append("")
        out.append(f"**Total excluded value: {_usd(unpaid_total)}**")
        # If we have Artnet partial-paid, show the partial details
        artnet_paid = next((d for d in paid_deals_block if d.get("is_partial")), None)
        artnet_remainder = next((d for d in unpaid_overrides if d.get("is_partial_remainder")), None)
        if artnet_paid and artnet_remainder:
            total_val = float(artnet_paid.get("total_deal_value_usd") or 0) or (
                float(artnet_paid.get("amount_usd") or 0) + float(artnet_remainder.get("amount_usd") or 0)
            )
            out.append("")
            out.append(f"**Artnet partial-payment detail**")
            out.append(f"- Total deal value: {_usd(total_val)}")
            out.append(f"- Paid portion (in Payment Received): {_usd(artnet_paid.get('amount_usd') or 0)}")
            out.append(f"- Pending portion (unpaid remainder): {_usd(artnet_remainder.get('amount_usd') or 0)}")
        return "\n".join(out)

    def _am_joy_manager_stepbystep_answer() -> str:
        """Detailed step-by-step walkthrough of Joy's AM manager commission.

        Includes performance, base commission, Krish Services Group pending commission,
        and the Wipro Q4 clawback. Always shows the final = −$78 with carry-forward note.
        """
        mgr = pinned.get("manager_incentive") or {}
        team_target = float(mgr.get("team_target_usd") or pinned.get("team_target_usd") or 0)
        team_ach = float(mgr.get("team_achievement_usd") or pinned.get("team_achievement_usd") or 0)
        team_ach_pct = float(mgr.get("team_achievement_pct") or 0) or (
            (team_ach / team_target * 100.0) if team_target else 0.0
        )
        team_paid = float(mgr.get("team_payment_received_usd") or 0)
        rate_pct = float(mgr.get("commission_rate_pct") or 1.5)
        manager_q1 = float(mgr.get("manager_q1_amount_usd") or 0) or round(team_paid * rate_pct / 100.0, 2)
        pending_total = float(mgr.get("total_pending_commission_usd") or 0)
        clawback_total = float(mgr.get("total_clawback_usd") or 0)
        q1_gross = float(mgr.get("q1_gross_commission_usd") or 0) or (manager_q1 + pending_total)
        final = float(mgr.get("final_payout_usd"))  # locked at −78
        carry = float(mgr.get("carry_forward_usd") or 0)
        mgr_name = (mgr.get("manager_name") or "Joy").strip()
        pending_items = mgr.get("pending_commissions") or []
        clawback_items = mgr.get("clawbacks") or []
        elig_min = int((pinned.get("commission") or {}).get("eligibility_min_pct") or 50)

        out: list[str] = [
            f"## {period_lbl} — {mgr_name}'s Manager Commission, Step by Step",
            "",
            f"{mgr_name} is **both a rep AND the AM manager** for the Account Management team. Her entire Q1 commission flows through the "
            "`manager_incentive` block in the pinned JSON (not as a regular rep entry under `commission.reps`). Below is exactly how her "
            f"**{_usd(final)}** Q1 result is built.",
            "",
            "---",
            "",
            "### Step 1 — Pull the team's Q1 performance",
            "",
            f"- **Team target**: {_usd(team_target)}",
            f"- **Team achievement (booked)**: {_usd(team_ach)} → **{team_ach_pct:.0f}% of quota**",
            f"- **Team payment received**: {_usd(team_paid)} (after excluding 4 Not-Paid deals + Artnet partial remainder)",
            "",
            f"Eligibility floor for manager base commission is **{elig_min}%** of team target. At **{team_ach_pct:.0f}%**, Joy IS eligible for the base.",
            "",
            "---",
            "",
            "### Step 2 — Calculate Q1 base manager commission",
            "",
            f"AM manager incentive policy: **{rate_pct:g}% of team payment received**.",
            "",
            "```",
            f"Base = Team Payment Received × {rate_pct:g}%",
            f"     = {team_paid:,.2f} × {rate_pct:g}%",
            f"     = {manager_q1:,.2f}",
            "```",
            "",
            f"**Q1 base commission: {_usd(manager_q1)}**",
            "",
            "---",
            "",
        ]
        # Step 3 — Q4 pending (Krish Services Group)
        if pending_items:
            out.append("### Step 3 — Add Q4 pending commissions (paid in Q1 at the Q4 rate)")
            out.append("")
            out.append(
                "These are deals that closed in Q4 FY2025 but whose payment was received in Q1. Per policy they're paid at the "
                "**Q4 rate (1%)**, not the current Q1 rate."
            )
            out.append("")
            out.append("| Deal | Period | Deal amount | Q4 rate | Commission |")
            out.append("| --- | --- | ---: | ---: | ---: |")
            for p in pending_items:
                dn = _esc_md((p.get("deal_name") or "").strip())
                per = _esc_md((p.get("period") or "").strip())
                deal_amt = float(p.get("deal_amount_usd") or 0)
                rate = float(p.get("rate_pct") or 0)
                comm = float(p.get("commission_usd") or 0)
                out.append(f"| **{dn}** | {per} | {_usd(deal_amt)} | {rate:g}% | **+{_usd(comm)}** |")
            out.append("")
            out.append(f"**Subtotal pending: +{_usd(pending_total)}**")
            out.append("")
            out.append(f"Q1 gross commission so far: Base {_usd(manager_q1)} + Pending {_usd(pending_total)} = **{_usd(q1_gross)}**.")
            out.append("")
            out.append("---")
            out.append("")
        # Step 4 — Wipro clawback
        if clawback_items:
            out.append("### Step 4 — Apply the Q4 Wipro slab-adjustment clawback")
            out.append("")
            out.append(
                "The Wipro Q4 FY2025 deal was processed at a **4% slab** at the time of payout, but the correct slab is **1%**. "
                "The 3% over-payment is clawed back from Q1's manager commission."
            )
            out.append("")
            out.append("| Refund / adjustment | Period | Deal amount | Rate adjustment | Clawback |")
            out.append("| --- | --- | ---: | ---: | ---: |")
            for cb in clawback_items:
                lbl = _esc_md((cb.get("label") or cb.get("deal_name") or "").strip())
                per = _esc_md((cb.get("period") or "").strip())
                refund_amt = float(cb.get("deal_amount_usd") or 0)
                rate = float(cb.get("rate_pct") or 0)
                ded = float(cb.get("deduction_usd") or 0)
                out.append(f"| **{lbl}** | {per} | {_usd(refund_amt)} | {rate:g}% | **−{_usd(ded)}** |")
            out.append("")
            out.append(f"**Subtotal clawback: −{_usd(clawback_total)}**")
            out.append("")
            out.append("---")
            out.append("")
        # Final
        out.append("### Step 5 — Add it all up")
        out.append("")
        out.append("```")
        out.append(f"Q1 base                  {manager_q1:>12,.2f}")
        if pending_items:
            out.append(f"+ Q4 pending (Krish)     {pending_total:>12,.2f}")
        if clawback_items:
            out.append(f"− Q4 Wipro clawback      {clawback_total:>12,.2f}")
        out.append(f"─────────────────────────────────────")
        out.append(f"Final {period_lbl} payout   {final:>12,.2f}")
        out.append("```")
        out.append("")
        if final < 0:
            out.append(
                f"### ⚠️ Final Q1 FY2026 payout: **{_usd(final)}**"
            )
            out.append("")
            out.append(
                f"**No payout for Q1.** The **{_usd(final)}** balance carries forward to the next compensation cycle "
                f"({_usd(carry)} if locked in `carry_forward_usd`). Joy doesn't owe anything — the deficit just rolls "
                "into the next quarter's starting balance and is absorbed once new manager commission accrues."
            )
        else:
            out.append(f"### ✅ Final Q1 FY2026 payout: **{_usd(final)}** (≈ {_inr(final)})")
        return "\n".join(out)

    def _am_rep_detail_answer(rep_first_name: str) -> str:
        """Detailed per-rep Q1 commission answer.

        Covers performance vs quota, eligibility check, base commission math, the 10%
        deduction policy, any manage-deal incentive (e.g. Vivin's Q4 Vendasta CF), and
        the rep's specific paid + unpaid deal lists. Strictly scoped to one rep — never
        shows other reps' numbers.
        """
        comm_reps = (pinned.get("commission") or {}).get("reps") or []
        rep_top_list = pinned.get("reps") or []
        ded_pct_team = float((pinned.get("commission") or {}).get("team_deduction_pct") or 10)
        elig_floor = int((pinned.get("commission") or {}).get("eligibility_min_pct") or 50)

        # Find rep in commission.reps (preferred — has full math) or in reps[] (top level)
        rep = next(
            (r for r in comm_reps if (r.get("name") or "").lower().startswith(rep_first_name.lower())),
            None,
        )
        rep_top = next(
            (r for r in rep_top_list if (r.get("name") or "").lower().startswith(rep_first_name.lower())),
            None,
        )
        # Collect this rep's deals
        rep_paid = [d for d in paid_deals_block if (d.get("deal_owner") or "").lower().startswith(rep_first_name.lower())]
        rep_unpaid = [d for d in unpaid_overrides if (d.get("deal_owner") or "").lower().startswith(rep_first_name.lower())]

        if not rep and not rep_top and not rep_paid and not rep_unpaid:
            return f"No data found for {rep_first_name.title()} in {period_lbl} AM."

        # ── Special handling for deal-owners NOT on the commissioned roster (e.g. Nikita)
        if not rep and not rep_top and (rep_paid or rep_unpaid):
            owner_full = (rep_paid[0].get("deal_owner") if rep_paid else rep_unpaid[0].get("deal_owner")) or rep_first_name.title()
            paid_sum_r = sum(float(d.get("amount_usd") or 0) for d in rep_paid)
            unpaid_sum_r = sum(float(d.get("amount_usd") or 0) for d in rep_unpaid)
            out = [
                f"### {owner_full} — {period_lbl} (AM team)",
                "",
                f"**{owner_full} is not on the Q1 commissioned AM roster** — she owns a Q1 deal in HubSpot but is not paid AM rep commission "
                f"directly. The deal she owns still contributes to the team Booked Amount / Payment Received calculation that flows into "
                f"Joy's manager incentive.",
                "",
                "**Deals owned**:",
                "",
                "| Deal | Close date | Status | Amount |",
                "| --- | --- | --- | ---: |",
            ]
            for d in (rep_paid + rep_unpaid):
                out.append(
                    f"| **{_esc_md((d.get('deal_name') or '').strip())}** | "
                    f"{_esc_md((d.get('close_date') or '').strip())} | "
                    f"{_esc_md((d.get('payment_status') or '').strip())} | "
                    f"{_usd(float(d.get('amount_usd') or 0))} |"
                )
            out.append("")
            out.append(f"**Payment received**: {_usd(paid_sum_r)} · **Pending**: {_usd(unpaid_sum_r)}")
            out.append("")
            out.append(
                f"**Personal AM commission for {period_lbl}: $0.00** (not on commissioned roster). "
                f"Once payment arrives for any pending deal, the dollars roll into the AM team total."
            )
            return "\n".join(out)

        # ── Standard commissioned rep flow
        rep_name = (rep.get("name") if rep else rep_top.get("name")) or rep_first_name.title()
        quota = float((rep or {}).get("quota_usd") or (rep_top or {}).get("target_usd") or 0)
        booked = float((rep or {}).get("revenue_achieved_usd") or (rep_top or {}).get("achievement_usd") or 0)
        paid_recv = float((rep or {}).get("payment_received_usd") or (rep_top or {}).get("payment_received_usd") or 0)
        ach_pct = (booked / quota * 100.0) if quota else 0.0
        if rep and rep.get("eligible_pct") is not None:
            ach_pct = float(rep.get("eligible_pct") or 0)
        base = float((rep or {}).get("base_commission_usd") or 0)
        ded_pct = float((rep or {}).get("deduction_pct") or 0)
        ded_amt = float((rep or {}).get("deduction_usd") or 0)
        md_total = float((rep or {}).get("manage_deal_usd") or 0)
        md_paid_now = float((rep or {}).get("manage_deal_paid_now_usd") or 0)
        md_pending = max(md_total - md_paid_now, 0)
        payout = float((rep or {}).get("total_payout_usd") or 0)
        note_v = ((rep or {}).get("note") or "").strip()
        eligible = ach_pct >= elig_floor

        out: list[str] = [
            f"## {rep_name} — {period_lbl} commission breakdown",
            "",
            f"This answer covers **{rep_name} only** — for the full team view, ask for the deal-owner breakdown.",
            "",
            "---",
            "",
            "### Step 1 — Performance vs quota",
            "",
            f"- **Individual quota**: {_usd(quota)}",
            f"- **Revenue achieved (booked)**: {_usd(booked)} → **{ach_pct:.0f}% of quota**",
            f"- **Payment received**: {_usd(paid_recv)}",
            "",
            "---",
            "",
            "### Step 2 — Eligibility check",
            "",
            f"AM Q1 eligibility floor: **{elig_floor}% of individual quota**.",
            "",
        ]
        if eligible:
            out.append(
                f"{rep_name} is at **{ach_pct:.0f}%** of quota → meets the {elig_floor}% floor → **eligible** for Q1 commission. ✅"
            )
            out.append("")
            out.append("---")
            out.append("")
            out.append("### Step 3 — Base commission (1% of Payment Received)")
            out.append("")
            out.append("```")
            out.append(f"Base = Payment Received × 1%")
            out.append(f"     = {paid_recv:,.2f} × 1%")
            out.append(f"     = {base:,.2f}")
            out.append("```")
            out.append("")
            if ded_amt > 0:
                out.append("---")
                out.append("")
                out.append(f"### Step 4 — Apply the {int(ded_pct_team)}% deduction")
                out.append("")
                out.append(
                    f"AM policy: reps who didn't close a Q1 manage deal incur a **{int(ded_pct)}% deduction** on their achieved commission. "
                    f"This applies to {rep_name}."
                )
                out.append("")
                out.append("```")
                out.append(f"Deduction = Base × {int(ded_pct)}%")
                out.append(f"          = {base:,.2f} × {int(ded_pct)}%")
                out.append(f"          = {ded_amt:,.2f}")
                out.append("```")
                out.append("")
                out.append(f"After deduction: **{_usd(base - ded_amt)}**.")
                out.append("")
            if md_total > 0:
                out.append("---")
                out.append("")
                out.append("### Step 5 — Q4 manage deal incentive (paid in Q1)")
                out.append("")
                out.append(
                    f"{rep_name} has a **Q4 FY2025 manage deal** whose payment was received in Q1 — Vendasta CloudFuze Manage Deal. "
                    "Per policy, it's paid at the **10% manage-deal rate**."
                )
                out.append("")
                out.append("```")
                out.append(f"Manage deal incentive = Manage Deal Amount × 10%")
                out.append(f"                      = {md_total:,.2f} × 10%")
                out.append(f"                      = {md_paid_now:,.2f}")
                out.append("```")
                out.append("")
                if md_pending > 0:
                    out.append(f"_(Of the {_usd(md_total)} total manage deal incentive earned, only **{_usd(md_paid_now)}** is paid this cycle — the remaining {_usd(md_pending)} releases when payment arrives.)_")
                    out.append("")
            out.append("---")
            out.append("")
            out.append("### Step 6 — Add it all up")
            out.append("")
            out.append("```")
            out.append(f"Base commission           {base:>12,.2f}")
            if ded_amt > 0:
                out.append(f"− {int(ded_pct)}% deduction            {ded_amt:>12,.2f}")
            if md_paid_now > 0:
                out.append(f"+ Manage deal paid now    {md_paid_now:>12,.2f}")
            out.append(f"──────────────────────────────────────")
            out.append(f"Final Q1 payout           {payout:>12,.2f}")
            out.append("```")
            out.append("")
            out.append(f"### ✅ Final {period_lbl} payout: **{_usd(payout)}** (≈ {_inr(payout)})")
        else:
            out.append(
                f"{rep_name} is at **{ach_pct:.0f}%** of quota — **below the {elig_floor}% floor**. **Not eligible** for Q1 base commission. ❌"
            )
            out.append("")
            out.append("---")
            out.append("")
            out.append("### Step 3 — Commission result")
            out.append("")
            out.append(
                f"Because {rep_name} did not meet the {elig_floor}% eligibility floor, the base commission for Q1 is **$0.00**. Her deals "
                f"still **count toward the team's Booked Amount and Payment Received** (which flow into Joy's manager incentive), but they "
                f"don't earn {rep_name} personal commission this quarter."
            )
            out.append("")
            out.append(f"### ❌ Final {period_lbl} payout: **$0.00**")

        # Show the rep's Q1 deal list at the end for full context
        if rep_paid or rep_unpaid:
            out.append("")
            out.append("---")
            out.append("")
            out.append(f"### {rep_name}'s Q1 deals")
            out.append("")
            out.append("| Deal | Close date | Status | Amount |")
            out.append("| --- | --- | --- | ---: |")
            for d in rep_paid:
                out.append(
                    f"| **{_esc_md((d.get('deal_name') or '').strip())}** | "
                    f"{_esc_md((d.get('close_date') or '').strip())} | "
                    f"{_esc_md((d.get('payment_status') or '').strip())} | "
                    f"{_usd(float(d.get('amount_usd') or 0))} |"
                )
            for d in rep_unpaid:
                out.append(
                    f"| **{_esc_md((d.get('deal_name') or '').strip())}** | "
                    f"{_esc_md((d.get('close_date') or '').strip())} | "
                    f"{_esc_md((d.get('payment_status') or '').strip())} | "
                    f"{_usd(float(d.get('amount_usd') or 0))} |"
                )
            rep_paid_sum = sum(float(d.get("amount_usd") or 0) for d in rep_paid)
            rep_unpaid_sum = sum(float(d.get("amount_usd") or 0) for d in rep_unpaid)
            unique_deal_n = len(rep_paid) + sum(1 for d in rep_unpaid if not d.get("is_partial_remainder"))
            out.append("")
            out.append(
                f"**{unique_deal_n} deal{'s' if unique_deal_n != 1 else ''}** · "
                f"**Received**: {_usd(rep_paid_sum)} · **Pending**: {_usd(rep_unpaid_sum)}"
            )

        if note_v:
            out.append("")
            out.append(f"_{_esc_md(note_v)}_")
        return "\n".join(out)

    if is_am_period:
        # AM-0 (highest priority): Joy's manager commission step-by-step.
        # Triggers on strong signals — Wipro clawback name, or Joy + (step/manager/commission).
        if (
            "wipro" in q
            or ("joy" in q and any(t in q for t in (
                "step by step", "step-by-step", "step by-step", "walk through", "walkthrough",
                "manager commission", "manager's commission", "manager incentive",
                "explain", "calculation", "in detail",
                "carry forward", "carry-forward",
            )))
            or "krish services" in q
            or "mars overage" in q
        ):
            return _am_joy_manager_stepbystep_answer()

        # AM-0.5 (also high priority): per-rep individual detail. Routes BEFORE the
        # deal-owner-breakdown handler so questions about ONE specific rep return that
        # rep's commission breakdown instead of the all-owners table. Multi-rep mentions
        # (e.g. "Vivin, Joy, Arundhati and Nikita") still fall through to the breakdown.
        rep_names_in_q = {
            n for n in ("vivin", "arundhati", "nikita")
            if n in q
        }
        # Joy is intentionally excluded here — Joy's questions go to AM-0 (manager step-by-step).
        if len(rep_names_in_q) == 1 and "joy" not in q:
            single = next(iter(rep_names_in_q))
            return _am_rep_detail_answer(single)
        # AM-A: deal counts
        if any(
            t in q for t in (
                "how many deals", "how many deal", "deal count", "deal counts",
                "number of deals", "number of deal", "total deals closed",
                "total deal closed", "total deals", "total deal",
                "deals closed", "deal closed",
            )
        ):
            return _am_deal_counts_answer()
        # AM-B: deal owner breakdown (multi-rep / all-owners questions only —
        # single-rep names are handled by AM-0.5 above so they get the per-rep detail).
        if any(
            t in q for t in (
                "deal owner", "by owner", "per owner", "owner breakdown",
                "all owners", "every owner", "each owner",
                "by rep", "per rep", "owner-level",
                "break down deals", "breakdown deals", "deals by",
            )
        ) or len(rep_names_in_q) > 1:
            return _am_deal_owner_answer()
        # AM-C: rep-level received vs pending
        if any(
            t in q for t in (
                "rep level", "rep-level", "received vs pending", "pending vs received",
                "how much pending", "how much received", "how much we receive",
                "received and pending", "at rep level",
            )
        ):
            return _am_rep_level_answer()
        # AM-D: list Not Paid / partial / specific AM deals
        if any(
            t in q for t in (
                "not paid", "not-paid", "unpaid", "partially paid", "partial payment",
                "artnet", "estée lauder", "estee lauder", "church & dwight", "church and dwight",
                "allied gold", "nerdstogo", "nerds to go", "carolina contracting",
                "which am deals", "am unpaid deals",
            )
        ):
            return _am_unpaid_deals_answer()

    # Manager-specific question (highest priority — handles "Joy", "Chitradip" by name)
    if mgr_name_low and mgr_name_low in q:
        return _format_manager_answer()

    # By rep name token
    for rs in rep_status_list:
        nm = (rs.get("name") or "").strip()
        first = (nm.split() or [""])[0].lower()
        if first and first in q:
            # If this rep is the manager (e.g. Joy is both rep & manager), return the manager view.
            if rs.get("is_manager"):
                return _format_manager_answer()

            payout_row = next((rp for rp in rep_payouts_list if (rp.get("name") or "").lower() == nm.lower()), None)
            # Build a rich, narrative explanation that covers:
            # (Note: when inr_rate=0 (e.g. ENT page), _usd_with_inr falls back to USD-only.)
            #  - eligibility (incl. policy floor)
            #  - base commission math
            #  - manage deal breakdown (incl. partial-paid scenarios — e.g. Lawrence's 2 deals / 1 paid)
            #  - adjustments / prior-period items
            #  - cross-period context (e.g. left_org later)
            quota_v = float((payout_row or {}).get("quota_usd") or rs.get("quota_usd") or 0)
            ach_v = float((payout_row or {}).get("revenue_achieved_usd") or rs.get("achievement_usd") or 0)
            paid_v = float((payout_row or {}).get("payment_received_usd") or rs.get("payment_received_usd") or 0)
            elig_pct_v = float((payout_row or {}).get("achievement_pct") or rs.get("achievement_pct") or 0)
            base_v = float((payout_row or {}).get("base_commission_usd") or 0)
            ded_v = float((payout_row or {}).get("deduction_usd") or 0)
            md_v = float((payout_row or {}).get("manage_deal_usd") or 0)
            md_paid_now_v = float((payout_row or {}).get("manage_deal_paid_now_usd") or 0)
            md_pending_v = max(md_v - md_paid_now_v, 0)
            payout_usd = float((payout_row or {}).get("total_payout_usd") or 0)
            note_v = ((payout_row or {}).get("note") or "").strip()
            adjustments = (payout_row or {}).get("adjustments_details") or []

            elig_floor = float(team_summary.get("eligibility_floor_pct") or 50)
            eligible_v = elig_pct_v >= elig_floor

            lines: list[str] = [
                f"### Why did {nm} receive {_usd_with_inr(payout_usd)} in {period_lbl}?",
                "",
                f"**Performance**",
                f"- Quota: {_usd(quota_v)}",
                f"- Revenue achieved: {_usd(ach_v)} ({elig_pct_v:.0f}%)",
                f"- Payment received: {_usd(paid_v)}",
                "",
                f"**Regular commission**",
            ]
            is_monthly = bool(period.get("month"))
            strict_floor_note = (
                f" The April monthly plan applies a **strict {int(elig_floor)}% floor** — "
                f"anything below {int(elig_floor)}% (even 59.9%) earns 0% commission."
                if is_monthly and int(elig_floor) == 60
                else ""
            )
            if eligible_v:
                lines.append(
                    f"- {nm} is at **{elig_pct_v:.0f}%** of quota, which meets the **{int(elig_floor)}% eligibility floor**. "
                    f"Base commission: **{_usd(base_v)}**."
                )
            else:
                lines.append(
                    f"- Achievement is **{elig_pct_v:.0f}%** of quota, **below the {int(elig_floor)}% eligibility floor**. "
                    f"{nm} is therefore **not eligible** for regular commission this period. Base commission: **{_usd(0)}**."
                    + strict_floor_note
                )
            if ded_v > 0:
                lines.append(f"- 10% deduction applied: **−{_usd(ded_v)}**.")

            # Manage deal breakdown (handles the Lawrence 2-deals / 1-paid case)
            if md_v > 0 or md_paid_now_v > 0:
                lines.append("")
                lines.append("**Manage deals**")
                if md_v > 0 and md_paid_now_v > 0 and md_v != md_paid_now_v:
                    deals_count = int(round(md_v / 500)) if md_v else 0
                    paid_count = int(round(md_paid_now_v / 500)) if md_paid_now_v else 0
                    lines.append(
                        f"- {nm} closed **{deals_count} manage deal(s)** at $500 each → **{_usd(md_v)} earned total**."
                    )
                    lines.append(
                        f"- However, at the time {period_lbl} commissions were calculated, payment had been received "
                        f"for only **{paid_count}** of those deals. So **{_usd(md_paid_now_v)}** was paid this cycle."
                    )
                    lines.append(
                        f"- The remaining **{_usd(md_pending_v)}** is **pending next cycle** (payment for the other deal not yet received)."
                    )
                else:
                    lines.append(f"- Manage deal incentive: **{_usd(md_v)}**.")
                lines.append(
                    "- Note: Manage Deal Incentive is paid regardless of overall quota achievement — "
                    "it has its own eligibility separate from the regular commission floor."
                )

            # Adjustments (prior-period rep-level items, e.g. Yieldstreet)
            if adjustments:
                lines.append("")
                lines.append("**Prior-period adjustments**")
                for a in adjustments:
                    deal = _esc_md((a.get("deal_name") or "").strip() or "(unnamed deal)")
                    per = _esc_md((a.get("period") or "").strip())
                    amt = float(a.get("commission_usd") or 0)
                    sign = "+" if amt >= 0 else "−"
                    deal_amt = float(a.get("paid_amount_usd") or a.get("deal_amount_usd") or 0)
                    rate = float(a.get("rate_pct") or 0)
                    math_part = f" *({deal_amt:,.2f} × {rate:g}%)*" if (deal_amt and rate) else ""
                    lines.append(
                        f"- **{deal}** "
                        + (f"({per}) " if per else "")
                        + f"— {sign}{_usd(abs(amt))}{math_part}"
                    )

            # If this rep IS the manager (e.g. Anthony for ENT), also surface the
            # manager-level pending commissions and clawbacks since they belong to them.
            rep_is_manager_role = (nm.lower() == (mgr_summary.get("manager_name") or "").lower()) and (mgr_summary.get("manager_name") or "")
            mgr_pending_items = (pinned.get("manager_incentive") or {}).get("pending_commissions") or []
            mgr_clawbacks_items = (pinned.get("manager_incentive") or {}).get("clawbacks") or []
            if rep_is_manager_role and (mgr_pending_items or mgr_clawbacks_items):
                if mgr_pending_items:
                    lines.append("")
                    lines.append("**Q4 pending commissions (paid in this period)**")
                    for p in mgr_pending_items:
                        dn = _esc_md((p.get("deal_name") or "").strip())
                        per = _esc_md((p.get("period") or "").strip())
                        deal_amt = float(p.get("deal_amount_usd") or 0)
                        rate = float(p.get("rate_pct") or 0)
                        comm = float(p.get("commission_usd") or 0)
                        lines.append(
                            f"- **{dn}** ({per}) — {_usd(deal_amt)} × {rate:g}% = **+{_usd(comm)}**"
                        )
                    lines.append(f"  **Subtotal: +{_usd(float(mgr_summary.get('pending_total_usd') or 0))}**")
                if mgr_clawbacks_items:
                    lines.append("")
                    lines.append("**Refunds / clawbacks**")
                    for cb in mgr_clawbacks_items:
                        lbl = _esc_md((cb.get("label") or "").strip())
                        per = _esc_md((cb.get("period") or "").strip())
                        deal_amt = float(cb.get("deal_amount_usd") or 0)
                        rate = float(cb.get("rate_pct") or 0)
                        ded = float(cb.get("deduction_usd") or 0)
                        lines.append(
                            f"- **{lbl}** ({per}) — {_usd(deal_amt)} × {rate:g}% = **−{_usd(ded)}**"
                        )
                    lines.append(f"  **Subtotal: −{_usd(float(mgr_summary.get('clawback_total_usd') or 0))}**")

            # Final result
            lines.append("")
            lines.append(f"**Final payout for {period_lbl}: {_usd_with_inr(payout_usd)}**")

            # JSON note (if present) often has the authoritative narrative
            if note_v:
                lines.append("")
                lines.append(f"_{_esc_md(note_v)}_")

            # Cross-period context (e.g. left org in April)
            cross_note = _cross_period_note_for_rep(first)
            if cross_note:
                lines.append("")
                lines.append(cross_note)

            return "\n".join(lines)

    # Aggregate questions
    team_summary = summary.get("team") or {}
    mgr_is_rep = bool(summary.get("manager_is_also_a_rep"))
    if any(t in q for t in ("total payout", "total commission", "grand total", "team payout", "team commission", "all payouts")):
        if mgr_is_rep:
            # When the only "manager" is also a rep (e.g. ENT / Anthony), the rep payout
            # already includes the manager-level pending commissions and clawbacks. Showing
            # a separate "Manager final payout" line would double-count and confuse readers.
            return (
                f"**{period_lbl} — total payout**\n"
                f"- **Total payout: {_usd_with_inr(summary['grand_total']['payout_usd'])}**"
            )
        return (
            f"**{period_lbl} — total payout**\n"
            f"- Rep payouts (sum): {_usd(summary['rep_payouts_totals']['sum_total_payout_usd'])}\n"
            f"- Manager final payout: {_usd(mgr_summary.get('final_payout_usd') or 0)}\n"
            f"- **Grand total: {_usd_with_inr(summary['grand_total']['payout_usd'])}**"
        )

    if any(t in q for t in ("team achievement", "team target", "team performance", "team total", "achievement", "achieved")):
        return (
            f"**{period_lbl} — team performance**\n"
            f"- Target: {_usd(team_summary.get('target_usd') or 0)}\n"
            f"- Achievement (effective, with exceptions): {_usd(team_summary.get('achievement_effective_usd') or 0)} ({team_summary.get('achievement_pct'):.1f}%)\n"
            f"- HubSpot raw: {_usd(team_summary.get('achievement_raw_usd') or 0)} + Exceptions: {_usd(team_summary.get('total_exception_usd') or 0)}\n"
            f"- Payment received: {_usd(team_summary.get('payment_received_effective_usd') or 0)}"
        )

    # Comprehensive: clawbacks, adjustments, exceptions, pending commissions, prior-period items.
    # Users often use these terms interchangeably — surface everything when any of them is asked.
    prior_keywords = (
        "clawback", "claw back",
        "exception", "exceptions",
        "pending", "carry over", "carry-over", "carry forward", "carry-forward",
        "adjustment", "adjustments",
        "prior period", "prior-period",
        "deduction reason", "why deducted",
        "additional payout",
        # specific deal names users have referenced
        "willow", "wealth", "wipro", "noventiq", "krish", "vendasta", "yieldstreet",
        "soules", "legal soft", "washington post",
    )
    if any(t in q for t in prior_keywords):
        all_items = summary.get("all_prior_period_items") or []
        if not all_items:
            return (
                f"No clawbacks, adjustments, exceptions, or pending commissions were applied in "
                f"{period_lbl} for this team."
            )

        # Group by kind for a clean breakdown.
        grouped: dict[str, list[dict]] = {}
        for it in all_items:
            grouped.setdefault(it.get("kind", "Other"), []).append(it)

        order = [
            "Manager clawback",
            "Manager pending commission",
            "Rep adjustment",
            "Team exception",
        ]
        kind_labels = {
            "Manager clawback": "Manager clawbacks (deductions from prior-period adjustments)",
            "Manager pending commission": "Manager pending commissions (paid in this period from prior-period deals)",
            "Rep adjustment": "Rep-level adjustments (e.g. prior-quarter deal commissions credited this month)",
            "Team exception": "Team exceptions (e.g. server splits / SOW carve-outs added to team total)",
        }

        out: list[str] = [f"**{period_lbl} — prior-period adjustments overview**", ""]
        net_total = 0.0
        for kind in order:
            items = grouped.get(kind) or []
            if not items:
                continue
            out.append(f"### {kind_labels.get(kind, kind)}")
            sub_total = 0.0
            for it in items:
                amt = float(it.get("amount_usd") or 0)
                sub_total += amt
                deal = _esc_md((it.get("deal_name") or "").strip() or "(unnamed deal)")
                person = _esc_md((it.get("rep_or_manager") or "").strip())
                per = _esc_md((it.get("period") or "").strip())
                deal_amt = float(it.get("deal_amount_usd") or 0)
                rate = float(it.get("rate_pct") or 0)
                sign = "+" if amt >= 0 else "−"
                math_str = ""
                if deal_amt and rate:
                    math_str = f" *(\\${deal_amt:,.2f} × {rate:g}%)*"
                note = _esc_md((it.get("note") or "").strip())
                out.append(
                    f"- **{deal}** "
                    + (f"→ {person} " if person else "")
                    + (f"({per}) " if per else "")
                    + f"— {sign}{_usd(abs(amt))}{math_str}"
                    + (f"\n  _{note}_" if note else "")
                )
            net_total += sub_total
            out.append(f"  **Subtotal: {'+' if sub_total >= 0 else '−'}{_usd(abs(sub_total))}**")
            out.append("")

        out.append(f"**Net impact on this period's payouts: {'+' if net_total >= 0 else '−'}{_usd(abs(net_total))}**")
        return "\n".join(out)

    # INR conversion questions
    if any(t in q for t in ("inr", "rupee", "rupees", "₹", "indian")):
        if not inr_rate:
            return "No INR exchange rate set in this pinned period."
        if mgr_is_rep:
            return (
                f"**{period_lbl} — payouts in INR @ ₹{inr_rate:g}/USD**\n"
                f"- **Total payout: {_inr(summary['grand_total']['payout_usd'])}**"
            )
        return (
            f"**{period_lbl} — payouts in INR @ ₹{inr_rate:g}/USD**\n"
            f"- Rep payouts (sum): {_inr(summary['rep_payouts_totals']['sum_total_payout_usd'])}\n"
            f"- Manager final payout: {_inr(mgr_summary.get('final_payout_usd') or 0)}\n"
            f"- **Grand total: {_inr(summary['grand_total']['payout_usd'])}**"
        )

    # Eligibility questions
    if any(t in q for t in ("eligible", "eligibility", "qualif")):
        not_elig = [rs for rs in rep_status_list if "not eligible" in (rs.get("status") or "").lower() or "left org" in (rs.get("status") or "").lower()]
        eligible = [rs for rs in rep_status_list if rs not in not_elig and not rs.get("is_manager")]
        is_monthly = bool(period.get("month"))
        elig_floor = int(team_summary.get("eligibility_floor_pct") or 50)
        lines = []
        # Lead with the policy rule so the answer is unambiguous.
        if is_monthly and elig_floor == 60:
            lines.append(
                f"**Eligibility rule for {period_lbl}:** Sales reps must achieve a minimum of **{elig_floor}%** of their individual quota to qualify for commission. "
                f"Anyone below {elig_floor}% — even **59.9%** — earns **0%** commission."
            )
        else:
            lines.append(
                f"**Eligibility rule for {period_lbl}:** Sales reps must achieve a minimum of **{elig_floor}%** of their individual quota to qualify for commission."
            )
        lines.append("")
        if eligible:
            lines.append("**Eligible reps:**")
            for rs in eligible:
                lines.append(f"- {rs.get('name')} — {rs.get('achievement_pct'):.0f}% achievement")
        if not_elig:
            lines.append("\n**Not eligible:**")
            for rs in not_elig:
                lines.append(f"- {rs.get('name')} — {rs.get('status')}")
        lines.append(f"\nManager ({mgr_summary.get('manager_name')}): "
                     + ("✅ eligible for base" if mgr_summary.get('is_eligible_for_base') else "❌ not eligible for base"))
        return "\n".join(lines)

    # Each-rep summary
    if any(t in q for t in ("each rep", "all reps", "per rep", "everyone", "rep breakdown", "list reps")):
        if not rep_payouts_list:
            return "No rep payouts computed yet for this period."
        lines = []
        for rp in rep_payouts_list:
            nm = rp.get("name") or "—"
            payout = float(rp.get("total_payout_usd") or 0)
            lines.append(f"- **{nm}** — {_usd(payout)} ({_inr(payout)})")
        lines.append(f"\n**Sum:** {_usd(summary['rep_payouts_totals']['sum_total_payout_usd'])}")
        return f"**{period_lbl} — rep payouts**\n" + "\n".join(lines)

    # Default: brief overview
    if mgr_is_rep:
        # ENT (Anthony) case: only one person, who is both rep and manager. The rep payout
        # already includes pending/clawback adjustments, so don't show a separate manager line.
        return (
            f"**{period_lbl} — quick overview**\n"
            f"- Team target: {_usd(team_summary.get('target_usd') or 0)}\n"
            f"- Team achievement: {_usd(team_summary.get('achievement_effective_usd') or 0)} ({team_summary.get('achievement_pct'):.1f}%)\n"
            f"- **Total payout: {_usd(summary['grand_total']['payout_usd'])}**\n\n"
            "Try asking: \"total payout\", \"why is the commission this amount\", "
            "\"pending commissions\", \"clawbacks\", or \"unpaid deals\"."
        )
    return (
        f"**{period_lbl} — quick overview**\n"
        f"- Team target: {_usd(team_summary.get('target_usd') or 0)}\n"
        f"- Team achievement: {_usd(team_summary.get('achievement_effective_usd') or 0)} ({team_summary.get('achievement_pct'):.1f}%)\n"
        f"- Rep payouts (sum): {_usd(summary['rep_payouts_totals']['sum_total_payout_usd'])}\n"
        f"- Manager final payout: {_usd(mgr_summary.get('final_payout_usd') or 0)}\n"
        f"- Grand total: {_usd(summary['grand_total']['payout_usd'])}\n\n"
        "Try asking: \"total payout\", \"each rep's commission\", \"why is X not eligible\", \"clawbacks\", "
        "\"pending commissions\", \"exceptions\", or \"convert to INR\"."
    )


def _ask_compensation_assistant(question: str, history: list[dict] | None = None) -> str:
    """Answer a question using Anthropic Claude with the full compensation-tool context.

    Falls back to a smart rule-based bot if no API key is configured (auto-detects which
    pinned period to load from the question text).
    """
    import json as _json

    api_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    if api_key and not api_key.startswith("sk-ant-"):
        api_key = ""

    anthropic_mod = None
    if api_key:
        try:
            import anthropic as anthropic_mod  # type: ignore
        except ImportError:
            anthropic_mod = None

    if anthropic_mod is None:
        # Smart fallback: route to the right pinned file based on question keywords.
        pinned = _detect_pinned_period_for_question(question)
        if pinned is not None:
            return _answer_pinned_period(pinned, question)
        # Try SMB Q1 as ultimate fallback for the rule-based bot.
        return _answer_smb_question(_load_pinned_smb_quarter(2026, 1) or {}, question)

    anthropic = anthropic_mod  # alias so existing code below keeps working

    model = (os.environ.get("ANTHROPIC_MODEL") or "claude-haiku-4-5-20251001").strip()
    context = _gather_compensation_chat_context()
    context_json = _json.dumps(context, default=str, ensure_ascii=False)
    # Hard cap on context size to keep token cost predictable.
    if len(context_json) > 60000:
        context_json = context_json[:60000] + "\n... (context truncated)"

    system_prompt = (
        "You are the in-app assistant for the CloudFuze Sales Compensation Tool. "
        "Answer questions about sales rep performance, individual & manager commissions, "
        "Q1 FY2026 SMB pinned numbers, AM and Enterprise team aggregates, the user roster, "
        "outbound meeting incentives, HubSpot deals/goals fetched this session, and the policy "
        "tiers (SMB Group A/B, AM, Enterprise, Manager pool, Outbound). "
        "Use ONLY the JSON context below. If something isn't in the context, say so plainly. "
        "Always show USD with a $ prefix; show INR with ₹ when it's helpful (use exchange_rate_inr_per_usd from the pinned commission block). "
        "Use markdown for clarity. Keep answers concise but complete; show the calculation when it helps. "
        "Today's date: 2026-05-05. Fiscal year = calendar year."
    )

    msgs: list[dict] = []
    for h in (history or [])[-6:]:
        if h.get("role") in ("user", "assistant") and (h.get("content") or "").strip():
            msgs.append({"role": h["role"], "content": h["content"]})
    msgs.append(
        {
            "role": "user",
            "content": (
                f"Question: {question}\n\n"
                f"---\nCompensation tool data (JSON):\n{context_json}"
            ),
        }
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model=model,
            max_tokens=1200,
            system=system_prompt,
            messages=msgs,
        )
        chunks = [c.text for c in (resp.content or []) if getattr(c, "type", "text") == "text"]
        return "\n".join(chunks).strip() or "(empty response)"
    except Exception as e:
        return f"Anthropic API error: {e}"


def _pinned_export_to_excel_bytes(sheets: dict) -> bytes:
    """Render a dict of ``{sheet_name: DataFrame}`` as an .xlsx byte string for download_button.

    Uses openpyxl so it works without the optional xlsxwriter dependency.
    """
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for name, df in sheets.items():
                # Sheet names are limited to 31 chars and may not contain certain characters.
                clean_name = (name or "Sheet")[:31].replace("/", " ").replace("\\", " ")
                df.to_excel(writer, sheet_name=clean_name, index=False)
    except Exception:
        # Fallback to xlsxwriter if openpyxl is missing.
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            for name, df in sheets.items():
                clean_name = (name or "Sheet")[:31].replace("/", " ").replace("\\", " ")
                df.to_excel(writer, sheet_name=clean_name, index=False)
    buf.seek(0)
    return buf.read()


def _gather_assistant_context() -> dict:
    """Collect a compact JSON-able snapshot of all data the Assistant needs to reason about.

    Includes pinned SMB Q1 2026 numbers, live rep_incentives across teams (compact),
    user roster, outbound meetings, policy summary, and HubSpot deals already fetched
    in this session (does NOT trigger live HubSpot calls).
    """
    ctx: dict = {}
    # Pinned SMB Q1 2026
    ctx["pinned_smb_q1_2026"] = _load_pinned_smb_quarter(2026, 1) or {}

    # Database — keep things compact.
    try:
        users = get_all_users_with_teams() or []
    except Exception:
        users = []
    ctx["users"] = [
        {
            "user_id": u.get("user_id"),
            "full_name": u.get("full_name"),
            "email": u.get("email"),
            "team_name": u.get("team_name"),
            "role": u.get("role"),
            "compensation_group": u.get("compensation_group"),
            "hubspot_quota_usd": u.get("hubspot_quota_usd"),
            "is_active": u.get("is_active"),
        }
        for u in users
    ]

    try:
        incentives = get_rep_incentives() or []
    except Exception:
        incentives = []
    # Aggregate by (team, period) and per-rep totals to keep payload small.
    team_period_totals: dict = {}
    rep_totals: dict = {}
    for r in incentives:
        team = (r.get("team_name") or "").strip() or "?"
        period = (r.get("calculation_period") or "").strip() or "?"
        key = f"{team}__{period}"
        agg = team_period_totals.setdefault(
            key,
            {"team_name": team, "period": period, "rep_count": 0, "total_revenue_usd": 0.0, "total_paid_amount_usd": 0.0, "incentive_amount_usd": 0.0},
        )
        agg["rep_count"] += 1
        agg["total_revenue_usd"] += float(r.get("total_revenue") or 0)
        agg["total_paid_amount_usd"] += float(r.get("total_paid_amount") or 0)
        agg["incentive_amount_usd"] += float(r.get("incentive_amount") or 0)

        rk = f"{r.get('user_id')}__{period}"
        rt = rep_totals.setdefault(
            rk,
            {
                "user_id": r.get("user_id"),
                "full_name": r.get("full_name"),
                "team_name": team,
                "period": period,
                "quota_usd": float(r.get("quota") or 0),
                "total_revenue_usd": 0.0,
                "total_paid_amount_usd": 0.0,
                "incentive_amount_usd": 0.0,
                "achievement_pct": r.get("achievement_pct"),
                "commission_pct": r.get("commission_pct"),
            },
        )
        rt["total_revenue_usd"] += float(r.get("total_revenue") or 0)
        rt["total_paid_amount_usd"] += float(r.get("total_paid_amount") or 0)
        rt["incentive_amount_usd"] += float(r.get("incentive_amount") or 0)
    ctx["rep_incentives_team_totals"] = list(team_period_totals.values())
    ctx["rep_incentives_per_rep"] = list(rep_totals.values())

    # Outbound meetings
    try:
        outbound = get_all_outbound_meetings() or []
    except Exception:
        outbound = []
    ctx["outbound_meetings"] = [
        {
            "outbound_id": o.get("outbound_id"),
            "rep_name": o.get("rep_name"),
            "rep_email": o.get("rep_email"),
            "region": o.get("region"),
            "meeting_date": str(o.get("meeting_date")) if o.get("meeting_date") else None,
            "incentive_amount_usd": float(o.get("incentive_amount") or 0),
            "notes": o.get("notes"),
            "created_at": str(o.get("created_at")) if o.get("created_at") else None,
        }
        for o in outbound
    ]

    # Policy summary
    try:
        from commission_policy import (
            ACCOUNT_MANAGEMENT_TEAM_NAME,
            AM_QUOTA_ACHIEVEMENT_ENABLED,
            ENTERPRISE_TEAM_NAME,
            OUTBOUND_ELIGIBLE_REGIONS,
            OUTBOUND_MEETING_PAYOUT_NOTE,
            OUTBOUND_MEETING_PAYOUT_ROWS,
            OUTBOUND_POLICY_LABEL,
            SMB_ACHIEVEMENT_TIERS,
            SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
            SMB_TEAM_NAME,
            TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT,
            TEAM_QUARTERLY_TARGETS_USD,
        )
        ctx["policy"] = {
            "team_names": {
                "smb": SMB_TEAM_NAME,
                "account_management": ACCOUNT_MANAGEMENT_TEAM_NAME,
                "enterprise": ENTERPRISE_TEAM_NAME,
            },
            "quarterly_team_targets_usd": dict(TEAM_QUARTERLY_TARGETS_USD or {}),
            "smb_achievement_tiers": list(SMB_ACHIEVEMENT_TIERS or []),
            "smb_min_achievement_for_commission_pct": float(SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT or 0),
            "am_quota_achievement_enabled": bool(AM_QUOTA_ACHIEVEMENT_ENABLED),
            "manager_team_pool_thresholds_pct": list(TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT or []),
            "outbound_policy_label": OUTBOUND_POLICY_LABEL,
            "outbound_eligible_regions": [{"code": c, "label": l} for c, l in (OUTBOUND_ELIGIBLE_REGIONS or [])],
            "outbound_meeting_payout_tiers": list(OUTBOUND_MEETING_PAYOUT_ROWS or []),
            "outbound_meeting_payout_note": OUTBOUND_MEETING_PAYOUT_NOTE,
        }
    except Exception:
        ctx["policy"] = {}

    # HubSpot — only what's already in session_state from earlier fetches.
    hs_session: dict = {}
    for ctx_key in ("smb", "am", "ent"):
        df = st.session_state.get(f"hubspot_last_fetch_deals_df_{ctx_key}")
        if df is not None and not df.empty:
            try:
                hs_session[ctx_key] = {
                    "deal_count": int(df.shape[0]),
                    "total_amount_usd": float(df["Amount"].astype(float).sum()) if "Amount" in df.columns else 0.0,
                    "summary": st.session_state.get(f"hubspot_last_fetch_summary_{ctx_key}", ""),
                    "deals_preview": df.head(40).to_dict("records"),
                }
            except Exception:
                pass
    if hs_session:
        ctx["hubspot_fetched_in_session"] = hs_session

    return ctx


_ASSISTANT_SYSTEM_PROMPT = """You are the in-app Q&A assistant for the CloudFuze Sales Compensation Tool.
You answer questions about sales rep performance, commissions, manager incentives, outbound meetings,
HubSpot deals, and the tool's commission policy. You will be given a JSON context block that contains
the entire current state of the tool's data.

Rules:
- Use ONLY the data in the context. If something is not in the context, say so plainly — do not invent numbers.
- For SMB Q1 2026, the source of truth is `pinned_smb_q1_2026` (NOT `rep_incentives_*`). Always prefer pinned values for that quarter.
- Format USD amounts as `$1,234.56` and INR as `₹1,234`. Use the exchange_rate_inr_per_usd from pinned commission block when converting USD→INR (default 86 if absent).
- Use markdown — short sections, bullet lists, bold key numbers.
- Keep answers concise: lead with the direct answer, then 2-6 supporting bullets. No fluff.
- When the user asks about totals, show the math (e.g. "Reps: $X + Manager: $Y = $Z").
- If the user asks about a specific rep by first name, match against full_name in the context.
- Never expose internal IDs (user_id, outbound_id) unless the user explicitly asks.
- If the question is ambiguous, ask one clarifying question instead of guessing.
"""


def _ask_assistant_llm(question: str, history: list[dict], context: dict) -> str:
    """Call Anthropic Claude with the question + recent history + tool context.

    Falls back to a 'no API key' message if ANTHROPIC_API_KEY is unset, and to a
    helpful error if the anthropic SDK is not installed.
    """
    api_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    if not api_key:
        try:
            secret = st.secrets.get("ANTHROPIC_API_KEY") if hasattr(st, "secrets") else None
            if secret:
                api_key = str(secret).strip()
        except Exception:
            pass
    if not api_key:
        return (
            "**ANTHROPIC_API_KEY is not set.**\n\n"
            "Add it to your `.env` (or `.streamlit/secrets.toml`) to enable AI answers. "
            "Get a key at https://console.anthropic.com → API keys.\n\n"
            "_Falling back to suggested-prompt list. Try one of those for quick answers from pinned data._"
        )

    try:
        import anthropic  # type: ignore
    except ImportError:
        return (
            "**The `anthropic` Python package isn't installed.**\n\n"
            "Run: `pip install anthropic` (or `pip install -r requirements.txt`)."
        )

    import json as _json
    model = (os.environ.get("ANTHROPIC_MODEL") or "claude-haiku-4-5-20251001").strip()
    # Trim very large fields to keep token usage low.
    ctx_json = _json.dumps(context, default=str)
    if len(ctx_json) > 60000:
        ctx_json = ctx_json[:60000] + "\n\n[…context truncated to fit token budget…]"

    msgs: list[dict] = []
    for h in history[-10:]:
        if h.get("role") in ("user", "assistant"):
            msgs.append({"role": h["role"], "content": str(h.get("content") or "")})
    msgs.append(
        {
            "role": "user",
            "content": f"Question: {question}\n\nCurrent tool data (JSON):\n```json\n{ctx_json}\n```",
        }
    )

    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model=model,
            max_tokens=1024,
            system=_ASSISTANT_SYSTEM_PROMPT,
            messages=msgs,
        )
        parts = []
        for block in (resp.content or []):
            if getattr(block, "type", "") == "text":
                parts.append(block.text)
        return ("\n".join(parts).strip()) or "_(no response)_"
    except Exception as e:
        return f"**Anthropic API error:** {e}\n\nDouble-check your `ANTHROPIC_API_KEY` and network connectivity."


_ASSISTANT_SUGGESTED_PROMPTS = (
    "What is the total Q1 2026 payout for SMB including the manager?",
    "Break down Chitradip's commission with the clawbacks.",
    "Who are the top performers across SMB, AM, and Enterprise?",
    "Which reps are not eligible for commission this quarter and why?",
    "Why is Lawrence's payout only $500 — show the manage-deal math.",
    "Summarize the SMB tier policy and the eligibility threshold.",
    "How many outbound meetings have been logged this quarter and by whom?",
    "What is the team payment-collection rate this quarter (SMB)?",
    "Show all SMB rep payouts in INR.",
    "Compare each team's revenue vs quarterly target.",
)


def render_assistant_page() -> None:
    """Sidebar 'Assistant' page: full-context Q&A across SMB/AM/ENT/Outbound/Policy/HubSpot."""
    st.markdown("# Assistant")
    st.caption(
        "Ask anything about your compensation data — SMB · AM · ENT · Outbound · Policy · HubSpot deals from this session. "
        "Powered by Claude (Anthropic). Set `ANTHROPIC_API_KEY` in `.env` to enable AI answers."
    )

    # Status row: tells the user whether AI is wired up.
    api_key_present = bool(
        (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
        or (
            (st.secrets.get("ANTHROPIC_API_KEY") if hasattr(st, "secrets") else None)
            if True
            else None
        )
    )
    status_col1, status_col2 = st.columns([3, 1])
    with status_col1:
        if api_key_present:
            st.success("AI mode: Claude (Anthropic) is connected.")
        else:
            st.warning("AI mode is OFF — set `ANTHROPIC_API_KEY` in your `.env` to enable. Without it the bot returns guidance, not answers.")
    with status_col2:
        if st.button("Clear chat", key="assistant_clear", use_container_width=True):
            st.session_state["assistant_chat_history"] = []
            st.rerun()

    st.markdown("**Try one of these:**")
    cols = st.columns(2)
    for i, prompt in enumerate(_ASSISTANT_SUGGESTED_PROMPTS):
        with cols[i % 2]:
            if st.button(prompt, key=f"assistant_sug_{i}", use_container_width=True):
                history = st.session_state.setdefault("assistant_chat_history", [])
                history.append({"role": "user", "content": prompt})
                with st.spinner("Thinking…"):
                    ctx = _gather_assistant_context()
                    answer = _ask_assistant_llm(prompt, history, ctx)
                history.append({"role": "assistant", "content": answer})
                st.rerun()

    st.markdown("---")
    history = st.session_state.setdefault("assistant_chat_history", [])
    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_input := st.chat_input("Ask a question…"):
        history.append({"role": "user", "content": user_input})
        with st.spinner("Thinking…"):
            ctx = _gather_assistant_context()
            answer = _ask_assistant_llm(user_input, history, ctx)
        history.append({"role": "assistant", "content": answer})
        st.rerun()


_SMB_BOT_SUGGESTED_PROMPTS = (
    "What is the total payout for Q1 including manager?",
    "How much did Chitradip earn after clawbacks?",
    "Who is the top performer this quarter?",
    "Who is not eligible for commission?",
    "Why is Lawrence's payout only $500?",
    "What clawbacks were applied to Chitradip?",
    "What is the team achievement vs target?",
    "How much did Vicky earn in INR?",
    "Show me cash collection by rep",
    "Show me each rep's commission",
)


def _answer_smb_question(pinned: dict, question: str) -> str:
    """Rule-based Q&A over the pinned SMB JSON. Returns a Markdown answer."""
    if not pinned:
        return "No pinned SMB data available. Make sure `policy/smb_q1_2026_fixed.json` exists."

    q = (question or "").lower().strip()
    if not q:
        return "Please enter a question."

    year = pinned.get("year")
    quarter = pinned.get("quarter")
    period = f"Q{quarter} {year}"
    reps_top = pinned.get("reps") or []
    comm = pinned.get("commission") or {}
    comm_reps = comm.get("reps") or []
    mgr = pinned.get("manager_incentive") or {}

    try:
        inr_rate = float(comm.get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0
    try:
        ded_pct = float(comm.get("team_deduction_pct") or 0)
    except (TypeError, ValueError):
        ded_pct = 0.0
    elig_min = float(comm.get("eligibility_min_pct") or 50)

    rep_payouts_total = sum(float(r.get("total_payout_usd") or 0) for r in comm_reps)
    base_total = sum(float(r.get("base_commission_usd") or 0) for r in comm_reps)
    manage_earned_total = sum(float(r.get("manage_deal_usd") or 0) for r in comm_reps)
    manage_paid_now_total = sum(float(r.get("manage_deal_paid_now_usd") or r.get("manage_deal_usd") or 0) for r in comm_reps)
    deduction_total = round(base_total * (ded_pct / 100.0), 2)

    mgr_name = (mgr.get("manager_name") or "Manager").strip()
    mgr_base = float(mgr.get("manager_q1_amount_usd") or 0)
    mgr_clawback = float(mgr.get("total_clawback_usd") or 0)
    mgr_final = float(mgr.get("final_payout_usd") or (mgr_base - mgr_clawback))
    mgr_rate = float(mgr.get("commission_rate_pct") or 0)
    mgr_team_paid = float(mgr.get("team_payment_received_usd") or 0)

    team_target = float(pinned.get("team_target_usd") or 0)
    team_ach = float(pinned.get("team_achievement_usd") or 0)
    team_ach_pct = (team_ach / team_target * 100.0) if team_target else 0.0
    team_collection_pct = (mgr_team_paid / team_ach * 100.0) if team_ach else 0.0

    grand_total = rep_payouts_total + mgr_final

    def _usd(x: float) -> str:
        return f"${x:,.2f}"

    def _inr(x: float) -> str:
        return f"₹{x * inr_rate:,.0f}" if inr_rate else "—"

    def _both(x: float) -> str:
        return f"{_usd(x)} ({_inr(x)})" if inr_rate else _usd(x)

    # ---- Intent: total payout (with/without manager) ----
    if any(k in q for k in ("total payout", "total commission", "total we paid", "how much did we pay", "grand total")):
        include_mgr = any(k in q for k in ("manager", "manag", "chit", "including", "with manager", "all"))
        if include_mgr:
            return (
                f"**{period} grand total payout (reps + manager):** {_both(grand_total)}\n\n"
                f"- Sales reps total: {_both(rep_payouts_total)}\n"
                f"- {mgr_name} (after clawbacks): {_both(mgr_final)}"
            )
        return (
            f"**{period} sales rep payouts (excluding manager):** {_both(rep_payouts_total)}\n\n"
            f"_Add the manager: {_both(mgr_final)} → grand total {_both(grand_total)}_"
        )

    # ---- Intent: manager / Chitradip ----
    if any(k in q for k in ("manager", "chit", "chitradip")):
        clawbacks = mgr.get("clawbacks") or []
        cb_lines = "\n".join(
            f"  - {(cb.get('label') or '').strip()} → −${float(cb.get('deduction_usd') or 0):,.0f}"
            for cb in clawbacks
        )
        return (
            f"**{mgr_name}'s {period} commission**\n\n"
            f"- Base commission: {_both(mgr_base)} ({mgr_rate:.2f}% × ${mgr_team_paid:,.0f} team payment received)\n"
            f"- Clawbacks: −${mgr_clawback:,.0f}\n{cb_lines}\n"
            f"- **Final payout:** {_both(mgr_final)}"
        )

    # ---- Intent: top performer ----
    if any(k in q for k in ("top performer", "highest", "best rep", "who performed best", "top rep")):
        if not reps_top:
            return "No rep data available."
        ranked = sorted(
            reps_top,
            key=lambda r: (float(r.get("achievement_usd") or 0) / float(r.get("target_usd") or 1)),
            reverse=True,
        )
        top = ranked[0]
        nm = (top.get("name") or "").strip()
        ach = float(top.get("achievement_usd") or 0)
        tgt = float(top.get("target_usd") or 0)
        pct = (ach / tgt * 100.0) if tgt else 0
        comm_row = next((c for c in comm_reps if (c.get("name") or "").strip() == nm), None)
        payout = float((comm_row or {}).get("total_payout_usd") or 0)
        return (
            f"**Top performer in {period}:** {nm} — **{pct:.0f}%** of quota "
            f"(${ach:,.0f} of ${tgt:,.0f})\n\n"
            f"Commission this cycle: {_both(payout)}"
        )

    # ---- Intent: who is not eligible ----
    if any(k in q for k in ("not eligible", "ineligible", "below 50", "didn't qualify", "not qualifying", "not earning")):
        not_elig = [c for c in comm_reps if float(c.get("eligible_pct") or 0) < elig_min]
        if not not_elig:
            return f"All SMB reps met the {elig_min:.0f}% threshold this quarter."
        lines = "\n".join(
            f"- **{(c.get('name') or '').strip()}** — {float(c.get('eligible_pct') or 0):.0f}% achievement"
            for c in not_elig
        )
        return (
            f"**{len(not_elig)} rep(s) below the {elig_min:.0f}% eligibility threshold in {period}:**\n\n{lines}\n\n"
            f"_Eligibility rule: {(comm.get('eligibility_text') or '').strip()}_"
        )

    # ---- Intent: Lawrence / manage deal ----
    if any(k in q for k in ("lawrence", "larry", "manage deal", "managed deal", "pending")):
        larry = next((c for c in comm_reps if "lawrence" in (c.get("name") or "").lower()), None)
        if not larry:
            return "Lawrence not found in pinned data."
        earned = float(larry.get("manage_deal_usd") or 0)
        paid_now = float(larry.get("manage_deal_paid_now_usd") or 0)
        pending = max(earned - paid_now, 0)
        return (
            f"**Lawrence Lewis — manage deals in {period}**\n\n"
            f"- Closed **2 manage deals** at $500 each → **${earned:,.0f}** total earned\n"
            f"- Paid this cycle: **${paid_now:,.0f}** (1 deal)\n"
            f"- Pending next cycle: **${pending:,.0f}** (payment for the second deal not received at duration of payout)\n"
            f"- **Total payout this cycle:** {_both(float(larry.get('total_payout_usd') or 0))}"
        )

    # ---- Intent: clawbacks ----
    if any(k in q for k in ("clawback", "clawbacks", "deduction reason", "why deducted", "manager deductions")):
        clawbacks = mgr.get("clawbacks") or []
        if not clawbacks:
            return "No clawbacks were applied this quarter."
        lines = "\n".join(
            f"- **{(cb.get('label') or '').strip()}** — ${float(cb.get('deal_amount_usd') or 0):,.2f} × "
            f"{float(cb.get('rate_pct') or 0):g}% = **−${float(cb.get('deduction_usd') or 0):,.0f}** "
            f"_({(cb.get('note') or '').strip()})_"
            for cb in clawbacks
        )
        return (
            f"**{mgr_name} clawbacks for {period}** (total **−${mgr_clawback:,.0f}**):\n\n{lines}\n\n"
            f"Net effect: ${mgr_base:,.0f} (base) − ${mgr_clawback:,.0f} = **{_both(mgr_final)}**"
        )

    # ---- Intent: team achievement / target / collection ----
    if any(k in q for k in ("team achievement", "team target", "team progress", "achievement vs", "vs target", "achieved vs target")):
        return (
            f"**{period} team performance:**\n\n"
            f"- Target: ${team_target:,.0f}\n"
            f"- Achieved (booked): ${team_ach:,.0f} → **{team_ach_pct:.1f}%** of target\n"
            f"- Payment received (collected): ${mgr_team_paid:,.0f} → **{team_collection_pct:.1f}%** of achieved"
        )

    if any(k in q for k in ("collection", "paid vs achieved", "cash collected", "payment received")):
        lines = []
        for r in reps_top:
            nm = (r.get("name") or "").strip()
            ach = float(r.get("achievement_usd") or 0)
            paid = float(r.get("payment_received_usd") or 0)
            ratio = (paid / ach * 100.0) if ach else 0
            lines.append(f"- **{nm}** — ${paid:,.0f} of ${ach:,.0f} ({ratio:.0f}%)")
        return (
            f"**Cash collection by rep — {period}:**\n\n"
            + "\n".join(lines)
            + f"\n\n_Team total: ${mgr_team_paid:,.0f} of ${team_ach:,.0f} ({team_collection_pct:.1f}%)_"
        )

    # ---- Intent: each rep's commission ----
    if any(k in q for k in ("each rep", "all reps", "per rep", "by rep", "every rep", "everyone")):
        lines = []
        for c in comm_reps:
            nm = (c.get("name") or "").strip()
            payout = float(c.get("total_payout_usd") or 0)
            achv = float(c.get("eligible_pct") or 0)
            lines.append(f"- **{nm}** ({achv:.0f}%) — {_both(payout)}")
        return (
            f"**Commission payout per rep — {period}** (sum {_both(rep_payouts_total)}):\n\n"
            + "\n".join(lines)
        )

    # ---- Intent: specific rep by name ----
    rep_match = None
    for c in comm_reps:
        nm = (c.get("name") or "").lower()
        first = nm.split(" ")[0] if nm else ""
        if nm and (nm in q or (first and first in q)):
            rep_match = c
            break
    if rep_match:
        nm = (rep_match.get("name") or "").strip()
        achv = float(rep_match.get("eligible_pct") or 0)
        rev = float(rep_match.get("revenue_achieved_usd") or 0)
        base = float(rep_match.get("base_commission_usd") or 0)
        payout = float(rep_match.get("total_payout_usd") or 0)
        ded = round(base * (ded_pct / 100.0), 2)
        manage = float(rep_match.get("manage_deal_usd") or 0)
        note = (rep_match.get("note") or "").strip()
        return (
            f"**{nm} — {period} commission**\n\n"
            f"- Group: {(rep_match.get('group') or '').strip()}\n"
            f"- Quota: ${float(rep_match.get('quota_usd') or 0):,.0f}\n"
            f"- Revenue achieved: ${rev:,.0f} ({achv:.0f}% of quota)\n"
            f"- Base commission: ${base:,.0f}\n"
            f"- {ded_pct:.0f}% deduction: −${ded:,.0f}\n"
            f"- Manage deal earned: ${manage:,.0f}\n"
            f"- **Total payout:** {_both(payout)}\n"
            + (f"\n_{note}_" if note else "")
        )

    # ---- Intent: INR conversion ----
    if any(k in q for k in ("inr", "rupee", "₹", "indian")):
        return (
            f"**{period} payout in INR @ ₹{inr_rate:g}/USD:**\n\n"
            f"- Sales reps total: ₹{rep_payouts_total * inr_rate:,.0f}\n"
            f"- {mgr_name}: ₹{mgr_final * inr_rate:,.0f}\n"
            f"- **Grand total: ₹{grand_total * inr_rate:,.0f}**"
        )

    # ---- Fallback: list capabilities ----
    sample = "\n".join(f"- {p}" for p in _SMB_BOT_SUGGESTED_PROMPTS[:6])
    return (
        "Sorry, I didn't catch that. I can answer questions about pinned **SMB Q1 2026** data — for example:\n\n"
        f"{sample}\n\nTry rephrasing, or click one of the suggested prompts."
    )


_ASSISTANT_SUGGESTED_PROMPTS = (
    "What is the total payout for Q1 2026 (reps + manager)?",
    "How much did Chitradip earn after clawbacks?",
    "Who is the top performer this quarter?",
    "Who is below the 50% eligibility threshold?",
    "Why is Lawrence's payout only $500?",
    "What are the SMB Group A vs Group B commission tiers?",
    "What's the AM team's total revenue and incentive in the database?",
    "List all sales reps in the Enterprise team.",
    "How many outbound meetings have been logged this quarter?",
    "What is the outbound meeting payout structure?",
    "Convert Vicky's payout to INR.",
    "Show clawbacks applied to Chitradip with their reasons.",
)


def render_compensation_assistant_page() -> None:
    """Sidebar 'Assistant' page: chat with Claude over the full Compensation Tool data."""
    st.subheader("Compensation Assistant")
    api_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    # Detect whether AI mode is actually available (key + SDK both present).
    _ai_available = False
    if api_key and api_key.startswith("sk-ant-"):
        try:
            import anthropic  # noqa: F401
            _ai_available = True
        except ImportError:
            _ai_available = False

    if _ai_available:
        model = (os.environ.get("ANTHROPIC_MODEL") or "claude-haiku-4-5-20251001").strip()
        st.caption(
            f"Powered by Anthropic Claude (`{model}`). Answers questions using your pinned data, "
            "AM/ENT team aggregates from the database, the user roster, outbound meetings, the commission policy, "
            "and any HubSpot deals you've fetched in this session."
        )
    else:
        st.caption(
            "Running in offline mode. I can answer questions about pinned periods (SMB Q1, AM Q1, SMB April, AM April) — "
            "ask about a specific rep, manager, team total, or eligibility. For free-form questions across the whole tool, "
            "set `ANTHROPIC_API_KEY` in `.env` (optional)."
        )

    cols = st.columns(2)
    for i, prompt in enumerate(_ASSISTANT_SUGGESTED_PROMPTS):
        with cols[i % 2]:
            if st.button(prompt, key=f"assistant_sug_{i}", use_container_width=True):
                hist = st.session_state.setdefault("compensation_chat_history", [])
                hist.append({"role": "user", "content": prompt})
                with st.spinner("Thinking…"):
                    answer = _ask_compensation_assistant(prompt, history=hist[:-1])
                hist.append({"role": "assistant", "content": answer})
                st.rerun()

    history = st.session_state.setdefault("compensation_chat_history", [])
    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_input := st.chat_input("Ask anything about commissions, reps, policy, outbound, HubSpot…"):
        history.append({"role": "user", "content": user_input})
        with st.spinner("Thinking…"):
            answer = _ask_compensation_assistant(user_input, history=history[:-1])
        history.append({"role": "assistant", "content": answer})
        st.rerun()

    if history:
        if st.button("Clear chat", key="assistant_chat_clear"):
            st.session_state["compensation_chat_history"] = []
            st.rerun()


def render_smb_chat_assistant(pinned: dict) -> None:
    """Render a chat box at the bottom of the SMB page that answers questions from the pinned JSON."""
    if not pinned:
        return
    st.markdown("---")
    st.markdown("### Ask about Q1 2026")
    st.caption(
        "I answer questions using the pinned data in `policy/smb_q1_2026_fixed.json` "
        "(reps, commissions, manager incentive, clawbacks). For broader questions across all teams, "
        "use the **Assistant** in the left sidebar."
    )

    # Suggested prompts as buttons (filling in the input).
    cols = st.columns(2)
    for i, prompt in enumerate(_SMB_BOT_SUGGESTED_PROMPTS):
        with cols[i % 2]:
            if st.button(prompt, key=f"smb_bot_sug_{i}", use_container_width=True):
                st.session_state.setdefault("smb_chat_history", []).append({"role": "user", "content": prompt})
                st.session_state["smb_chat_history"].append(
                    {"role": "assistant", "content": _answer_smb_question(pinned, prompt)}
                )
                st.rerun()

    history = st.session_state.setdefault("smb_chat_history", [])
    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_input := st.chat_input("Ask anything about Q1 2026 — payouts, clawbacks, eligibility…"):
        history.append({"role": "user", "content": user_input})
        history.append({"role": "assistant", "content": _answer_smb_question(pinned, user_input)})
        st.rerun()

    if history:
        if st.button("Clear chat", key="smb_chat_clear"):
            st.session_state["smb_chat_history"] = []
            st.rerun()


def _compute_pinned_summary(pinned: dict) -> dict:
    """Pre-compute every derived total / percentage / eligibility status from pinned data.

    The chat assistant sends this alongside the raw JSON so Claude can USE these values
    instead of recomputing them (which is where math errors happen).
    """
    reps = pinned.get("reps") or []
    comm = pinned.get("commission") or {}
    comm_reps = comm.get("reps") or []
    mgr = pinned.get("manager_incentive") or {}
    try:
        inr_rate = float(comm.get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0
    elig_min = float(comm.get("eligibility_min_pct") or 50)

    # Aggregate exceptions (Joy's Washington Post split, etc.)
    total_exception_usd = 0.0
    exception_lines: list[dict] = []
    for r in reps:
        exc = r.get("exception")
        if exc:
            try:
                share = float(exc.get("joy_share_usd") or exc.get("share_usd") or 0)
            except (TypeError, ValueError):
                share = 0.0
            if share:
                total_exception_usd += share
                exception_lines.append(
                    {
                        "rep": r.get("name"),
                        "deal_name": exc.get("deal_name"),
                        "deal_amount_usd": float(exc.get("deal_amount_usd") or 0),
                        "share_usd": share,
                        "note": exc.get("note"),
                    }
                )

    team_target = float(pinned.get("team_target_usd") or 0)
    raw_booked = sum(float(r.get("achievement_usd") or 0) for r in reps)
    raw_paid = sum(float(r.get("payment_received_usd") or 0) for r in reps)
    effective_booked = raw_booked + total_exception_usd
    effective_paid = raw_paid + total_exception_usd
    team_pct = (effective_booked / team_target * 100.0) if team_target else 0.0

    # Per-rep status
    rep_status: list[dict] = []
    for r in reps:
        nm = (r.get("name") or "").strip()
        quota = float(r.get("target_usd") or 0)
        ach = float(r.get("achievement_usd") or 0)
        paid = float(r.get("payment_received_usd") or 0)
        is_mgr = bool(r.get("is_manager"))
        left_org = bool(r.get("left_org"))
        ach_pct = (ach / quota * 100.0) if quota else 0.0
        if left_org:
            status_label = "Left org — no payout (deals still count toward team total)"
        elif is_mgr:
            status_label = "Manager — see manager_incentive section, not the rep commission summary"
        elif ach_pct < elig_min:
            status_label = f"Not eligible — {ach_pct:.0f}% below the {int(elig_min)}% floor"
        else:
            status_label = f"Eligible — {ach_pct:.0f}% achievement"
        rep_status.append(
            {
                "name": nm,
                "quota_usd": quota,
                "achievement_usd": ach,
                "payment_received_usd": paid,
                "achievement_pct": round(ach_pct, 2),
                "is_manager": is_mgr,
                "left_org": left_org,
                "status": status_label,
            }
        )

    # Commission summary aggregates.
    rep_payouts: list[dict] = []
    sum_base = 0.0
    sum_deduction = 0.0
    sum_manage = 0.0
    sum_payout = 0.0
    monthly_tiers = pinned.get("monthly_tiers")
    is_monthly_dynamic = bool(pinned.get("month")) and not comm_reps and bool(monthly_tiers)

    if is_monthly_dynamic:
        # Monthly path: rep payouts are computed dynamically from reps[] + monthly_tiers.
        ind_tiers = (monthly_tiers or {}).get("individual_tiers") or []
        md_pct_only_am = float((monthly_tiers or {}).get("manage_deal_pct_am_only") or 0)
        team_lbl_low = (mgr.get("team_label") or "").lower()
        is_am_team = ("account" in team_lbl_low) or team_lbl_low.startswith("am")
        for r in reps:
            if bool(r.get("is_manager")):
                continue
            nm = (r.get("name") or "").strip() or "—"
            quota_v = float(r.get("target_usd") or 0)
            ach_v = float(r.get("achievement_usd") or 0)
            paid_v = float(r.get("payment_received_usd") or 0)
            md_v = float(r.get("manage_deal_usd") or 0)
            left_org_v = bool(r.get("left_org"))
            ach_pct_v = (ach_v / quota_v * 100.0) if quota_v else 0.0
            slab_v, _slab_lbl = _monthly_slab_pct(ind_tiers, ach_pct_v)
            eligible_v = (ach_pct_v >= elig_min) and not left_org_v

            base_v = round(paid_v * slab_v / 100.0, 2) if eligible_v else 0.0
            md_amt_v = round(md_v * md_pct_only_am / 100.0, 2) if (is_am_team and md_v > 0 and eligible_v) else 0.0
            adj_total_v = round(sum(float(a.get("commission_usd") or 0) for a in (r.get("adjustments") or [])), 2)
            payout_v = round(base_v + md_amt_v + adj_total_v, 2)
            # Honour per-rep override (e.g. Yogi's $1,004 rounded total) so the chat assistant
            # stays in sync with the UI table.
            _override = r.get("total_payout_override_usd")
            if _override is not None:
                try:
                    payout_v = round(float(_override), 2)
                except (TypeError, ValueError):
                    pass

            sum_base += base_v
            sum_manage += md_amt_v
            sum_payout += payout_v
            status_note = (
                "Left org — no payout (deals still count toward team total)" if left_org_v
                else ("Not eligible — below threshold" if not eligible_v else "")
            )
            rep_payouts.append(
                {
                    "name": nm,
                    "quota_usd": quota_v,
                    "achievement_usd": ach_v,
                    "payment_received_usd": paid_v,
                    "achievement_pct": round(ach_pct_v, 2),
                    "slab_pct": round(slab_v, 2),
                    "eligible": eligible_v,
                    "left_org": left_org_v,
                    "base_commission_usd": base_v,
                    "manage_deal_usd": md_amt_v,
                    "adjustments_usd": adj_total_v,
                    "adjustments_details": list(r.get("adjustments") or []),
                    "total_payout_usd": payout_v,
                    "total_payout_inr": round(payout_v * inr_rate, 2) if inr_rate else 0,
                    "status_note": status_note,
                }
            )
    else:
        # Quarterly path: use pre-stored values from commission.reps[].
        for c in comm_reps:
            base = float(c.get("base_commission_usd") or 0)
            try:
                ded_v = float(c.get("deduction_usd"))
            except (TypeError, ValueError):
                try:
                    ded_pct = float(c.get("deduction_pct"))
                except (TypeError, ValueError):
                    ded_pct = float(comm.get("team_deduction_pct") or 0)
                ded_v = round(base * ded_pct / 100.0, 2)
            md_v = float(c.get("manage_deal_usd") or 0)
            md_paid_now = float(c.get("manage_deal_paid_now_usd") or 0)
            payout = float(c.get("total_payout_usd") or 0)
            quota_v = float(c.get("quota_usd") or 0)
            rev_v = float(c.get("revenue_achieved_usd") or 0)
            paid_v = float(c.get("payment_received_usd") or 0)
            elig_pct_v = float(c.get("eligible_pct") or 0)
            sum_base += base
            sum_deduction += ded_v
            sum_manage += md_v
            sum_payout += payout
            rep_payouts.append(
                {
                    "name": c.get("name"),
                    "group": c.get("group") or c.get("compensation_group") or "",
                    "quota_usd": quota_v,
                    "revenue_achieved_usd": rev_v,
                    "payment_received_usd": paid_v,
                    "achievement_pct": elig_pct_v,
                    "base_commission_usd": base,
                    "deduction_usd": ded_v,
                    "manage_deal_usd": md_v,
                    "manage_deal_paid_now_usd": md_paid_now,
                    "total_payout_usd": payout,
                    "total_payout_inr": round(payout * inr_rate, 2) if inr_rate else 0,
                    "note": (c.get("note") or "").strip(),
                    "adjustments_details": list(c.get("adjustments") or []),
                }
            )

    # Manager incentive (final + breakdown)
    mgr_base = float(mgr.get("manager_q1_amount_usd") or 0)
    mgr_pending = float(
        mgr.get("total_pending_commission_usd")
        or sum(float(p.get("commission_usd") or 0) for p in (mgr.get("pending_commissions") or []))
    )
    mgr_clawback = float(
        mgr.get("total_clawback_usd")
        or sum(float(c.get("deduction_usd") or 0) for c in (mgr.get("clawbacks") or []))
    )
    mgr_final = mgr.get("final_payout_usd")
    if mgr_final is None:
        mgr_final = mgr_base + mgr_pending - mgr_clawback
    mgr_final = float(mgr_final)

    # Detect when the "manager" is the same person as one of the reps in commission.reps[]
    # (e.g. Anthony Raymond for ENT — sole rep & treated as manager in the JSON for pending
    # commission / clawback bookkeeping). If so, the rep's total_payout_usd ALREADY includes
    # those manager-level pending/clawback adjustments, so we must NOT double-count them in
    # grand_total.
    mgr_name_low = (mgr.get("manager_name") or "").strip().lower()
    manager_is_also_a_rep = bool(mgr_name_low) and any(
        (rp.get("name") or "").strip().lower() == mgr_name_low for rp in rep_payouts
    )
    if manager_is_also_a_rep:
        grand_total = sum_payout
    else:
        grand_total = sum_payout + max(mgr_final, 0)

    # Consolidate ALL prior-period adjustments so the chat bot can list them in one go.
    # This includes: rep-level adjustments, manager-level clawbacks, manager-level pending
    # commissions, and team-level exceptions. Users sometimes call all of these "clawbacks".
    prior_items: list[dict] = []
    for r in reps:
        nm = (r.get("name") or "").strip()
        for adj in (r.get("adjustments") or []):
            prior_items.append({
                "kind": "Rep adjustment",
                "rep_or_manager": nm,
                "deal_name": adj.get("deal_name"),
                "period": adj.get("period"),
                "amount_usd": float(adj.get("commission_usd") or 0),
                "deal_amount_usd": float(adj.get("paid_amount_usd") or adj.get("deal_amount_usd") or 0),
                "rate_pct": float(adj.get("rate_pct") or 0),
                "direction": "+" if float(adj.get("commission_usd") or 0) >= 0 else "−",
                "note": adj.get("note"),
            })
    for cb in (mgr.get("clawbacks") or []):
        prior_items.append({
            "kind": "Manager clawback",
            "rep_or_manager": mgr.get("manager_name"),
            "deal_name": cb.get("label") or cb.get("deal_name"),
            "period": cb.get("period"),
            "amount_usd": -float(cb.get("deduction_usd") or 0),
            "deal_amount_usd": float(cb.get("deal_amount_usd") or 0),
            "rate_pct": float(cb.get("rate_pct") or 0),
            "direction": "−",
            "note": cb.get("note"),
        })
    for p in (mgr.get("pending_commissions") or []):
        prior_items.append({
            "kind": "Manager pending commission",
            "rep_or_manager": mgr.get("manager_name"),
            "deal_name": p.get("deal_name"),
            "period": p.get("period"),
            "amount_usd": float(p.get("commission_usd") or 0),
            "deal_amount_usd": float(p.get("deal_amount_usd") or 0),
            "rate_pct": float(p.get("rate_pct") or 0),
            "direction": "+",
            "note": p.get("note"),
        })
    for ex in exception_lines:
        prior_items.append({
            "kind": "Team exception",
            "rep_or_manager": ex.get("rep"),
            "deal_name": ex.get("deal_name"),
            "period": "current",
            "amount_usd": float(ex.get("share_usd") or 0),
            "deal_amount_usd": float(ex.get("deal_amount_usd") or 0),
            "rate_pct": 0,
            "direction": "+",
            "note": ex.get("note"),
        })

    return {
        "period": {
            "year": pinned.get("year"),
            "quarter": pinned.get("quarter"),
            "month": pinned.get("month"),
            "month_label": pinned.get("month_label"),
        },
        "team": {
            "target_usd": team_target,
            "achievement_raw_usd": round(raw_booked, 2),
            "payment_received_raw_usd": round(raw_paid, 2),
            "total_exception_usd": round(total_exception_usd, 2),
            "achievement_effective_usd": round(effective_booked, 2),
            "payment_received_effective_usd": round(effective_paid, 2),
            "achievement_pct": round(team_pct, 2),
            "eligibility_floor_pct": elig_min,
        },
        "exceptions": exception_lines,
        "rep_status": rep_status,
        "rep_payouts": rep_payouts,
        "rep_payouts_totals": {
            "sum_base_commission_usd": round(sum_base, 2),
            "sum_deduction_usd": round(sum_deduction, 2),
            "sum_manage_deal_usd": round(sum_manage, 2),
            "sum_total_payout_usd": round(sum_payout, 2),
            "sum_total_payout_inr": round(sum_payout * inr_rate, 2) if inr_rate else 0,
        },
        "manager_incentive_summary": {
            "manager_name": mgr.get("manager_name"),
            "base_usd": round(mgr_base, 2),
            "pending_total_usd": round(mgr_pending, 2),
            "clawback_total_usd": round(mgr_clawback, 2),
            "final_payout_usd": round(mgr_final, 2),
            "final_payout_inr": round(mgr_final * inr_rate, 2) if inr_rate else 0,
            "carry_forward_usd": float(mgr.get("carry_forward_usd") or 0),
            "is_eligible_for_base": team_pct >= float(mgr.get("manager_team_minimum_pct") or 60),
            "calculation_note": (mgr.get("calculation_note") or "").strip(),
        },
        "grand_total": {
            "payout_usd": round(grand_total, 2),
            "payout_inr": round(grand_total * inr_rate, 2) if inr_rate else 0,
        },
        "manager_is_also_a_rep": manager_is_also_a_rep,
        "all_prior_period_items": prior_items,
        "inr_rate_per_usd": inr_rate,
    }


def _ask_pinned_with_context(
    question: str,
    pinned: dict,
    history: list[dict] | None,
    stream_placeholder=None,
) -> str:
    """Answer a question using Claude with the pinned dict + pre-computed summary as context.

    When ``stream_placeholder`` is provided (a ``st.empty()`` slot), the response streams
    progressively into it (ChatGPT-style). Otherwise returns the full response at once.
    Falls back to the rule-based SMB bot if no API key is configured (SMB Q1 only).
    """
    import json as _json

    api_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    # Treat placeholder keys (any value that doesn't look like a real Anthropic key) as "no key".
    if api_key and not api_key.startswith("sk-ant-"):
        api_key = ""

    # Try to import the SDK only when a real key is present. If either is missing, fall back
    # silently to the rule-based bot — no scary error message.
    anthropic_mod = None
    if api_key:
        try:
            import anthropic as anthropic_mod  # type: ignore
        except ImportError:
            anthropic_mod = None

    if anthropic_mod is None:
        return _answer_pinned_period(pinned, question)

    # Default to Sonnet for math accuracy on this assistant; user can override via env.
    model = (os.environ.get("ANTHROPIC_PINNED_CHAT_MODEL") or os.environ.get("ANTHROPIC_MODEL") or "claude-sonnet-4-6").strip()

    summary = _compute_pinned_summary(pinned)
    summary_json = _json.dumps(summary, default=str, ensure_ascii=False, indent=2)
    raw_json = _json.dumps(pinned, default=str, ensure_ascii=False, indent=2)
    if len(raw_json) > 25000:
        raw_json = raw_json[:25000] + "\n... (raw JSON truncated)"

    quarter = pinned.get("quarter")
    month = pinned.get("month")
    year = pinned.get("year")
    if quarter:
        period_desc = f"Q{quarter} FY{year}"
    elif month:
        period_desc = pinned.get("month_label") or f"{year}-{month:02d}"
    else:
        period_desc = "this period"
    team_desc = (pinned.get("manager_incentive") or {}).get("team_label") or ""

    system_prompt = (
        f"You are a helpful, conversational AI assistant — like ChatGPT — for the **{team_desc} {period_desc}** sales "
        "compensation page. You answer questions naturally and thoroughly, in a friendly tone. You're allowed "
        "to elaborate, give context, and follow up with related insights the user might find useful.\n\n"
        "**CRITICAL — STRICT TEAM ISOLATION (this is a hard rule, not a guideline):**\n"
        f"This assistant is SCOPED ONLY to the **{team_desc}** team for **{period_desc}**. You MUST refuse "
        "to answer questions about other teams. The teams in this product are: SMB, Account Management (AM), "
        "Enterprise (ENT), and Outbound. They have DIFFERENT policies, different rate cards, different reps, "
        "different managers, and different pinned data files.\n"
        f"- If asked about a team OTHER than {team_desc} (e.g. \"What is Anthony's commission?\" on the AM page, "
        "or \"Who is the SMB manager?\" on the ENT page, or anything about Chitradip / Lawrence / Yogi on the "
        "AM or ENT pages, or anything about Anthony on the SMB/AM pages), respond with: "
        "\"I'm scoped to the {team_desc} page only. For that question, please open the relevant team's page "
        "and ask there.\" Then briefly tell them which page to use (SMB / AM / ENT / Outbound).\n"
        "- NEVER pull figures from one team to answer about another. NEVER conflate Joy (AM manager) with "
        "Chitradip (SMB manager) or with Anthony (ENT sole rep). Each is on their own page.\n"
        "- The `pinned` JSON you're given is the SINGLE SOURCE OF TRUTH for this page. Do not reference "
        "deals, rates, or people that aren't in this JSON.\n\n"
        "**HOW TO USE THE DATA — CRITICAL FOR ACCURACY:**\n"
        "1. A `summary` object is provided below with ALL totals, percentages, eligibility statuses, and payouts "
        "pre-computed by Python. USE these values directly. NEVER do your own arithmetic — Python's math is reliable, "
        "yours isn't.\n"
        "2. The `pinned` raw JSON is for context (notes, rules text). Don't derive totals from `pinned.reps[]` — "
        "use `summary.team.*` and `summary.rep_status`.\n"
        "3. Per-page conventions:\n"
        "   • Reps with `is_manager: true` are excluded from the rep commission summary (they appear under manager_incentive).\n"
        "   • Reps with `left_org: true` get NO payout but their deals still count toward team totals.\n"
        "   • Exceptions (e.g. Washington Post server split) are pre-added to `summary.team.achievement_effective_usd` "
        "and `payment_received_effective_usd`.\n"
        "   • Pending commissions are prior-period deals paid in this period; clawbacks are deductions from prior periods.\n"
        "   • **Eligibility floor for monthly plan (April onwards) is 60% — STRICT.** A rep at 59.9%, 59%, or anything below "
        "60% earns 0% commission. Always explicitly state this when explaining a non-eligible rep on a monthly page. "
        "The Q1 quarterly floor is 50%.\n"
        "   • **ENT-specific:** the Enterprise team has only one person — Anthony Raymond. There is NO separate manager "
        "incentive; the `manager_incentive` block in the JSON represents Anthony's own commission. ENT slab bands: 0–75.99% → 7%, "
        "76–100% → 9%, 101–125% → 11%, 126%+ → 13%. ENT uses USD only (no INR conversion). Never describe Anthony as the "
        "'manager' or use 'manager pool' language for ENT — just call it 'Anthony's commission'.\n"
        "   • **CRITICAL — manager_is_also_a_rep guard:** when `summary.manager_is_also_a_rep` is `true` (always true for ENT), "
        "the rep's `total_payout_usd` ALREADY includes the pending commissions and clawbacks listed under `manager_incentive`. "
        "Do NOT add `rep_payouts_totals.sum_total_payout_usd` and `manager_incentive_summary.final_payout_usd` together — "
        "that double-counts. In this case the correct total is simply `summary.grand_total.payout_usd` (which equals the rep's "
        "single total_payout_usd). Never write a 'Manager final payout' line, never refer to 'manager pool', and never say "
        "'Grand total = rep + manager' for ENT. Just present **Total payout: $X** with the breakdown of base + pending − clawback.\n"
        "   • **AM-specific:** the Account Management team for Q1 2026 has 3 reps + 1 manager (Joy is both a rep and the "
        "manager — `is_manager: true`). Reps: Arundhati Sen, Vivin Joseph, Joy Prakash. Nikita Shekher also owns a Q1 deal "
        "(NerdsToGo) even though she's not a commissioned AM rep. AM Q1 commission rate is 1% of payment received (with a "
        "10% deduction for reps without a manage deal). AM eligibility floor is 50% of individual quota. The manager Joy is "
        "paid 1.5% of team payment received minus the Wipro Q4 clawback. AM uses USD with INR conversion at ₹86/USD.\n"
        "   • **AM deal-level data:** `pinned.paid_deals` (33 entries — includes 1 partially-paid Artnet deal of $11,566 with "
        "`is_partial: true`) and `pinned.unpaid_deals_overrides` (4 entries — Allied Gold, Estée Lauder, Church & Dwight, "
        "NerdsToGo, plus the Artnet partial remainder of $11,566 with `is_partial_remainder: true`). Total = 37 deals "
        "closed in Q1: 33 with payment received (incl. 1 partial), 4 with no payment. Deal owners: Vivin (22), Joy (11), "
        "Arundhati (3), Nikita (1). When asked about deal counts, owner breakdown, or rep-level pending vs received, surface "
        "this data with the per-owner totals.\n"
        "   • **Partial-payment math (Artnet only):** Artnet appears in BOTH `paid_deals` ($11,566 paid portion, `is_partial: true`) "
        "AND `unpaid_deals_overrides` ($11,566 unpaid remainder, `is_partial_remainder: true`). The total deal value is "
        "$23,132. For deal counts, treat Artnet as ONE deal (not two). For dollar totals, both halves are real.\n"
        "   • **`summary.all_prior_period_items`** is the SINGLE LIST of every prior-period item: manager clawbacks, "
        "manager pending commissions, rep-level adjustments (e.g. Yogi's Yieldstreet/Willow Wealth), and team exceptions "
        "(e.g. Washington Post). When asked about ANY of: clawback, adjustment, exception, pending, carry-over, or a "
        "specific deal name like 'Willow Wealth' / 'Wipro' / 'Krish', list everything from this array grouped by kind. "
        "Users often use the words 'clawback' / 'adjustment' / 'exception' interchangeably — always list all of them.\n"
        "4. Format: USD with `$` and thousands separators ($2,561.34). INR with `₹` using `summary.inr_rate_per_usd`.\n"
        "5. Show your reasoning step-by-step and cite the actual numbers — users want to see how the figure was built.\n"
        "6. If the question can't be answered from the provided data, say so honestly. Don't invent.\n\n"
        "**STYLE — be ChatGPT-like:**\n"
        "- Open with a direct answer, then explain.\n"
        "- Use markdown headings, bullet lists, and short paragraphs.\n"
        "- Wrap calculations in code blocks for clarity.\n"
        "- Volunteer relevant context: e.g. if asked about Joy's payout, mention the Wipro clawback and the Washington Post exception even if not explicitly asked.\n"
        "- Offer a follow-up suggestion at the end (\"Want me to also compare last quarter's payout?\" — only if relevant).\n"
        "- Keep tone friendly, not robotic.\n\n"
        "**WORKED EXAMPLES:**\n\n"
        "Q (SMB/AM): \"What is the total payout for this month including the manager?\"\n"
        "A: The total payout is **$X** (≈ ₹Y @ ₹86/USD).\n"
        "Here's how it breaks down:\n"
        "- Sales rep payouts (sum from `rep_payouts_totals.sum_total_payout_usd`): $A\n"
        "- Manager final payout (`manager_incentive_summary.final_payout_usd`): $B\n"
        "- Grand total: $X\n"
        "(Note: Vivin shows $0 because he left the org — his deals still count toward the team though.)\n\n"
        "Q (AM): \"Why is Joy eligible this month?\"\n"
        "A: Joy is eligible because the team achievement is **{team_pct}%**, which is above the **60% floor**.\n"
        "- Team achievement (effective): $X (raw HubSpot $Y + Washington Post exception $17,500)\n"
        "- That puts the team in the 81–99% band → 3% rate\n"
        "- Joy's manager commission: $X × 3% = $...\n"
        "- After Q1 carry-forward ($78) and pending Q1 deals ($427.22 from Krish, Artnet, Church & Dwight), her final is $2,561.34.\n\n"
        "**ENT WORKED EXAMPLES (Anthony Raymond, Q1 FY2026):**\n\n"
        "Q (ENT): \"What is Anthony's total Q1 FY2026 payout, including Q4 pending deals?\"\n"
        "A: Anthony's total Q1 FY2026 payout is **$22,055.57**.\n\n"
        "**Performance**\n"
        "- Quota: $1,000,000\n"
        "- Revenue achieved (booked): $453,168.40 (45.3% of quota)\n"
        "- Payment received (after excluding 3 Not-Paid deals totaling $198,874.50): **$254,293.90**\n\n"
        "**Base commission (Q1)**\n"
        "- 45.3% lands in the **0–75.99% ENT slab band → 7%** commission rate.\n"
        "- Commission is calculated on **Payment Received**, not booked revenue:\n"
        "- $254,293.90 × 7% = **$17,800.57**\n\n"
        "**Q4 pending commissions (paid in Q1 at 10% Q4 rate)**\n"
        "- Armstrong World Industries — BDO USA, LLP — Partner Deal ($28,008 × 10%) = +$2,801\n"
        "- Chryselys — Egnyte to MS SharePoint (Remaining 25%) ($5,292 × 10%) = +$529\n"
        "- Forvis Mazars ($18,000 × 10%) = +$1,800\n"
        "- **Subtotal: +$5,130**\n\n"
        "**Refunds / clawbacks**\n"
        "- Cambrex — Dropbox to MS (Q4 cancelled / refunded $8,747 × 10%) = **−$875**\n\n"
        "**Final math**\n"
        "- $17,800.57 + $5,130 − $875 = **$22,055.57**\n\n"
        "Q (ENT): \"Which Q1 deals were Not Paid? How are they excluded from Payment Received?\"\n"
        "A: Three Q1 deals were marked **Not Paid** in HubSpot at the time the Q1 commission was calculated. "
        "Per ENT policy, commission is calculated on **Payment Received**, not booked revenue — so these deals are excluded from the commission base until payment arrives.\n\n"
        "| Deal | Close date | Amount |\n"
        "| --- | --- | --- |\n"
        "| The Estée Lauder Companies Inc. — Dropbox to SharePoint Online | 2026-02-06 | $160,000.00 |\n"
        "| Experian Information Solutions, Inc. — Slack to Teams AIO | 2026-03-31 | $4,356.00 |\n"
        "| CloudSoft — Google to MS | 2026-03-26 | $34,518.50 |\n\n"
        "**Total unpaid: $198,874.50**\n\n"
        "**Math:** Booked $453,168.40 − Not Paid $198,874.50 = **Payment Received $254,293.90**. "
        "Anthony's Q1 base commission is therefore $254,293.90 × 7% = **$17,800.57** — not $453,168.40 × 7%. "
        "Once payment arrives for any of these deals, the corresponding commission is released in a future cycle.\n\n"
        "Q (ENT): \"What is the Cambrex refund and how is it deducted?\"\n"
        "A: The **Cambrex — Dropbox to MS** deal was a Q4 FY2025 deal that was later cancelled and the customer was refunded $8,747. "
        "The commission that was previously paid out on that deal at the **Q4 rate (10%)** is now clawed back from Q1:\n"
        "- Refund amount: $8,747\n"
        "- Rate previously paid: 10% (Q4 FY2025 rate)\n"
        "- Clawback: $8,747 × 10% = **−$875**\n\n"
        "**How it flows into Anthony's total:** Q1 base $17,800.57 + Q4 pending +$5,130 − Cambrex clawback −$875 = **$22,055.57**.\n\n"
        "Q (ENT): \"Explain the ENT commission calculation step by step.\"\n"
        "A: Use this exact 7-step structure with numbered headings — never abbreviate when the user asks "
        "for a step-by-step. Each step has a heading, policy context, and math.\n\n"
        "## Q1 FY2026 — ENT commission calculation, step by step\n\n"
        "### Step 1 — Pull the team's Q1 performance\n"
        "- Quota: $1,000,000\n"
        "- Booked revenue: $453,168.40 → 45.3% of quota\n\n"
        "### Step 2 — Exclude Q1 deals that were Not Paid\n"
        "ENT policy: commission is calculated on **Payment Received**, not booked revenue. Any deal still "
        "marked Not Paid in HubSpot at the time the commission is run is excluded from the base.\n"
        "- The Estée Lauder Companies — Dropbox to SharePoint ($160,000.00)\n"
        "- Experian — Slack to Teams AIO ($4,356.00)\n"
        "- CloudSoft — Google to MS ($34,518.50)\n"
        "- **Total Not-Paid: $198,874.50**\n"
        "Math: $453,168.40 − $198,874.50 = **Payment Received $254,293.90**.\n\n"
        "### Step 3 — Determine the slab band\n"
        "ENT slab bands: 0–75.99% → 7%, 76–100% → 9%, 101–125% → 11%, 126%+ → 13%.\n"
        "At 45.3% of quota, Anthony lands in the 0–75.99% band → **7%** commission rate. "
        "(ENT has a 0% eligibility floor — any positive revenue qualifies.)\n\n"
        "### Step 4 — Calculate Q1 base commission\n"
        "Base = Payment Received × Slab % = $254,293.90 × 7% = **$17,800.57**.\n"
        "We use Payment Received, NOT booked revenue.\n\n"
        "### Step 5 — Add Q4 pending commissions (paid in Q1 at the Q4 10% rate)\n"
        "- Armstrong World Industries — BDO USA, LLP ($28,008 × 10%) = +$2,801\n"
        "- Chryselys — Egnyte to MS SharePoint (Remaining 25%) ($5,292 × 10%) = +$529\n"
        "- Forvis Mazars ($18,000 × 10%) = +$1,800\n"
        "- **Subtotal pending: +$5,130**\n\n"
        "### Step 6 — Deduct Q4 refunds / clawbacks\n"
        "Cambrex — Dropbox to MS (cancelled Q4 deal, $8,747 refunded × 10%) = **−$875**.\n\n"
        "### Step 7 — Add it all up\n"
        "Q1 base $17,800.57 + Q4 pending +$5,130 − Q4 clawback −$875 = **$22,055.57**.\n\n"
        "### ✅ Final Q1 FY2026 payout: **$22,055.57** (USD only — no INR conversion for ENT)\n\n"
        "**AM WORKED EXAMPLES (Joy / Vivin / Arundhati, Q1 FY2026):**\n\n"
        "Q (AM): \"How many deals did the AM team close, how many paid, how many not paid?\"\n"
        "A:\n"
        "- 📦 **Total deals closed**: **37**\n"
        "- ✅ **Payment received**: **33 deal(s)** — total $100,107.40 (includes 1 partially-paid Artnet at $11,566)\n"
        "- ❌ **Payment not received**: **4 deal(s)** — excluded value $56,697 (Allied Gold $1,800, Estée Lauder $26,416, Church & Dwight $14,615, NerdsToGo $2,300, plus Artnet partial remainder $11,566)\n\n"
        "Q (AM): \"Break down deals by owner — how many did each rep close?\"\n"
        "A:\n"
        "| Deal Owner | Total | Payment received | Payment not received | Booked Amount | $ Received | $ Pending |\n"
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: |\n"
        "| **Vivin Joseph** | 22 | 21 | 1 | $61,981 | $60,181 | $1,800 |\n"
        "| **Joy Prakash** | 11 | 9 (1-Partial Paid) | 2 (1-Partial Payment not received) | $88,920 | $36,323 | $52,597 |\n"
        "| **Arundhati Sen** | 3 | 3 | 0 | $3,600 | $3,600 | $0 |\n"
        "| **Nikita Shekher** | 1 | 0 | 1 | $2,300 | $0 | $2,300 |\n"
        "| **Total (AM team)** | **37** | **33** | **4** | **$156,801** | **$100,104** | **$56,697** |\n\n"
        "Q (AM): \"How much is pending vs received at rep level?\"\n"
        "A: At rep level, the $ pending = sum of their Not-Paid deal amounts (incl. partial remainders). Received = sum of paid amounts (incl. partial paid portions).\n"
        "- **Vivin Joseph**: Received $60,181 · Pending $1,800 (Allied Gold)\n"
        "- **Joy Prakash**: Received $36,323 · Pending $52,597 (Estée Lauder $26,416 + Church & Dwight $14,615 + Artnet partial remainder $11,566)\n"
        "- **Arundhati Sen**: Received $3,600 · Pending $0\n"
        "- **Nikita Shekher**: Received $0 · Pending $2,300 (NerdsToGo - DropBox to MS for Carolina Contracting)\n"
        "- **Team total**: Received **$100,104** · Pending **$56,697**\n\n"
        "Q (AM): \"Why is Arundhati Sen not eligible?\"\n"
        "A: This is a question about **Arundhati only** — do NOT show other reps' figures.\n\n"
        "**Performance vs quota**\n"
        "- Individual quota: $50,000\n"
        "- Revenue achieved (booked): $3,600 → **7% of quota**\n"
        "- Payment received: $3,600\n\n"
        "**Eligibility check**\n"
        "AM Q1 eligibility floor: **50% of individual quota**. Arundhati is at **7%** — far below the 50% floor → **not eligible**. ❌\n\n"
        "**Commission result**: Because Arundhati did not meet the 50% floor, her base commission for Q1 is **$0.00**. Her 3 deals "
        "(Ampleo $1,000, Greene IS $1,000, IT Solutions for Renovus $1,600) still count toward team Booked Amount / Payment Received "
        "(which flow into Joy's manager incentive), but they don't earn Arundhati personal commission this quarter.\n\n"
        "### ❌ Final Q1 FY2026 payout: **$0.00**\n\n"
        "Q (AM): \"Explain Vivin Joseph's Q1 commission — base, 10% deduction, and the Vendasta manage deal.\"\n"
        "A: This is a question about **Vivin only** — do NOT show other reps' figures.\n\n"
        "**Performance vs quota**\n"
        "- Individual quota: $100,000\n"
        "- Revenue achieved (booked): $61,981 → **62% of quota** ✅ above 50% floor\n"
        "- Payment received: $60,181 (Allied Gold $1,800 is Not Paid → excluded)\n\n"
        "**Base commission (1% of Payment Received)**\n"
        "Base = $60,181 × 1% = **$602**\n\n"
        "**10% deduction** (AM policy — reps without a Q1 manage deal incur a 10% deduction)\n"
        "Deduction = $602 × 10% = **−$60**\n"
        "After deduction: $542\n\n"
        "**Q4 Vendasta CloudFuze Manage Deal incentive** (paid in Q1 at the 10% manage-deal rate)\n"
        "Manage deal = $6,300 × 10% = **+$630**\n\n"
        "**Final math**\n"
        "$602 − $60 + $630 = **$1,172**\n\n"
        "### ✅ Final Q1 FY2026 payout: **$1,172.00** (≈ ₹100,792 @ ₹86/USD)\n\n"
        "Q (AM): \"Explain Joy's manager commission calculation step by step.\"\n"
        "A: Joy is both a rep and the AM manager. Her commission flows entirely through `manager_incentive`:\n"
        "1. **Team Payment Received**: $100,107 (booked $156,801 − $56,698 not paid).\n"
        "2. **Manager rate**: 1.5% of team payment received → $100,107 × 1.5% = $1,502.\n"
        "3. **Q4 pending**: +$55 (Krish Services Group MARS overage — Q4 deal paid in Q1 at the Q4 1% rate).\n"
        "4. **Q4 Wipro clawback**: $54,515 × 3% = −$1,635 (Wipro Q4 deal cancelled; the slab was originally 4% but the correct slab is 1%, so 3% × $54,515 is clawed back).\n"
        "5. **Final**: $1,502 + $55 − $1,635 = **−$78**. No payout for Q1 — the −$78 balance carries forward to the next cycle.\n\n"
        "**SMB WORKED EXAMPLES (Chitradip + 7 reps, Q1 FY2026):**\n\n"
        "Q (SMB): \"How many deals did the SMB team close — paid, not paid, cancelled?\"\n"
        "A:\n"
        "- 📦 **Total effective deals closed**: **44** (cancellations excluded)\n"
        "- ✅ **Payment received**: **40 deal(s)** — total **$354,597** (official from `manager_incentive.team_payment_received_usd`)\n"
        "- ❌ **Payment not received**: **4 deal(s)** — excluded value **$71,336** (Legal Soft $20,706, CTI Technology $3,262, Allied Gold $2,500, OnCloud $44,868)\n"
        "- ⚠️ **Cancelled**: **1 deal** — Keytel Systems $2,000 (closed Won 2026-03-31 but later cancelled — removed from team achievement)\n\n"
        "Team-level: Booked Amount **$425,934** · Payment Received **$354,597** · Pending **$71,336**.\n\n"
        "Q (SMB): \"Explain Lawrence Lewis's 2 manage deals.\"\n"
        "A: Lawrence Lewis closed **2 manage deals** in Q1 FY2026:\n"
        "1. **The Commit Partnership - CF Manage** ($1,319, Paid, Closed Won 2026-03-10) — payment received → manage deal commission of $500 **paid this cycle**.\n"
        "2. **Legal Soft - Migrate(GWS to GWS) + Manage** ($20,706, Paid in HubSpot but payment not yet processed at the commission-cutoff date) — This deal combines a Migration ($9,942) and a Manage portion ($10,764). Because the payment hasn't been received, the $500 manage commission is **pending next cycle** and will release once payment arrives.\n\n"
        "**Math**: 2 manage deals × $500 = **$1,000 total earned**. **$500 paid this cycle** (Commit Partnership); **$500 pending** (Legal Soft).\n\n"
        "Lawrence is at **49% of quota** ($73,235.50 / $150,000), **below the 50% eligibility floor**, so his regular base commission is $0. However, manage deal incentive is paid regardless of overall quota achievement — it has its own eligibility separate from the regular commission floor.\n\n"
        "**Final Q1 payout for Lawrence: $500** (manage deal portion). _Additionally, Lawrence left the organization in April, so no further payouts were processed afterward._\n\n"
        "Q (SMB): \"Why was the Keytel Systems deal cancelled?\"\n"
        "A: **Keytel Systems - Box to SharePoint** ($2,000, Owner: Yogesh Vig) was Closed Won on 2026-03-31, but the deal was **subsequently cancelled**. The impact:\n"
        "- **Team Booked Amount** drops from $427,934 to **$425,934** (−$2,000)\n"
        "- **Yogesh Vig's individual achievement** drops from $73,563 (73.6% of quota) to **$71,563** (71.6% of quota)\n"
        "- Payment Received is **unaffected** (the deal was Not Paid anyway — no money was collected)\n\n"
        "Q (SMB): \"Explain Chitradip's manager commission step by step (incl. all 3 Q4 clawbacks).\"\n"
        "A:\n"
        "1. **Team target**: $850,000. **Team payment received**: $354,597. Team achievement: 50% → eligible.\n"
        "2. **Manager rate**: 1.5% (flat) → Base = $354,597 × 1.5% = **$5,319**\n"
        "3. **Q4 clawbacks** (3 items):\n"
        "   - Q4 Royston / Synergy Employment: $4,395.35 × 1% = **−$44** ($4,395.35 counted but not received | Chit 1% slab)\n"
        "   - Q4 Yogi / CEG Solutions ($15,009) + Yieldstreet ($17,068): $32,077 × 1% = **−$321** (Chit 1% slab)\n"
        "   - Noventiq Q3 deduction: $13,410 × 1.5% = **−$201** (1.5% on $13,410)\n"
        "   - **Total clawback: −$566**\n"
        "4. **Final**: $5,319 − $566 = **$4,753** ✅ (≈ ₹408,758 @ ₹86/USD)\n"
    )

    msgs: list[dict] = []
    for h in (history or [])[-6:]:
        if h.get("role") in ("user", "assistant") and (h.get("content") or "").strip():
            msgs.append({"role": h["role"], "content": h["content"]})
    msgs.append(
        {
            "role": "user",
            "content": (
                f"Question: {question}\n\n"
                f"---\n**SUMMARY (use these pre-computed values):**\n```json\n{summary_json}\n```\n\n"
                f"**RAW PINNED JSON (for context / notes):**\n```json\n{raw_json}\n```"
            ),
        }
    )

    anthropic = anthropic_mod  # alias so existing code below keeps working
    try:
        client = anthropic.Anthropic(api_key=api_key)
        if stream_placeholder is not None:
            # Stream the response so it appears progressively (ChatGPT-style).
            buf = ""
            with client.messages.stream(
                model=model,
                max_tokens=1500,
                system=system_prompt,
                messages=msgs,
            ) as stream:
                for text_chunk in stream.text_stream:
                    buf += text_chunk
                    try:
                        stream_placeholder.markdown(buf + "▌")
                    except Exception:
                        pass
            try:
                stream_placeholder.markdown(buf)
            except Exception:
                pass
            return buf.strip() or "(empty response)"
        else:
            # Non-streaming (suggested-prompt buttons or any caller without a placeholder).
            resp = client.messages.create(
                model=model,
                max_tokens=1500,
                system=system_prompt,
                messages=msgs,
            )
            chunks = [c.text for c in (resp.content or []) if getattr(c, "type", "text") == "text"]
            return "\n".join(chunks).strip() or "(empty response)"
    except Exception as e:
        return f"Anthropic API error: {e}"


def render_pinned_chat_assistant(pinned: dict, session_prefix: str = "pinned") -> None:
    """Inline chat assistant that answers questions about ANY pinned compensation period.

    Auto-detects whether the data is quarterly (has ``quarter``) or monthly (has ``month``)
    and labels the section accordingly. Uses a unique session key per period so chats on
    different pages don't share history.
    """
    if not pinned:
        return

    year = pinned.get("year")
    quarter = pinned.get("quarter")
    month = pinned.get("month")
    if quarter:
        period_key = f"q{quarter}_{year}"
        period_lbl = f"Q{quarter} FY{year}"
    elif month:
        period_key = f"m{month:02d}_{year}"
        period_lbl = pinned.get("month_label") or f"{year}-{month:02d}"
    else:
        period_key = "period"
        period_lbl = "this period"

    team_label = (pinned.get("manager_incentive") or {}).get("team_label") or ""
    tl_low = team_label.lower()
    sp_low = session_prefix.lower()
    if "account" in tl_low or sp_low.startswith("am"):
        team_short = "AM"
    elif "enterprise" in tl_low or "ent" in tl_low or sp_low.startswith("ent"):
        team_short = "ENT"
    elif "outbound" in tl_low or sp_low.startswith("ob") or sp_low.startswith("outbound"):
        team_short = "Outbound"
    else:
        team_short = "SMB"

    session_key = f"chat_history_{session_prefix}_{team_short.lower()}_{period_key}"

    st.markdown("---")
    st.markdown(f"### Ask about {team_short} · {period_lbl}")
    api_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    _ai_ok = False
    if api_key and api_key.startswith("sk-ant-"):
        try:
            import anthropic  # noqa: F401
            _ai_ok = True
        except ImportError:
            _ai_ok = False
    if _ai_ok:
        st.caption(
            "AI-powered Q&A using the pinned data on this page (Anthropic Claude). "
            "For broader questions across all teams, use the **Assistant** in the left sidebar."
        )
    else:
        st.caption(
            f"Ask about {team_short} {period_lbl} reps, manager, totals, eligibility, or specific rows. "
            "Running in offline mode — answers come from the pinned data on this page."
        )

    # Team + Period aware suggested prompts. Each set is strictly scoped to ITS team AND
    # the currently-selected period (Q1 vs April vs May etc.) — no cross-team mixing, and
    # the prompts always reference the specific numbers/events from the active pinned file.
    _is_q1 = bool(quarter) and int(quarter) == 1
    _is_apr = bool(month) and int(month) == 4
    _is_may = bool(month) and int(month) == 5
    _is_monthly = bool(month)

    if team_short == "ENT":
        # ENT is Q1-only for now (no monthly ENT plan).
        suggested = [
            f"What is Anthony's total Q1 FY2026 payout, including Q4 pending deals?",
            f"Explain the ENT commission calculation step by step.",
            f"Which Q1 deals were Not Paid? How are they excluded from Payment Received?",
            f"List all paid Q1 ENT deals — which ones contributed to the $254,293.90 Payment Received?",
            f"How many deals did the ENT team close, how many were paid, and how many weren't?",
            f"What are the Q4 pending commissions paid in {period_lbl}?",
            f"What is the Cambrex refund and how is it deducted?",
            f"Summarize the ENT Q1 commission in 3 bullets.",
        ]
    elif team_short == "AM":
        if _is_apr:
            suggested = [
                f"What is the total AM team payout for {period_lbl}, including Joy's manager incentive?",
                "Explain Joy's April manager commission step by step (with Washington Post exception).",
                "What is the Washington Post server-split exception, and how does it flow into Joy's April payout?",
                "Which Q1 FY2026 deals were paid in April — NerdsToGo, Church & Dwight, Artnet — and what commission was earned?",
                "What is the Q1 carry-forward Wipro clawback ($78) and how is it deducted from April?",
                "How many April AM deals were closed, and how many were paid?",
                "Break down April AM deals by owner — Vivin, Joy, Arundhati.",
                "Why is Vivin's April payout $0 — did he leave the organization?",
                f"Summarize the AM {period_lbl} commission in 3 bullets.",
            ]
        elif _is_may:
            suggested = [
                f"What is the total AM team payout for {period_lbl}, including Joy's manager incentive?",
                "Why is Joy earning $0 for May (team is at 15.4%, well below the 60% floor)?",
                "Why is Arundhati Sen not eligible for May commission ($6,229 / $35,000 = 17.80%)?",
                "Deepak joined the AM team mid-May with a $25K quota — is he eligible for any commission?",
                "How many May AM deals were closed, and which one wasn't paid (SLR Consulting - Slack to Teams)?",
                f"Break down {period_lbl} AM deals by owner — Joy, Arundhati, Deepak.",
                "Explain Joy's May performance — 4 deals totaling $9,176 booked, $6,800 received.",
                f"Summarize the AM {period_lbl} commission in 3 bullets.",
            ]
        else:
            # Default: AM Q1 (existing)
            suggested = [
                f"What is the total AM team payout for {period_lbl}, including Joy's manager incentive?",
                "Explain Joy's manager commission calculation step by step (incl. Wipro clawback).",
                "How many deals did the AM team close — how many paid, how many not paid?",
                "Break down deals by owner — how many did Vivin, Joy, Arundhati and Nikita close?",
                "At rep level, how much was received vs how much is pending per rep?",
                "Which AM Q1 deals are Not Paid (or only partially paid)? Show Artnet, Estée Lauder, Church & Dwight, Allied Gold, NerdsToGo.",
                "Explain Vivin Joseph's Q1 commission — base, 10% deduction, and the Vendasta manage deal.",
                "Why is Arundhati Sen not eligible for commission this quarter?",
                "Summarize the AM Q1 commission in 3 bullets.",
            ]
    elif team_short == "Outbound":
        suggested = [
            f"What is the total outbound meeting payout for {period_lbl}?",
            "How are outbound meeting incentives calculated (per region)?",
            "Which regions are eligible for outbound incentives?",
            "What are the rules for No-Show / Meeting Validation?",
            "Show outbound meeting counts by rep.",
            "Summarize the outbound payout structure in 3 bullets.",
        ]
    else:
        # SMB (ENT and AM handled above with team-specific prompts).
        if _is_apr:
            suggested = [
                f"What is the total SMB team payout for {period_lbl}, including Chitradip's manager incentive?",
                "Explain Yogi's April commission — Yieldstreet Q4 pending and rounded total $1,004.",
                "Why did Lawrence Lewis's April eligibility change — did he leave the organization?",
                "What is Chitradip's April manager commission (Willow Wealth $170.68 + Legal Soft $310.60 pending)?",
                "How many April SMB deals were closed, and how many had payment received?",
                f"Break down {period_lbl} SMB deals by owner (Vicky, Yogi, Kritika, Lawrence, Deepak, Rutuja).",
                f"Who is eligible for SMB commission in {period_lbl} (strict 60% floor)?",
                f"Summarize the SMB {period_lbl} commission in 3 bullets.",
            ]
        elif _is_may:
            suggested = [
                f"What is the total SMB team payout for {period_lbl}, including Chitradip's manager incentive?",
                "Explain Kritika's May commission — $1,979.40 gross with the $581.50 clawback → $1,397.90 net.",
                "Explain Yogi's May commission — May base ($513) + April SBG-Corp deal ($319) + Pending Commission ($581.50) = $1,413.50.",
                "Rutuja is at 65% attainment against her $25K quota — what commission does she earn?",
                "Why is Chitradip earning $0 for May (team is at 51.6%, below the 60% floor)?",
                "Who is Lennis Brown (Deactivated User) and how does his ICS Data deal factor into the team?",
                "How many May SMB deals were closed, and which one wasn't paid (Protecht - Box to SharePoint)?",
                f"Break down {period_lbl} SMB deals by owner (Kritika, Yogi, Rutuja, Vicky, Lennis).",
                f"Summarize the SMB {period_lbl} commission in 3 bullets.",
            ]
        elif _is_q1:
            suggested = [
                f"What is the total SMB team payout for {period_lbl}, including Chitradip's manager incentive?",
                "Explain Chitradip's manager commission calculation step by step (incl. all 3 Q4 clawbacks).",
                "How many deals did the SMB team close — paid, not paid, cancelled?",
                "Break down SMB deals by owner — Vicky, Kritika, Lawrence, Yogesh, Deepak, Royston, Kartik.",
                "Explain Lawrence Lewis's 2 manage deals — why only $500 was paid this cycle and $500 is pending.",
                "Why was the Keytel Systems deal cancelled and how does it affect Yogi's achievement?",
                "Who is eligible for SMB commission this Q1, and who is not? Why?",
                "Show me each SMB rep's commission breakdown including INR conversion.",
            ]
        else:
            # Fallback for other SMB periods (Q2, Q3, other months)
            suggested = [
                f"What is the total SMB team payout for {period_lbl}, including Chitradip's manager incentive?",
                f"Explain Chitradip's manager commission calculation for {period_lbl}.",
                f"How many SMB deals were closed in {period_lbl} — paid vs not paid?",
                f"Break down {period_lbl} SMB deals by owner.",
                f"Who is eligible for SMB commission in {period_lbl}?",
                f"Summarize the SMB {period_lbl} commission in 3 bullets.",
            ]

    cols = st.columns(2)
    for i, prompt in enumerate(suggested):
        with cols[i % 2]:
            if st.button(prompt, key=f"{session_key}_sug_{i}", use_container_width=True):
                hist = st.session_state.setdefault(session_key, [])
                hist.append({"role": "user", "content": prompt})
                with st.spinner("Thinking…"):
                    answer = _ask_pinned_with_context(prompt, pinned, hist[:-1])
                hist.append({"role": "assistant", "content": answer})
                st.rerun()

    history = st.session_state.setdefault(session_key, [])
    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_input := st.chat_input(
        f"Ask anything about {team_short} {period_lbl}…",
        key=f"{session_key}_input",
    ):
        history.append({"role": "user", "content": user_input})
        # Render the user message immediately, then stream the answer below.
        with st.chat_message("user"):
            st.markdown(user_input)
        with st.chat_message("assistant"):
            stream_placeholder = st.empty()
            answer = _ask_pinned_with_context(
                user_input, pinned, history[:-1], stream_placeholder=stream_placeholder
            )
        history.append({"role": "assistant", "content": answer})
        st.rerun()

    if history:
        if st.button("Clear chat", key=f"{session_key}_clear"):
            st.session_state[session_key] = []
            st.rerun()


def render_pinned_smb_manager_view(pinned: dict) -> None:
    """Pinned manager incentive view for SMB Q1 2026 (Chitradip): per-rep table, manager metrics, highlighted clawbacks, final payout.

    User-level filtering: SALES_REPs see only their own row/deals. Chitradip (manager)
    and ADMIN see everything. If a SALES_REP is looking at this view and they aren't
    Chitradip, the manager block is hidden (only their personal row/deals shown).
    """
    _current_user = st.session_state.get("user") if hasattr(st, "session_state") else None
    _user_matchers = _get_user_rep_name_matchers(_current_user)
    if _user_matchers:
        pinned = _filter_pinned_for_user(pinned, _current_user)
        st.info(
            f"🔒 **Personal view** — showing only your data for Q{pinned.get('quarter')} FY{pinned.get('year')}. "
            "Team totals are shown for context."
        )
    mgr = pinned.get("manager_incentive") or {}
    if not mgr:
        return

    reps = pinned.get("reps") or []
    year = pinned.get("year")
    quarter = pinned.get("quarter")
    mgr_name = (mgr.get("manager_name") or "").strip() or "Manager"
    try:
        inr_rate = float((pinned.get("commission") or {}).get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0

    st.markdown(f"### Q{quarter} FY{year} · {mgr_name}")
    st.caption("Pinned values — read from `policy/smb_q{quarter}_{year}_fixed.json`. Edit that file to change.".format(quarter=quarter, year=year))

    # ---- Per-rep table (Achieved / Payment received / Target) ----
    rep_html_rows: list[str] = []
    sum_ach = 0.0
    sum_pay = 0.0
    sum_tgt = 0.0
    short_name = {
        "Lawrence Lewis": "Larry",
        "Yogi": "Yogesh Vig",
    }
    for r in reps:
        nm = (r.get("name") or "").strip()
        disp_nm = short_name.get(nm, nm)
        ach = float(r.get("achievement_usd") or 0)
        pay = float(r.get("payment_received_usd") or 0)
        tgt = float(r.get("target_usd") or 0)
        sum_ach += ach
        sum_pay += pay
        sum_tgt += tgt
        rep_html_rows.append(
            f"<tr>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>{html_module.escape(disp_nm)}</td>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${ach:,.0f}</td>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${pay:,.0f}</td>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${tgt:,.0f}</td>"
            f"</tr>"
        )
    rep_html_rows.append(
        f"<tr style='background:#dcfce7;font-weight:500;'>"
        f"<td style='padding:10px 12px;'>Total</td>"
        f"<td style='padding:10px 12px;text-align:right;'>${sum_ach:,.0f}</td>"
        f"<td style='padding:10px 12px;text-align:right;'>${sum_pay:,.0f}</td>"
        f"<td style='padding:10px 12px;text-align:right;'>${sum_tgt:,.0f}</td>"
        f"</tr>"
    )
    rep_table_html = (
        '<div style="width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:18px;">'
        '<table style="border-collapse:collapse;font-size:13px;width:100%;min-width:560px;">'
        '<thead><tr style="background:#fef9c3;">'
        '<th style="text-align:left;padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:500;">Sales Rep</th>'
        '<th style="text-align:right;padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:500;">Achieved ($)</th>'
        '<th style="text-align:right;padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:500;">Payment received</th>'
        '<th style="text-align:right;padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:500;">Target ($)</th>'
        '</tr></thead>'
        f'<tbody>{"".join(rep_html_rows)}</tbody>'
        '</table></div>'
    )
    st.markdown(rep_table_html, unsafe_allow_html=True)

    # ---- Manager metrics ----
    team_ach_pct = float(mgr.get("team_achievement_pct") or 0)
    team_ach_amt = float(mgr.get("team_achievement_usd") or 0)
    team_paid = float(mgr.get("team_payment_received_usd") or 0)
    rate_pct = float(mgr.get("commission_rate_pct") or 0)
    base_amt = float(mgr.get("manager_q1_amount_usd") or 0)
    basis_label = (mgr.get("commission_basis_label") or "Team payment received").strip()

    # Optional Joy-style fields: personal commission + manage deal incentive.
    personal_paid = float(mgr.get("personal_payment_received_usd") or 0)
    personal_rate = float(mgr.get("personal_commission_rate_pct") or 0)
    personal_comm = float(mgr.get("personal_commission_usd") or 0)
    md_value = float(mgr.get("manage_deal_value_usd") or 0)
    md_rate = float(mgr.get("manage_deal_rate_pct") or 0)
    md_comm = float(mgr.get("manage_deal_commission_usd") or 0)
    has_personal = personal_comm > 0
    has_manage_deal = md_comm > 0

    metric_rows: list[tuple[str, str]] = [
        ("Team Achievement %", f"{team_ach_pct:.0f}%"),
        ("Team Achievement Amount", f"${team_ach_amt:,.0f}"),
    ]
    if has_personal:
        # Joy-style: show only her personal payment received and rate (no team payment row).
        metric_rows.extend(
            [
                ("Payment Received", f"${personal_paid:,.0f}"),
                ("Commission Rate", f"{personal_rate:.2f}%"),
                (f"Q{quarter} Commission", f"<strong>${personal_comm:,.0f}</strong>"),
            ]
        )
    else:
        # Chit-style: team payment + team-based rate.
        metric_rows.extend(
            [
                ("Team Payment Received", f"${team_paid:,.0f}"),
                ("Commission Rate", f"{rate_pct:.2f}%"),
            ]
        )

    if has_manage_deal:
        metric_rows.extend(
            [
                ("Manage Deal Value", f"${md_value:,.0f}"),
                ("Manage Deal Rate", f"{md_rate:g}%"),
                ("Manage Deal Incentive", f"<strong>${md_comm:,.0f}</strong>"),
            ]
        )

    metric_rows.append(
        (
            f"{mgr_name} Q{quarter} Amount",
            f"<strong>${base_amt:,.0f}</strong>"
            + (f" &nbsp;<span style='color:#6b7280;font-weight:400;'>(₹{(base_amt * inr_rate):,.0f} @ ₹{inr_rate:g})</span>" if inr_rate else ""),
        )
    )

    metric_html = "".join(
        f"<tr><td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;'>{html_module.escape(lbl)}</td>"
        f"<td style='padding:8px 12px;border-bottom:1px solid #f1f1ef;text-align:right;font-variant-numeric:tabular-nums;'>{val}</td></tr>"
        for lbl, val in metric_rows
    )
    metric_table_html = (
        '<div style="width:100%;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:18px;">'
        '<table style="border-collapse:collapse;font-size:13px;width:100%;">'
        '<thead><tr style="background:#fef9c3;">'
        '<th style="text-align:left;padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:500;">Metric</th>'
        '<th style="text-align:right;padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:500;">Value</th>'
        '</tr></thead>'
        f'<tbody>{metric_html}</tbody>'
        '</table></div>'
    )
    st.markdown(f"#### {mgr_name} commission — base calculation")
    if has_personal and has_manage_deal:
        st.caption(
            f"Base = Personal commission (\\${personal_paid:,.0f} × {personal_rate:.2f}% = \\${personal_comm:,.0f}) "
            f"+ Manage Deal (\\${md_value:,.0f} × {md_rate:g}% = \\${md_comm:,.0f}) = \\${base_amt:,.0f}"
        )
    elif has_personal:
        st.caption(
            f"Base = \\${personal_paid:,.0f} × {personal_rate:.2f}% = \\${personal_comm:,.0f}"
        )
    else:
        st.caption(f"Base = {basis_label} × Commission Rate = \\${team_paid:,.0f} × {rate_pct:.2f}% = \\${base_amt:,.0f}")
    st.markdown(metric_table_html, unsafe_allow_html=True)

    # ---- Pending commissions (positive adjustments from prior periods) ----
    pending = mgr.get("pending_commissions") or []
    try:
        pending_total = float(mgr.get("total_pending_commission_usd") or sum(float(p.get("commission_usd") or 0) for p in pending))
    except (TypeError, ValueError):
        pending_total = sum(float(p.get("commission_usd") or 0) for p in pending)
    if pending:
        st.markdown("#### Pending commissions (prior periods)")
        st.caption("Prior-period deals whose payment was received in this period. Paid regardless of current eligibility.")
        pend_rows_html: list[str] = []
        for p in pending:
            dn = html_module.escape((p.get("deal_name") or "").strip())
            per = html_module.escape((p.get("period") or "").strip())
            amt = float(p.get("deal_amount_usd") or 0)
            rate = float(p.get("rate_pct") or 0)
            comm = float(p.get("commission_usd") or 0)
            note = html_module.escape((p.get("note") or "").strip())
            pend_rows_html.append(
                f"<tr style='background:#ecfdf5;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;'><strong>{dn}</strong><br/>"
                f"<span style='font-size:12px;color:#14532d;'>{note}</span></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;'>{per}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;'>${amt:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;'>{rate:g}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;font-variant-numeric:tabular-nums;color:#14532d;'><strong>+${comm:,.2f}</strong></td>"
                f"</tr>"
            )
        pend_rows_html.append(
            f"<tr style='background:#bbf7d0;font-weight:600;'>"
            f"<td colspan='4' style='padding:10px 12px;'>Total pending commission</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#14532d;'>+${pending_total:,.2f}</td>"
            f"</tr>"
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #86efac;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;min-width:680px;'>"
            "<thead><tr style='background:#a7f3d0;color:#064e3b;'>"
            "<th style='text-align:left;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Deal</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Period</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Deal amount</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Slab rate</th>"
            "<th style='text-align:right;padding:10px 12px;border-bottom:2px solid #6ee7b7;font-weight:600;'>Commission</th>"
            f"</tr></thead><tbody>{''.join(pend_rows_html)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Clawbacks (highlighted) ----
    clawbacks = mgr.get("clawbacks") or []
    total_cb = float(mgr.get("total_clawback_usd") or 0)
    if clawbacks:
        st.markdown("#### Clawbacks")
        st.caption("Deductions applied to the manager's base commission for prior-period adjustments.")
        cb_rows: list[str] = []
        for cb in clawbacks:
            lbl = html_module.escape((cb.get("label") or "").strip())
            note = html_module.escape((cb.get("note") or "").strip())
            try:
                deal_amt = float(cb.get("deal_amount_usd") or 0)
            except (TypeError, ValueError):
                deal_amt = 0.0
            try:
                rate_v = float(cb.get("rate_pct") or 0)
            except (TypeError, ValueError):
                rate_v = 0.0
            try:
                ded = float(cb.get("deduction_usd") or 0)
            except (TypeError, ValueError):
                ded = 0.0
            cb_rows.append(
                f"<tr style='background:#fff1f2;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;'><strong>{lbl}</strong>"
                f"<br/><span style='font-size:12px;color:#7f1d1d;'>{note}</span></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;font-variant-numeric:tabular-nums;'>${deal_amt:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;font-variant-numeric:tabular-nums;'>{rate_v:g}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;font-variant-numeric:tabular-nums;color:#b91c1c;'><strong>−${ded:,.0f}</strong></td>"
                f"</tr>"
            )
        cb_rows.append(
            f"<tr style='background:#fee2e2;font-weight:600;'>"
            f"<td style='padding:10px 12px;' colspan='3'>Total clawback</td>"
            f"<td style='padding:10px 12px;text-align:right;font-variant-numeric:tabular-nums;color:#991b1b;'>−${total_cb:,.0f}</td>"
            f"</tr>"
        )
        cb_table_html = (
            '<div style="width:100%;overflow-x:auto;border:2px solid #fecaca;border-radius:8px;margin-bottom:18px;">'
            '<table style="border-collapse:collapse;font-size:13px;width:100%;min-width:680px;">'
            '<thead><tr style="background:#fecaca;color:#7f1d1d;">'
            '<th style="text-align:left;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;">Clawback reason</th>'
            '<th style="text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;">Received Amount (Q4)</th>'
            '<th style="text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;">Slab rate</th>'
            '<th style="text-align:right;padding:10px 12px;border-bottom:2px solid #fca5a5;font-weight:600;">Deduction</th>'
            '</tr></thead>'
            f'<tbody>{"".join(cb_rows)}</tbody>'
            '</table></div>'
        )
        st.markdown(cb_table_html, unsafe_allow_html=True)

    # ---- Calculation summary (step-by-step) ----
    explicit_final = mgr.get("final_payout_usd")
    if explicit_final is not None:
        final = float(explicit_final)
    else:
        final = float(base_amt + pending_total - total_cb)
    carry_forward = float(mgr.get("carry_forward_usd") or 0)
    q1_gross = float(mgr.get("q1_gross_commission_usd") or (base_amt + pending_total))

    calc_rows = [
        ("Q1 base commission", base_amt, "#1f2937"),
    ]
    if pending_total:
        calc_rows.append((f"+ Pending commissions (prior periods)", pending_total, "#15803d"))
        calc_rows.append((f"= Q1 gross commission", q1_gross, "#1f2937"))
    if total_cb:
        calc_rows.append((f"− Clawback (prior-period adjustment)", -total_cb, "#b91c1c"))
    calc_rows.append((f"= Final Q{quarter} {year} balance", final, "#14532d" if final >= 0 else "#7f1d1d"))

    st.markdown("#### Calculation summary")
    calc_html_rows = "".join(
        f"<tr><td style='padding:8px 14px;border-bottom:1px solid #f1f1ef;'>{html_module.escape(lbl)}</td>"
        f"<td style='padding:8px 14px;border-bottom:1px solid #f1f1ef;text-align:right;font-variant-numeric:tabular-nums;color:{color};font-weight:{ 'bold' if lbl.startswith('=') else '400' };'>{'−' if amt < 0 else '+' if lbl.startswith('+') else ''}${abs(amt):,.2f}</td></tr>"
        for lbl, amt, color in calc_rows
    )
    st.markdown(
        "<div style='width:100%;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:18px;'>"
        "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
        f"<tbody>{calc_html_rows}</tbody></table></div>",
        unsafe_allow_html=True,
    )

    # ---- Final payout banner ----
    st.markdown("#### Final payout")

    # Compute INR equivalent once so both branches (and the Excel export below) can use it.
    final_inr = round(final * inr_rate, 2) if inr_rate else 0.0

    if final < 0:
        # Negative net — no payout this cycle, balance carries forward.
        display_payout = 0.0
        cf_amount = final if carry_forward == 0 else carry_forward
        st.markdown(
            f"<div style='background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:14px 18px;color:#7f1d1d;'>"
            f"<div style='font-size:18px;font-weight:600;'>Total Q{quarter} FY{year} Payout: $0.00</div>"
            f"<div style='font-size:14px;font-weight:500;color:#991b1b;margin-top:4px;'>Net balance: <strong>${final:,.2f}</strong> — carried forward to next compensation cycle.</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    else:
        inr_suffix = (
            f"<div style='font-size:14px;font-weight:500;color:#166534;margin-top:4px;'>≈ ₹{final_inr:,.0f} (INR @ ₹{inr_rate:g}/USD)</div>"
            if inr_rate
            else ""
        )
        st.markdown(
            f"<div style='background:#bbf7d0;border:1px solid #4ade80;border-radius:8px;padding:14px 18px;color:#14532d;'>"
            f"<div style='font-size:18px;font-weight:600;'>Total Q{quarter} FY{year} Payout: ${final:,.2f}</div>"
            f"{inr_suffix}"
            f"</div>",
            unsafe_allow_html=True,
        )
    explicit_calc = (mgr.get("calculation_note") or "").strip()
    if explicit_calc:
        # Escape $ so Markdown doesn't treat it as LaTeX math mode.
        st.caption("Calculation: " + explicit_calc.replace("$", "\\$"))
    else:
        st.caption(
            (f"Calculation: \\${base_amt:,.0f} (base) "
             + (f"+ \\${pending_total:,.0f} (pending) " if pending_total else "")
             + f"− \\${total_cb:,.0f} (clawback) = \\${final:,.2f}")
        )

    # ---- Excel export ----
    rep_df = pd.DataFrame(
        [
            {
                "Sales Rep": short_name.get((r.get("name") or "").strip(), (r.get("name") or "").strip()),
                "Achieved ($)": float(r.get("achievement_usd") or 0),
                "Payment received ($)": float(r.get("payment_received_usd") or 0),
                "Target ($)": float(r.get("target_usd") or 0),
            }
            for r in reps
        ]
    )
    metrics_df = pd.DataFrame(
        [
            {"Metric": "Team Achievement %", "Value": f"{team_ach_pct:.0f}%"},
            {"Metric": "Team Achievement Amount", "Value": f"${team_ach_amt:,.0f}"},
            {"Metric": "Payment Received", "Value": f"${team_paid:,.0f}"},
            {"Metric": "Commission Rate", "Value": f"{rate_pct:.2f}%"},
            {"Metric": f"{mgr_name} Q{quarter} Amount (USD)", "Value": f"${base_amt:,.0f}"},
            {"Metric": f"{mgr_name} Q{quarter} Amount (INR)", "Value": f"₹{(base_amt * inr_rate):,.0f}" if inr_rate else "—"},
        ]
    )
    cb_df = pd.DataFrame(
        [
            {
                "Clawback reason": (cb.get("label") or "").strip(),
                "Note": (cb.get("note") or "").strip(),
                "Deal amount ($)": float(cb.get("deal_amount_usd") or 0),
                "Slab rate (%)": float(cb.get("rate_pct") or 0),
                "Deduction ($)": -float(cb.get("deduction_usd") or 0),
            }
            for cb in clawbacks
        ]
        + [{"Clawback reason": "Total clawback", "Note": "", "Deal amount ($)": "", "Slab rate (%)": "", "Deduction ($)": -total_cb}]
    )
    payout_df = pd.DataFrame(
        [
            {"Item": f"{mgr_name} base commission (USD)", "Amount": base_amt},
            {"Item": "Total clawback (USD)", "Amount": -total_cb},
            {"Item": f"Final payout (USD)", "Amount": final},
            {"Item": f"Final payout (INR @ ₹{inr_rate:g})", "Amount": final_inr},
        ]
    )
    xls_bytes = _pinned_export_to_excel_bytes(
        {
            "Sales Reps": rep_df,
            "Manager metrics": metrics_df,
            "Clawbacks": cb_df,
            "Final payout": payout_df,
        }
    )
    st.download_button(
        label="📥 Export Manager commission to Excel",
        data=xls_bytes,
        file_name=f"smb_q{quarter}_{year}_manager_commission_{mgr_name.lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_smb_manager_xlsx",
    )


def render_pinned_smb_team_view(pinned: dict) -> None:
    """Pinned team view for SMB: clean charts + insights based on the pinned numbers (no live database)."""
    reps = pinned.get("reps") or []
    year = pinned.get("year")
    quarter = pinned.get("quarter")
    team_target = float(pinned.get("team_target_usd") or 0)
    team_ach = float(pinned.get("team_achievement_usd") or 0)
    mgr = pinned.get("manager_incentive") or {}
    team_paid = float(mgr.get("team_payment_received_usd") or 0)

    if not reps:
        st.info("No pinned rep data found.")
        return

    short_name = {"Lawrence Lewis": "Larry", "Yogi": "Yogesh Vig"}
    df_reps = pd.DataFrame([
        {
            "Rep": short_name.get((r.get("name") or "").strip(), (r.get("name") or "").strip()),
            "Achieved": float(r.get("achievement_usd") or 0),
            "Paid": float(r.get("payment_received_usd") or 0),
            "Target": float(r.get("target_usd") or 0),
        }
        for r in reps
    ])
    df_reps["Achievement %"] = (df_reps["Achieved"] / df_reps["Target"] * 100.0).round(1)
    df_reps["Collection %"] = ((df_reps["Paid"] / df_reps["Achieved"]).where(df_reps["Achieved"] > 0, 0) * 100.0).round(1)

    # ---- Top: KPI cards ----
    st.markdown(f"### Q{quarter} FY{year} — SMB Team view")
    team_ach_pct = round((team_ach / team_target * 100.0), 1) if team_target else 0.0
    team_collection_pct = round((team_paid / team_ach * 100.0), 1) if team_ach else 0.0
    eligible_count = int((df_reps["Achievement %"] >= 50).sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Team target", f"${team_target:,.0f}")
    k2.metric("Achieved", f"${team_ach:,.0f}", f"{team_ach_pct:.1f}% of target")
    k3.metric("Collected (paid)", f"${team_paid:,.0f}", f"{team_collection_pct:.1f}% of achieved")
    k4.metric("Eligible reps (≥50%)", f"{eligible_count} of {len(df_reps)}")
    top_rep_row = df_reps.loc[df_reps["Achievement %"].idxmax()] if len(df_reps) else None
    if top_rep_row is not None:
        k5.metric("Top performer", str(top_rep_row["Rep"]), f"{top_rep_row['Achievement %']:.0f}%")

    st.markdown("---")

    # ---- Chart 1: Achieved vs Target per rep (grouped bars) ----
    st.markdown("#### Achieved vs Target per rep")
    df_long = df_reps.melt(id_vars=["Rep"], value_vars=["Achieved", "Target"], var_name="Metric", value_name="USD")
    chart1 = (
        alt.Chart(df_long)
        .mark_bar()
        .encode(
            y=alt.Y("Rep:N", sort="-x", title=None),
            x=alt.X("USD:Q", title="USD", axis=alt.Axis(format="$~s")),
            color=alt.Color(
                "Metric:N",
                scale=alt.Scale(domain=["Achieved", "Target"], range=["#1d4ed8", "#cbd5e1"]),
                legend=alt.Legend(orient="top", title=None),
            ),
            yOffset="Metric:N",
            tooltip=["Rep", "Metric", alt.Tooltip("USD:Q", format="$,.0f")],
        )
        .properties(height=320)
    )
    st.altair_chart(chart1, use_container_width=True)

    # ---- Chart 2: Achievement % horizontal bars with tier color ----
    st.markdown("#### Achievement % vs individual quota")
    st.caption("Color shows the commission tier: red = below 50% (not eligible), amber = 50–75%, green = 76%+.")
    def _tier_color(p: float) -> str:
        if p < 50: return "Below 50% (not eligible)"
        if p < 76: return "50–75%"
        if p < 101: return "76–100%"
        return "101%+"

    df_pct = df_reps[["Rep", "Achievement %"]].copy()
    df_pct["Tier"] = df_pct["Achievement %"].apply(_tier_color)
    chart2 = (
        alt.Chart(df_pct)
        .mark_bar(cornerRadius=3)
        .encode(
            y=alt.Y("Rep:N", sort="-x", title=None),
            x=alt.X("Achievement %:Q", title="Achievement %", scale=alt.Scale(domain=[0, max(125, df_pct['Achievement %'].max() + 10)])),
            color=alt.Color(
                "Tier:N",
                scale=alt.Scale(
                    domain=["Below 50% (not eligible)", "50–75%", "76–100%", "101%+"],
                    range=["#ef4444", "#f59e0b", "#10b981", "#1d4ed8"],
                ),
                legend=alt.Legend(orient="top", title=None),
            ),
            tooltip=["Rep", alt.Tooltip("Achievement %:Q", format=".1f"), "Tier"],
        )
        .properties(height=300)
    )
    text2 = (
        alt.Chart(df_pct)
        .mark_text(align="left", dx=4, color="#111827", fontWeight=500)
        .encode(y=alt.Y("Rep:N", sort="-x"), x="Achievement %:Q", text=alt.Text("Achievement %:Q", format=".0f"))
    )
    st.altair_chart(chart2 + text2, use_container_width=True)

    # ---- Chart 3: Collection % (Paid ÷ Achieved) ----
    st.markdown("#### Cash collection — Payment received vs Achieved")
    st.caption("How much of each rep's booked revenue has actually been collected (paid).")
    # Pre-sort reps by Achieved descending so the chart's x-axis order matches.
    rep_order = df_reps.sort_values("Achieved", ascending=False)["Rep"].tolist()
    df_coll = df_reps[["Rep", "Achieved", "Paid"]].melt(id_vars="Rep", var_name="Metric", value_name="USD")
    chart3 = (
        alt.Chart(df_coll)
        .mark_bar()
        .encode(
            x=alt.X("Rep:N", sort=rep_order, title=None),
            y=alt.Y("USD:Q", title="USD", axis=alt.Axis(format="$~s")),
            color=alt.Color(
                "Metric:N",
                scale=alt.Scale(domain=["Achieved", "Paid"], range=["#94a3b8", "#0ea5e9"]),
                legend=alt.Legend(orient="top", title=None),
            ),
            xOffset="Metric:N",
            tooltip=["Rep", "Metric", alt.Tooltip("USD:Q", format="$,.0f")],
        )
        .properties(height=300)
    )
    st.altair_chart(chart3, use_container_width=True)

    # ---- Chart 4: Team progress doughnut (Paid / Achieved / Gap-to-target) ----
    st.markdown("#### Team progress to target")
    gap_to_target = max(team_target - team_ach, 0.0)
    receivable = max(team_ach - team_paid, 0.0)
    df_donut = pd.DataFrame(
        [
            {"Bucket": "Paid", "USD": team_paid},
            {"Bucket": "Booked but unpaid", "USD": receivable},
            {"Bucket": "Remaining to target", "USD": gap_to_target},
        ]
    )
    chart4 = (
        alt.Chart(df_donut)
        .mark_arc(innerRadius=70, outerRadius=130)
        .encode(
            theta=alt.Theta("USD:Q", stack=True),
            color=alt.Color(
                "Bucket:N",
                scale=alt.Scale(
                    domain=["Paid", "Booked but unpaid", "Remaining to target"],
                    range=["#10b981", "#f59e0b", "#cbd5e1"],
                ),
                legend=alt.Legend(orient="right", title=None),
            ),
            tooltip=["Bucket", alt.Tooltip("USD:Q", format="$,.0f")],
        )
        .properties(height=300)
    )
    st.altair_chart(chart4, use_container_width=True)

    # ---- Insights summary ----
    st.markdown("#### Insights")
    insight_items: list[str] = []
    insight_items.append(
        f"Team is at **{team_ach_pct:.1f}%** of the ${team_target:,.0f} target — **${team_ach:,.0f}** booked, **${team_paid:,.0f}** collected ({team_collection_pct:.1f}%)."
    )
    not_elig = df_reps[df_reps["Achievement %"] < 50]["Rep"].tolist()
    if not_elig:
        insight_items.append(
            f"**{len(not_elig)} rep(s)** below the 50% eligibility threshold and not earning commission this cycle: {', '.join(not_elig)}."
        )
    over_100 = df_reps[df_reps["Achievement %"] >= 100]["Rep"].tolist()
    if over_100:
        insight_items.append(f"**{len(over_100)} rep(s)** already at or beyond quota: {', '.join(over_100)}.")
    low_collection = df_reps[(df_reps["Achieved"] > 0) & (df_reps["Collection %"] < 50)]["Rep"].tolist()
    if low_collection:
        insight_items.append(
            f"**{len(low_collection)} rep(s)** have less than 50% of booked revenue collected (cash-flow risk): {', '.join(low_collection)}."
        )
    bullets_html = "".join(f"<li style='margin:4px 0;'>{itm}</li>" for itm in insight_items)
    st.markdown(
        f"<ul style='line-height:1.6;color:#1f2937;'>{bullets_html}</ul>",
        unsafe_allow_html=True,
    )


def _render_smb_goal_attainment_table() -> None:
    """SMB Goal Attainment: bullet charts vs goal (team + each rep)."""
    from commission_policy import SMB_TEAM_NAME, TEAM_QUARTERLY_TARGETS_USD, smb_individual_quota_usd_for_rep

    inv = get_rep_incentives() or []
    users = get_all_users_with_teams()

    start_d, end_d, range_label, year, quarter = _close_date_picker("ga_smb", default_preset="Last quarter")
    data_source = "Finalized incentives (database)"

    # If the selected range is a single calendar month, look for a pinned monthly file first.
    _month_of_range = _is_single_month_range(start_d, end_d)
    if _month_of_range:
        monthly_pinned = _load_pinned_monthly("smb", int(start_d.year), _month_of_range)
        if monthly_pinned is not None:
            render_pinned_monthly_team_view(monthly_pinned, "SMB")
            return

    period_labels = _quarter_month_labels(year, quarter)
    quota_period = period_labels[0]
    team_goal = float(TEAM_QUARTERLY_TARGETS_USD.get(SMB_TEAM_NAME, 0) or 0)

    pinned = _load_pinned_smb_quarter(int(year), int(quarter))
    if pinned is not None:
        # ---- User-level filter: everyone except ADMIN sees only their own data ----
        _current_user = st.session_state.get("user") if hasattr(st, "session_state") else None
        _user_role = (_current_user or {}).get("role") or "unknown"
        _user_name = (_current_user or {}).get("full_name") or (_current_user or {}).get("email") or "unknown"
        _user_matchers = _get_user_rep_name_matchers(_current_user)
        if _user_matchers:
            pinned = _filter_pinned_for_user(pinned, _current_user)
            st.info(
                f"🔒 **Personal view** — showing only your data for Q{quarter} FY{year} "
                f"(logged in as **{_user_name}**, role: {_user_role}). Team totals shown for context."
            )
        elif _user_role.upper() != "ADMIN":
            st.warning(
                f"⚠️ Filter not applied — logged in as **{_user_name}** (role: {_user_role}). "
                "If you expected to see only your own data, please contact your administrator to set your `full_name` on your user record."
            )
        rep_rows: list[dict] = []
        for r in (pinned.get("reps") or []):
            try:
                tgt_p = float(r.get("target_usd") or 0)
                ach_p = float(r.get("achievement_usd") or 0)
            except (TypeError, ValueError):
                tgt_p, ach_p = 0.0, 0.0
            nm = (r.get("name") or "").strip() or "—"
            rep_rows.append({"name": nm, "target": tgt_p, "attained": ach_p})

        team_target_p = float(pinned.get("team_target_usd") or team_goal or 0)
        _team_ach_override = pinned.get("team_achievement_usd")
        if _team_ach_override is not None:
            try:
                team_attained_p = float(_team_ach_override)
            except (TypeError, ValueError):
                team_attained_p = sum(r["attained"] for r in rep_rows)
        else:
            team_attained_p = sum(r["attained"] for r in rep_rows)

        parts = [
            '<div class="goal-attainment-wrap">',
            '<div class="ga-bullet-section">',
            '<p class="ga-bullet-section-title">Performance vs goal</p>',
            _bullet_chart_row_html(
                "SMB Team",
                f"Q{quarter} {year} total",
                team_attained_p,
                team_target_p,
            ),
        ]
        for r in rep_rows:
            av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar(r["name"]))}</div>'
            parts.append(
                _bullet_chart_row_html(
                    r["name"],
                    "Individual quota",
                    r["attained"],
                    r["target"],
                    initials_html=av,
                )
            )
        parts.extend(["</div>", "</div>"])
        st.markdown("".join(parts), unsafe_allow_html=True)

        # ---- Deal-count strip + deal-owner / per-owner breakdown (mirrors AM Q1 page) ----
        paid_deals_smb = pinned.get("paid_deals") or []
        unpaid_deals_smb = pinned.get("unpaid_deals_overrides") or []
        deals_summary_smb = pinned.get("deals_summary") or {}
        if paid_deals_smb or unpaid_deals_smb:
            # Separate cancelled deals — they were closed but no longer count in achievement.
            cancelled_deals_smb = [d for d in unpaid_deals_smb if d.get("is_cancelled")]
            non_cancelled_unpaid_smb = [d for d in unpaid_deals_smb if not d.get("is_cancelled") and not d.get("is_partial_remainder")]
            paid_count_smb = len(paid_deals_smb)
            unpaid_count_smb = len(non_cancelled_unpaid_smb)
            cancelled_count_smb = len(cancelled_deals_smb)
            # Total deals = effective closed count (excludes cancelled deals — they no longer count).
            total_closed_smb = int(deals_summary_smb.get("total_deals_closed") or (paid_count_smb + unpaid_count_smb))
            paid_sum_smb = sum(float(d.get("amount_usd") or 0) for d in paid_deals_smb)
            unpaid_sum_smb = sum(float(d.get("amount_usd") or 0) for d in non_cancelled_unpaid_smb)
            cancelled_sum_smb = sum(float(d.get("amount_usd") or 0) for d in cancelled_deals_smb)
            # Combined value = booked amount AFTER cancellations (= official team_achievement_usd)
            official_team_ach = float(pinned.get("team_achievement_usd") or (paid_sum_smb + unpaid_sum_smb))
            combined_value_smb = official_team_ach

            st.markdown(
                "<div style='display:grid;grid-template-columns:repeat(3, 1fr);gap:12px;margin-top:14px;'>"
                # Card 1 — Total closed
                "<div style='background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:14px 16px;'>"
                "<div style='font-size:12px;font-weight:600;color:#3730a3;text-transform:uppercase;letter-spacing:0.5px;'>📦 Total deals closed</div>"
                f"<div style='font-size:32px;font-weight:700;color:#1e3a8a;margin-top:4px;'>{total_closed_smb}</div>"
                f"<div style='font-size:12px;color:#475569;margin-top:2px;'>Combined value: ${combined_value_smb:,.2f}"
                + (f" <span style='color:#991b1b;'>(after ${cancelled_sum_smb:,.2f} cancellation)</span>" if cancelled_sum_smb else "")
                + "</div>"
                "</div>"
                # Card 2 — Payment received
                "<div style='background:#ecfdf5;border:1px solid #86efac;border-radius:10px;padding:14px 16px;'>"
                "<div style='font-size:12px;font-weight:600;color:#166534;text-transform:uppercase;letter-spacing:0.5px;'>✅ Payment received</div>"
                f"<div style='font-size:32px;font-weight:700;color:#14532d;margin-top:4px;'>{paid_count_smb} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
                f"<div style='font-size:12px;color:#166534;margin-top:2px;'>Total: ${paid_sum_smb:,.2f}</div>"
                "</div>"
                # Card 3 — Payment NOT received
                "<div style='background:#fef2f2;border:1px solid #fca5a5;border-radius:10px;padding:14px 16px;'>"
                "<div style='font-size:12px;font-weight:600;color:#991b1b;text-transform:uppercase;letter-spacing:0.5px;'>❌ Payment not received</div>"
                f"<div style='font-size:32px;font-weight:700;color:#7f1d1d;margin-top:4px;'>{unpaid_count_smb} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
                f"<div style='font-size:12px;color:#991b1b;margin-top:2px;'>Excluded value: ${unpaid_sum_smb:,.2f}</div>"
                "</div>"
                "</div>",
                unsafe_allow_html=True,
            )

            # Cancellation banner — explains the achievement reduction.
            if cancelled_deals_smb:
                cancellation_rows = []
                for d in cancelled_deals_smb:
                    dn = html_module.escape((d.get("deal_name") or "").strip())
                    owner = html_module.escape((d.get("deal_owner") or "").strip())
                    cd = html_module.escape((d.get("close_date") or "").strip())
                    amt = float(d.get("amount_usd") or 0)
                    reason = html_module.escape((d.get("reason") or "").strip())
                    cancellation_rows.append(
                        f"<li style='margin:6px 0;'><strong>{dn}</strong> — Rep: {owner} · Amount: <strong>${amt:,.2f}</strong> · Closed Won {cd}"
                        + (f"<br/><span style='font-size:12px;color:#78350f;'>{reason}</span>" if reason else "")
                        + "</li>"
                    )
                pre_cancel_total = official_team_ach + cancelled_sum_smb
                st.markdown(
                    "<div style='margin:14px 0 10px 0;padding:14px 18px;background:#fef3c7;border:2px solid #f59e0b;border-radius:10px;'>"
                    "<div style='font-size:13px;font-weight:700;color:#78350f;text-transform:uppercase;letter-spacing:0.5px;'>⚠️ Cancellation note</div>"
                    f"<div style='font-size:13px;color:#78350f;margin-top:6px;'><strong>{cancelled_count_smb} deal(s) subsequently cancelled — reducing team achievement from ${pre_cancel_total:,.2f} to ${official_team_ach:,.2f}.</strong></div>"
                    f"<ul style='margin:8px 0 0 0;padding-left:22px;color:#78350f;font-size:13px;'>{''.join(cancellation_rows)}</ul>"
                    "</div>",
                    unsafe_allow_html=True,
                )

            # ---- Deal owner breakdown table ----
            # Cancelled deals: don't count in Booked Amount or Pending. Track separately.
            smb_owners: dict[str, dict] = {}
            for d in paid_deals_smb:
                o = smb_owners.setdefault(d.get("deal_owner") or "—", {
                    "paid_n": 0, "unpaid_n": 0, "cancelled_n": 0,
                    "booked": 0.0, "paid_usd": 0.0, "pending_usd": 0.0, "cancelled_usd": 0.0,
                })
                amt = float(d.get("amount_usd") or 0)
                o["paid_n"] += 1
                o["paid_usd"] += amt
                o["booked"] += amt
            for d in unpaid_deals_smb:
                o = smb_owners.setdefault(d.get("deal_owner") or "—", {
                    "paid_n": 0, "unpaid_n": 0, "cancelled_n": 0,
                    "booked": 0.0, "paid_usd": 0.0, "pending_usd": 0.0, "cancelled_usd": 0.0,
                })
                amt = float(d.get("amount_usd") or 0)
                if d.get("is_cancelled"):
                    o["cancelled_n"] += 1
                    o["cancelled_usd"] += amt
                    # Do NOT add to booked or pending — cancellation removes from achievement
                else:
                    o["unpaid_n"] += 1
                    o["booked"] += amt
                    o["pending_usd"] += amt

            # ---- Build owner → official rep figure lookup -------------------------
            # The breakdown table uses official rep totals from `reps[]` (rounded to whole
            # dollar) for Booked Amount and Payment Received columns. This makes individual
            # rows + total reconcile to the same numbers the rest of the page shows (e.g.
            # Lawrence Lewis $73,236 from rounding $73,235.50, team total $425,934).
            # Pending stays as the deal-level sum because that's what the user verifies
            # against deal-by-deal.
            _smb_rep_lookup: dict[str, dict] = {}
            for r in (pinned.get("reps") or []):
                _smb_rep_lookup[(r.get("name") or "").strip().lower()] = r
            # Aliases — deal_owner names that differ from rep names.
            _smb_owner_alias = {
                "yogesh vig": "yogi",
            }
            def _smb_normalize_owner(owner_raw: str) -> str:
                """Strip '(Deactivated User)' and map aliases (Yogesh Vig → Yogi)."""
                base = (owner_raw or "").replace(" (Deactivated User)", "").strip().lower()
                return _smb_owner_alias.get(base, base)

            st.markdown("#### Deal owner breakdown")
            st.caption("Per-rep summary of SMB deals closed in Q1 — Total deals, Payment received vs not received, Booked Amount, $ received and $ pending. Booked Amount and $ Payment received use the official rep totals from the JSON (rounded to whole dollars).")
            smb_owner_rows: list[str] = []
            s_total = s_paid = s_unpaid = 0
            s_booked_official = s_paid_official = 0.0
            s_pending_usd = 0.0
            for owner, o in sorted(smb_owners.items(), key=lambda kv: (-(kv[1]["paid_n"] + kv[1]["unpaid_n"]), kv[0])):
                tot = o["paid_n"] + o["unpaid_n"]
                s_total += tot
                s_paid += o["paid_n"]
                s_unpaid += o["unpaid_n"]
                s_pending_usd += o["pending_usd"]
                # Resolve official rep figure for this owner if available
                rep_entry = _smb_rep_lookup.get(_smb_normalize_owner(owner))
                if rep_entry:
                    booked_display = round(float(rep_entry.get("achievement_usd") or 0))
                    paid_display = float(rep_entry.get("payment_received_usd") or 0)
                else:
                    booked_display = o["booked"]
                    paid_display = o["paid_usd"]
                s_booked_official += booked_display
                s_paid_official += paid_display
                smb_owner_rows.append(
                    "<tr>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;'><strong>{html_module.escape(owner)}</strong></td>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{tot}</td>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#166534;'><strong>{o['paid_n']}</strong></td>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#991b1b;'><strong>{o['unpaid_n']}</strong></td>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${booked_display:,.2f}</td>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#166534;'>${paid_display:,.2f}</td>"
                    f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;color:#991b1b;'>${o['pending_usd']:,.2f}</td>"
                    "</tr>"
                )
            # Total row uses official team_achievement_usd and team_payment_received_usd
            _official_team_booked = float(pinned.get("team_achievement_usd") or s_booked_official)
            _official_team_paid = float((pinned.get("manager_incentive") or {}).get("team_payment_received_usd") or s_paid_official)
            smb_owner_rows.append(
                "<tr style='background:#fef9c3;font-weight:700;'>"
                f"<td style='padding:10px 12px;'>Total (SMB team)</td>"
                f"<td style='padding:10px 12px;text-align:right;'>{s_total}</td>"
                f"<td style='padding:10px 12px;text-align:right;color:#166534;'>{s_paid}</td>"
                f"<td style='padding:10px 12px;text-align:right;color:#991b1b;'>{s_unpaid}</td>"
                f"<td style='padding:10px 12px;text-align:right;'>${_official_team_booked:,.2f}</td>"
                f"<td style='padding:10px 12px;text-align:right;color:#14532d;'>${_official_team_paid:,.2f}</td>"
                f"<td style='padding:10px 12px;text-align:right;color:#7f1d1d;'>${s_pending_usd:,.2f}</td>"
                "</tr>"
            )
            st.markdown(
                "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-top:6px;margin-bottom:18px;'>"
                "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
                "<thead><tr style='background:#e0f2fe;color:#075985;'>"
                "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Deal Owner</th>"
                "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Total deals</th>"
                "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Payment received deals</th>"
                "<th style='text-align:right;padding:10px 12px;font-weight:600;'># Payment not received</th>"
                "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Booked Amount</th>"
                "<th style='text-align:right;padding:10px 12px;font-weight:600;'>$ Payment received</th>"
                "<th style='text-align:right;padding:10px 12px;font-weight:600;'>$ Payment Pending</th>"
                f"</tr></thead><tbody>{''.join(smb_owner_rows)}</tbody></table></div>",
                unsafe_allow_html=True,
            )

            # ---- Per-owner expander with Excel export ----
            with st.expander("📋 See every SMB deal grouped by owner", expanded=False):
                smb_owner_to_paid: dict[str, list[dict]] = {}
                smb_owner_to_unpaid: dict[str, list[dict]] = {}
                for d in paid_deals_smb:
                    smb_owner_to_paid.setdefault(d.get("deal_owner") or "—", []).append(d)
                for d in unpaid_deals_smb:
                    smb_owner_to_unpaid.setdefault(d.get("deal_owner") or "—", []).append(d)

                # Excel export — all rows with Amount + Payment Receipt
                import pandas as _pd
                _smb_xls_rows: list[dict] = []
                for owner in sorted(set(list(smb_owner_to_paid.keys()) + list(smb_owner_to_unpaid.keys()))):
                    for d in smb_owner_to_paid.get(owner, []):
                        _smb_xls_rows.append({
                            "Deal Owner": owner,
                            "Deal": d.get("deal_name") or "",
                            "Close date": d.get("close_date") or "",
                            "Status": d.get("payment_status") or "Paid",
                            "Amount (USD)": float(d.get("amount_usd") or 0),
                            "Payment Receipt Amount (USD)": float(d.get("amount_usd") or 0),
                            "Pending (USD)": 0.0,
                        })
                    for d in smb_owner_to_unpaid.get(owner, []):
                        _smb_xls_rows.append({
                            "Deal Owner": owner,
                            "Deal": d.get("deal_name") or "",
                            "Close date": d.get("close_date") or "",
                            "Status": d.get("payment_status") or "Not Paid",
                            "Amount (USD)": float(d.get("amount_usd") or 0),
                            "Payment Receipt Amount (USD)": 0.0,
                            "Pending (USD)": float(d.get("amount_usd") or 0),
                        })
                try:
                    _smb_xls = _pinned_export_to_excel_bytes({"All SMB deals by owner": _pd.DataFrame(_smb_xls_rows)})
                    st.download_button(
                        label="📥 Export all SMB deals grouped by owner (Excel)",
                        data=_smb_xls,
                        file_name=f"smb_q{quarter}_{year}_deals_by_owner.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_smb_deals_by_owner_xlsx",
                    )
                except Exception:
                    pass

                grand_amount = 0.0
                grand_paid = 0.0
                grand_n = 0
                for owner in sorted(set(list(smb_owner_to_paid.keys()) + list(smb_owner_to_unpaid.keys()))):
                    p_list = smb_owner_to_paid.get(owner, [])
                    u_list = smb_owner_to_unpaid.get(owner, [])
                    # Split unpaid into "cancelled" vs "still pending"
                    u_non_cancelled = [d for d in u_list if not d.get("is_cancelled")]
                    u_cancelled = [d for d in u_list if d.get("is_cancelled")]
                    p_sum = sum(float(d.get("amount_usd") or 0) for d in p_list)
                    u_sum = sum(float(d.get("amount_usd") or 0) for d in u_non_cancelled)
                    cancelled_sum_o = sum(float(d.get("amount_usd") or 0) for d in u_cancelled)
                    # Owner Booked Amount excludes cancelled deals (they no longer count in achievement)
                    owner_booked = p_sum + u_sum
                    # Effective deal count: paid + non-cancelled unpaid (cancelled deals don't count)
                    deal_n_effective = len(p_list) + len(u_non_cancelled)
                    deal_n_with_cancelled = len(p_list) + len(u_list)  # used for headline only
                    grand_amount += owner_booked
                    grand_paid += p_sum
                    grand_n += deal_n_effective
                    received_chip = f"<span style='color:#166534;font-size:13px;font-weight:500;'>· ${p_sum:,.2f} received</span>"
                    not_received_chip = (
                        f" <span style='color:#991b1b;font-size:13px;font-weight:500;'>· ${u_sum:,.2f} not received</span>"
                        if u_sum > 0 else ""
                    )
                    cancelled_chip = (
                        f" <span style='color:#78350f;font-size:13px;font-weight:500;'>· ${cancelled_sum_o:,.2f} cancelled</span>"
                        if cancelled_sum_o > 0 else ""
                    )
                    st.markdown(
                        f"<div style='margin-top:14px;margin-bottom:6px;font-size:18px;font-weight:700;color:#0f172a;'>"
                        f"{html_module.escape(owner)} — {deal_n_effective} deal(s) {received_chip}{not_received_chip}{cancelled_chip}"
                        "</div>",
                        unsafe_allow_html=True,
                    )
                    rows_h: list[str] = []
                    for d in p_list:
                        dn = html_module.escape((d.get("deal_name") or "").strip())
                        amt = float(d.get("amount_usd") or 0)
                        cd = html_module.escape((d.get("close_date") or "").strip())
                        rows_h.append(
                            f"<tr style='background:#f0fdf4;'>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;'><strong>{dn}</strong></td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;'>{cd}</td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;color:#166534;'>Paid</td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;'><strong>${amt:,.2f}</strong></td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;color:#166534;'><strong>${amt:,.2f}</strong></td>"
                            "</tr>"
                        )
                    for d in u_list:
                        dn = html_module.escape((d.get("deal_name") or "").strip())
                        amt = float(d.get("amount_usd") or 0)
                        cd = html_module.escape((d.get("close_date") or "").strip())
                        status = html_module.escape((d.get("payment_status") or "Not Paid").strip())
                        is_cancelled_d = bool(d.get("is_cancelled"))
                        if is_cancelled_d:
                            bg, border, badge_color = "#fef3c7", "#fde68a", "#78350f"
                            amount_display = f"<s>${amt:,.2f}</s>"
                            paid_display = "$0.00 (cancelled)"
                        else:
                            bg, border, badge_color = "#fef2f2", "#fecaca", "#b91c1c"
                            amount_display = f"${amt:,.2f}"
                            paid_display = "$0.00"
                        rows_h.append(
                            f"<tr style='background:{bg};'>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid {border};'><strong>{dn}</strong></td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;'>{cd}</td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;color:{badge_color};'>{status}</td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;'><strong>{amount_display}</strong></td>"
                            f"<td style='padding:8px 12px;border-bottom:1px solid {border};text-align:right;color:{badge_color};'><strong>{paid_display}</strong></td>"
                            "</tr>"
                        )
                    # Owner total row (Booked Amount excludes cancelled deals)
                    rows_h.append(
                        "<tr style='background:#fef9c3;font-weight:700;'>"
                        f"<td style='padding:10px 12px;' colspan='2'>Total — {html_module.escape(owner)} ({deal_n_effective} deal{'s' if deal_n_effective != 1 else ''})"
                        + (f" <span style='color:#78350f;font-size:11px;font-weight:500;'>(excludes ${cancelled_sum_o:,.2f} cancelled)</span>" if cancelled_sum_o else "")
                        + "</td>"
                        f"<td style='padding:10px 12px;text-align:right;'></td>"
                        f"<td style='padding:10px 12px;text-align:right;'>${owner_booked:,.2f}</td>"
                        f"<td style='padding:10px 12px;text-align:right;color:#166534;'>${p_sum:,.2f}</td>"
                        "</tr>"
                    )
                    st.markdown(
                        "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:14px;'>"
                        "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
                        "<thead><tr style='background:#f1f5f9;color:#334155;'>"
                        "<th style='text-align:left;padding:8px 12px;font-weight:600;'>Deal</th>"
                        "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Close date</th>"
                        "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Status</th>"
                        "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Amount</th>"
                        "<th style='text-align:right;padding:8px 12px;font-weight:600;'>Payment Receipt Amount</th>"
                        f"</tr></thead><tbody>{''.join(rows_h)}</tbody></table></div>",
                        unsafe_allow_html=True,
                    )

                # Grand total — use official figures from JSON where available so the
                # headline reconciles cleanly with the rep totals (avoids $1 deal-level rounding).
                official_total_booked = float(pinned.get("team_achievement_usd") or grand_amount)
                official_total_paid = float((pinned.get("manager_incentive") or {}).get("team_payment_received_usd") or grand_paid)
                # Pending sums directly from the deal-level data (no double-rounding).
                grand_pending = sum(
                    float(d.get("amount_usd") or 0)
                    for d in unpaid_deals_smb
                    if not d.get("is_cancelled")
                )
                st.markdown(
                    "<div style='margin-top:6px;padding:14px 16px;background:#fef3c7;border:2px solid #f59e0b;border-radius:10px;'>"
                    "<div style='font-size:13px;font-weight:700;color:#78350f;text-transform:uppercase;letter-spacing:0.5px;'>🏷️ Grand total — SMB team (all owners)</div>"
                    "<div style='display:grid;grid-template-columns:repeat(4, 1fr);gap:12px;margin-top:8px;'>"
                    f"<div><div style='font-size:11px;color:#92400e;'>Total deals</div><div style='font-size:20px;font-weight:700;color:#78350f;'>{grand_n}</div></div>"
                    f"<div><div style='font-size:11px;color:#92400e;'>Total Amount</div><div style='font-size:20px;font-weight:700;color:#78350f;'>${official_total_booked:,.2f}</div></div>"
                    f"<div><div style='font-size:11px;color:#92400e;'>Payment Receipt</div><div style='font-size:20px;font-weight:700;color:#166534;'>${official_total_paid:,.2f}</div></div>"
                    f"<div><div style='font-size:11px;color:#92400e;'>Pending</div><div style='font-size:20px;font-weight:700;color:#991b1b;'>${grand_pending:,.2f}</div></div>"
                    "</div>"
                    "</div>",
                    unsafe_allow_html=True,
                )

            st.markdown("---")

        comm_block = pinned.get("commission") or {}
        if comm_block:
            from commission_policy import (
                SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
                smb_commission_pct_from_quota_achievement,
            )
            try:
                ded_pct = float(comm_block.get("team_deduction_pct") or 0)
            except (TypeError, ValueError):
                ded_pct = 0.0
            ded_reason = (comm_block.get("team_deduction_reason") or "").strip()
            elig_text = (comm_block.get("eligibility_text") or "").strip()
            elig_min_pct = float(comm_block.get("eligibility_min_pct") or SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT or 50)
            try:
                inr_rate = float(comm_block.get("exchange_rate_inr_per_usd") or 0)
            except (TypeError, ValueError):
                inr_rate = 0.0

            comm_rows: list[dict] = []
            base_total = 0.0
            deduction_total = 0.0
            manage_total = 0.0
            payout_total = 0.0

            for c in (comm_block.get("reps") or []):
                nm = (c.get("name") or "").strip() or "—"
                grp = (c.get("group") or c.get("compensation_group") or "").strip()
                try:
                    quota = float(c.get("quota_usd") or 0)
                except (TypeError, ValueError):
                    quota = 0.0
                try:
                    rev = float(c.get("revenue_achieved_usd") or 0)
                except (TypeError, ValueError):
                    rev = 0.0
                try:
                    achv_pct_v = float(c.get("eligible_pct") or 0)
                except (TypeError, ValueError):
                    achv_pct_v = 0.0
                try:
                    base = float(c.get("base_commission_usd") or 0)
                except (TypeError, ValueError):
                    base = 0.0
                try:
                    manage_deal_v = float(c.get("manage_deal_usd") or 0)
                except (TypeError, ValueError):
                    manage_deal_v = 0.0
                try:
                    manage_pay_v = float(c.get("manage_deal_paid_now_usd") or 0)
                except (TypeError, ValueError):
                    manage_pay_v = 0.0
                try:
                    payout = float(c.get("total_payout_usd") or 0)
                except (TypeError, ValueError):
                    payout = 0.0
                payout_inr = round(payout * inr_rate, 2)
                # Slab % comes from policy: tier rate applied for this rep's group + achievement
                try:
                    slab_pct = float(smb_commission_pct_from_quota_achievement(achv_pct_v, grp) or 0)
                except Exception:
                    slab_pct = 0.0
                if achv_pct_v < elig_min_pct:
                    slab_display = "Not eligible"
                elif slab_pct <= 0:
                    slab_display = "—"
                else:
                    # show as integer if whole, else 1 decimal
                    slab_display = (f"{slab_pct:.0f}%" if abs(slab_pct - round(slab_pct)) < 1e-6 else f"{slab_pct:.1f}%")
                deduction = round(base * (ded_pct / 100.0), 2)
                base_total += base
                deduction_total += deduction
                manage_total += manage_deal_v
                payout_total += payout
                comm_rows.append(
                    {
                        "Rep. Name": nm,
                        "Group": grp,
                        "Quota": quota,
                        "Revenue Achieved": rev,
                        "Achieved %": f"{achv_pct_v:.0f}%",
                        "Slab %": slab_display,
                        "Base Compensation": base,
                        f"{int(ded_pct)}% Deduction": deduction,
                        "Manage Deal": manage_deal_v,
                        "Manage Pay": manage_pay_v,
                        "Total Payout": payout,
                        "Exchange Rate": f"₹{inr_rate:g}/USD" if inr_rate else "—",
                        "Total Payout (INR)": payout_inr,
                        "Note": (c.get("note") or "").strip(),
                    }
                )

            st.markdown("---")
            st.markdown("#### Commission summary")
            if elig_text:
                st.caption(f"**Eligibility criteria:** {elig_text}")
            if ded_reason:
                st.caption(f"**Deduction reason:** {ded_reason}")

            usd_cols = {
                "Quota",
                "Revenue Achieved",
                "Base Compensation",
                f"{int(ded_pct)}% Deduction",
                "Manage Deal",
                "Manage Pay",
                "Total Payout",
            }
            inr_cols = {"Total Payout (INR)"}
            # Render as a real HTML table inside an overflow-x:auto wrapper so the
            # Note column is fully readable via a native horizontal scrollbar.
            _columns = [
                "Rep. Name",
                "Group",
                "Quota",
                "Revenue Achieved",
                "Achieved %",
                "Slab %",
                "Base Compensation",
                f"{int(ded_pct)}% Deduction",
                "Manage Deal",
                "Manage Pay",
                "Total Payout",
                "Exchange Rate",
                "Total Payout (INR)",
                "Note",
            ]
            _col_widths = {
                "Rep. Name": "150px",
                "Group": "80px",
                "Quota": "110px",
                "Revenue Achieved": "140px",
                "Achieved %": "90px",
                "Slab %": "100px",
                "Base Compensation": "150px",
                f"{int(ded_pct)}% Deduction": "130px",
                "Manage Deal": "110px",
                "Manage Pay": "110px",
                "Total Payout": "120px",
                "Exchange Rate": "120px",
                "Total Payout (INR)": "150px",
                "Note": "440px",
            }

            def _fmt_cell(col: str, val) -> str:
                if col in usd_cols:
                    try:
                        return f"${float(val):,.2f}"
                    except (TypeError, ValueError):
                        return html_module.escape(str(val))
                if col in inr_cols:
                    try:
                        return f"₹{float(val):,.2f}"
                    except (TypeError, ValueError):
                        return html_module.escape(str(val))
                return html_module.escape(str(val if val is not None else ""))

            head_cells = "".join(
                f'<th style="text-align:left;padding:8px 12px;border-bottom:1px solid #e5e7eb;background:#f5f3ff;font-weight:500;color:#374151;white-space:nowrap;min-width:{_col_widths[c]};">{html_module.escape(c)}</th>'
                for c in _columns
            )
            body_html_parts: list[str] = []
            for row in comm_rows:
                cells = "".join(
                    f'<td style="padding:8px 12px;border-bottom:1px solid #f1f1ef;vertical-align:top;min-width:{_col_widths[c]};">{_fmt_cell(c, row.get(c, ""))}</td>'
                    for c in _columns
                )
                body_html_parts.append(f"<tr>{cells}</tr>")
            table_html = (
                '<div style="width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;">'
                '<table style="border-collapse:collapse;font-size:13px;width:max-content;min-width:100%;">'
                f"<thead><tr>{head_cells}</tr></thead>"
                f'<tbody>{"".join(body_html_parts)}</tbody>'
                "</table></div>"
            )
            st.markdown(table_html, unsafe_allow_html=True)

            payout_total_inr = round(payout_total * inr_rate, 2)
            sc1, sc2, sc3, sc4, sc5 = st.columns(5)
            sc1.metric("Base compensation (sum)", f"${base_total:,.2f}")
            sc2.metric(f"{int(ded_pct)}% deduction (sum)", f"−${deduction_total:,.2f}")
            sc3.metric("Manage deal earned", f"${manage_total:,.2f}")
            sc4.metric("Total payout (USD)", f"${payout_total:,.2f}")
            sc5.metric(f"Total payout (INR @ ₹{inr_rate:g})", f"₹{payout_total_inr:,.2f}")

            # Excel export of the commission summary
            _xls_df = pd.DataFrame(comm_rows)
            _xls_summary = pd.DataFrame(
                [
                    {"Metric": "Base compensation (sum)", "USD": base_total},
                    {"Metric": f"{int(ded_pct)}% deduction (sum)", "USD": -deduction_total},
                    {"Metric": "Manage deal earned", "USD": manage_total},
                    {"Metric": "Total payout (USD)", "USD": payout_total},
                    {"Metric": f"Total payout (INR @ ₹{inr_rate:g})", "USD": payout_total_inr},
                ]
            )
            _xls_bytes = _pinned_export_to_excel_bytes(
                {"Commission summary": _xls_df, "Totals": _xls_summary}
            )
            st.download_button(
                label="📥 Export commission summary to Excel",
                data=_xls_bytes,
                file_name=f"smb_q{quarter}_{year}_commission_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_smb_comm_xlsx",
            )

            render_pinned_chat_assistant(pinned, session_prefix="smb_q1")

        return

    smb_candidates: list[dict] = []
    for u in users:
        if (u.get("team_name") or "").strip() != SMB_TEAM_NAME:
            continue
        if (u.get("compensation_group") or "").strip().upper() == "SMB_CHITRADIP":
            continue
        if u.get("role") != "SALES_REP":
            continue
        smb_candidates.append(u)

    use_hubspot = data_source.startswith("HubSpot")
    token = get_access_token() if use_hubspot else None
    hubspot_plan: dict | None = None
    rev_by_email: dict[str, float] = {}
    hubspot_failed = False

    if use_hubspot and token:
        smb_emails = {(u.get("email") or "").strip().lower() for u in smb_candidates if u.get("email")}
        try:
            mapped, _stats, _ = fetch_and_map_hubspot_deals(
                token,
                allowed_emails=smb_emails if smb_emails else None,
                year=int(year),
                quarter=int(quarter),
            )
            for row in mapped:
                owner = (row.get("deal_owner") or "").strip()
                em = owner.lower() if "@" in owner else ""
                if not em:
                    continue
                try:
                    amt = float(row.get("amount") or 0.0)
                except (TypeError, ValueError):
                    amt = 0.0
                rev_by_email[em] = rev_by_email.get(em, 0.0) + amt
            hubspot_plan = build_hubspot_goal_sync_plan(token, int(year), int(quarter), smb_candidates)
        except Exception as e:
            hubspot_failed = True
            st.error(f"HubSpot request failed: {e}")
    elif use_hubspot and not token:
        st.warning(
            "HubSpot mode needs a Private App token: set `HUBSPOT_ACCESS_TOKEN` in `.env` or `.streamlit/secrets.toml`, "
            "or use **Finalized incentives (database)**."
        )

    rep_rows: list[dict] = []
    for u in smb_candidates:
        name = (u.get("full_name") or "").strip() or "—"
        uid = int(u["user_id"])
        email = u.get("email")
        em_l = (email or "").strip().lower()
        hq = u.get("hubspot_quota_usd")
        try:
            policy_tgt = float(
                smb_individual_quota_usd_for_rep(
                    u.get("compensation_group"),
                    hq,
                    full_name=name,
                    email=email,
                    calculation_period=quota_period,
                )
            )
        except (TypeError, ValueError):
            policy_tgt = 0.0

        if use_hubspot and token and not hubspot_failed and hubspot_plan is not None:
            hs = hubspot_plan.get("by_user_id", {}).get(uid)
            if hs is not None:
                tgt = float(hs)
            else:
                tgt = policy_tgt
        else:
            tgt = policy_tgt

        if tgt <= 0:
            continue

        if use_hubspot and token and not hubspot_failed:
            rev = rev_by_email.get(em_l, 0.0) if em_l else 0.0
        else:
            rev = _sum_smb_rep_revenue_for_periods(inv, uid, period_labels)

        pct = (rev / tgt * 100.0) if tgt > 0 else 0.0
        rep_rows.append({"name": name, "target": tgt, "attained": rev, "pct": pct})

    rep_rows.sort(key=lambda r: r["name"].lower())

    if not rep_rows:
        st.info(
            "No SMB sales reps with individual targets yet. Set **compensation_group** (SMB_A / SMB_B) on users and "
            "configure quotas in **policy/** (`commission_policy.json` and the SMB quota list JSON). "
            "For HubSpot mode, ensure Sales Goals exist for reps or policy quotas apply."
        )
        return

    team_att = sum(r["attained"] for r in rep_rows)

    q_sub = f"Q{quarter} {year} total"
    team_sub = q_sub
    ind_sub = "Individual quota"
    if use_hubspot and token and not hubspot_failed:
        team_sub = f"Team revenue · {q_sub} · HubSpot"
        ind_sub = f"HubSpot Sales Goal · Q{quarter} {year}"

    parts = [
        '<div class="goal-attainment-wrap">',
        '<div class="ga-bullet-section">',
        '<p class="ga-bullet-section-title">Performance vs goal</p>',
        _bullet_chart_row_html(
            "SMB Team",
            team_sub,
            team_att,
            team_goal,
        ),
    ]
    for r in rep_rows:
        av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar(r["name"]))}</div>'
        parts.append(
            _bullet_chart_row_html(
                r["name"],
                ind_sub,
                r["attained"],
                r["target"],
                initials_html=av,
            )
        )
    parts.extend(["</div>", "</div>"])

    st.markdown("".join(parts), unsafe_allow_html=True)

    if use_hubspot and token and not hubspot_failed and hubspot_plan is not None:
        raw_count = int(hubspot_plan.get("raw_goal_target_count") or 0)
        matched_count = len(hubspot_plan.get("matched_rows") or [])
        st.caption(
            f"HubSpot Sales Goals fetched for Q{quarter} {year}: **{raw_count}** total goal slices, "
            f"matched to **{matched_count}** rep(s) by owner email."
        )
        miss = hubspot_plan.get("users_without_goals") or []
        if miss:
            names = ", ".join((m.get("full_name") or m.get("email") or "?") for m in miss[:8])
            extra = f" (+{len(miss) - 8} more)" if len(miss) > 8 else ""
            st.caption(
                f"No HubSpot Sales Goal matched for: {names}{extra}. Individual target uses policy / DB quota for those users."
            )
        unmatched_owners = hubspot_plan.get("hubspot_owners_with_goals_no_user") or []
        if unmatched_owners:
            owner_lines = []
            for o in unmatched_owners[:8]:
                em = o.get("email") or o.get("hubspot_owner_id") or "?"
                amt = o.get("target_usd") or 0
                owner_lines.append(f"{em} (${amt:,.0f})")
            extra2 = f" (+{len(unmatched_owners) - 8} more)" if len(unmatched_owners) > 8 else ""
            st.caption(
                "HubSpot goals exist for owners not in your User Management roster: "
                + ", ".join(owner_lines)
                + extra2
            )


def _render_account_management_goal_attainment() -> None:
    """Account Management: bullet charts vs policy team goal + individual quotas (policy names or HubSpot)."""
    from commission_policy import (
        ACCOUNT_MANAGEMENT_TEAM_NAME,
        AM_QUOTA_ACHIEVEMENT_ENABLED,
        TEAM_QUARTERLY_TARGETS_USD,
        am_individual_quota_usd_for_rep,
    )

    inv = get_rep_incentives() or []
    users = get_all_users_with_teams()

    start_d, end_d, range_label, year, quarter = _close_date_picker("ga_am", default_preset="Last quarter")

    # Monthly pinned data takes precedence when the range is a single month.
    _month_of_range = _is_single_month_range(start_d, end_d)
    if _month_of_range:
        monthly_pinned = _load_pinned_monthly("am", int(start_d.year), _month_of_range)
        if monthly_pinned is not None:
            render_pinned_monthly_team_view(monthly_pinned, "Account Management")
            return

    # Quarterly pinned: when the range matches a calendar quarter and an AM pinned file exists,
    # render the AM pinned quarterly view (mirrors SMB pinned quarter layout).
    _yr_pick = int(start_d.year)
    _qt_pick = _quarter_of_month(int(start_d.month))
    _qs, _qe = _quarter_start_end(_yr_pick, _qt_pick)
    if start_d == _qs and end_d == _qe:
        am_q_pinned = _load_pinned_am_quarter(_yr_pick, _qt_pick)
        if am_q_pinned is not None:
            render_pinned_am_quarterly_view(am_q_pinned)
            return

    period_labels = _quarter_month_labels(year, quarter)
    quota_period = period_labels[0]
    team_goal = float(TEAM_QUARTERLY_TARGETS_USD.get(ACCOUNT_MANAGEMENT_TEAM_NAME, 0) or 0)

    rep_rows: list[dict] = []
    for u in users:
        if (u.get("team_name") or "").strip() != ACCOUNT_MANAGEMENT_TEAM_NAME:
            continue
        if u.get("role") != "SALES_REP":
            continue
        name = (u.get("full_name") or "").strip() or "—"
        uid = int(u["user_id"])
        try:
            hq = u.get("hubspot_quota_usd")
            if AM_QUOTA_ACHIEVEMENT_ENABLED:
                tgt = float(
                    am_individual_quota_usd_for_rep(
                        hq,
                        full_name=name,
                        email=u.get("email"),
                        calculation_period=quota_period,
                    )
                )
            else:
                tgt = float(hq) if hq is not None and float(hq) > 0 else 0.0
        except (TypeError, ValueError):
            tgt = 0.0
        if tgt <= 0:
            continue
        rev = _sum_team_rep_revenue_for_periods(inv, uid, period_labels, ACCOUNT_MANAGEMENT_TEAM_NAME)
        rep_rows.append({"name": name, "target": tgt, "attained": rev})

    rep_rows.sort(key=lambda r: r["name"].lower())

    if not rep_rows:
        st.info(
            "No Account Management sales reps with a resolved **individual quota** yet. When AM quota rules are enabled in "
            "policy, Joy/Vivin/Arundhati use named quotas; otherwise set **hubspot_quota_usd** and **Team** = Account Management."
        )
        return

    team_att = sum(r["attained"] for r in rep_rows)

    q_sub = f"Q{quarter} {year} total"
    parts = [
        '<div class="goal-attainment-wrap">',
        '<div class="ga-bullet-section">',
        '<p class="ga-bullet-section-title">Performance vs goal</p>',
        _bullet_chart_row_html(
            "Account Management team",
            f"Team revenue · {q_sub}",
            team_att,
            team_goal,
        ),
    ]
    for r in rep_rows:
        av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar(r["name"]))}</div>'
        parts.append(
            _bullet_chart_row_html(
                r["name"],
                "Individual quota",
                r["attained"],
                r["target"],
                initials_html=av,
            )
        )
    parts.extend(["</div>", "</div>"])
    st.markdown("".join(parts), unsafe_allow_html=True)


def _render_ent_quarterly_pinned(pinned: dict) -> None:
    """Render the ENT Q1 2026 (or any pinned ENT quarter): bullet bars + commission summary.

    User-level filtering: SALES_REPs see only their own row/deals. ADMIN and SALES_MANAGER see full team.
    """
    _current_user = st.session_state.get("user") if hasattr(st, "session_state") else None
    _user_matchers = _get_user_rep_name_matchers(_current_user)
    if _user_matchers:
        pinned = _filter_pinned_for_user(pinned, _current_user)
        st.info(
            f"🔒 **Personal view** — showing only your data for Q{pinned.get('quarter')} FY{pinned.get('year')}. "
            "Team totals are shown for context."
        )
    year = pinned.get("year")
    quarter = pinned.get("quarter")
    team_target = float(pinned.get("team_target_usd") or 0)
    reps_top = pinned.get("reps") or []
    team_achievement = float(pinned.get("team_achievement_usd") or sum(float(r.get("achievement_usd") or 0) for r in reps_top))
    team_paid = float(pinned.get("team_payment_received_usd") or sum(float(r.get("payment_received_usd") or 0) for r in reps_top))
    team_pct = (team_achievement / team_target * 100.0) if team_target else 0.0
    comm = pinned.get("commission") or {}
    try:
        inr_rate = float(comm.get("exchange_rate_inr_per_usd") or 0)
    except (TypeError, ValueError):
        inr_rate = 0.0

    st.markdown(f"### Q{quarter} {year} — Enterprise Sales Target")
    st.caption(f"Pinned values — read from `policy/ent_q{quarter}_{year}_fixed.json`.")

    # ---- KPI cards ----
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Team target", f"${team_target:,.0f}")
    k2.metric("Achieved (booked)", f"${team_achievement:,.0f}", f"{team_pct:.1f}%")
    k3.metric("Payment received", f"${team_paid:,.0f}")
    k4.metric("Eligibility floor", f"{int(comm.get('eligibility_min_pct') or 0)}%")

    # ---- Deal-count strip (paid + unpaid → total closed) ----
    _paid_deals_top = pinned.get("paid_deals") or []
    _unpaid_deals_top = pinned.get("unpaid_deals_overrides") or []
    _paid_count = len(_paid_deals_top)
    _unpaid_count = len(_unpaid_deals_top)
    _closed_count = _paid_count + _unpaid_count
    _paid_sum_top = sum(float(p.get("amount_usd") or 0) for p in _paid_deals_top)
    _unpaid_sum_top = sum(float(u.get("amount_usd") or 0) for u in _unpaid_deals_top)
    if _closed_count > 0:
        st.markdown(
            "<div style='display:grid;grid-template-columns:repeat(3, 1fr);gap:12px;margin-top:14px;'>"
            # Card 1 — Total closed
            "<div style='background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#3730a3;text-transform:uppercase;letter-spacing:0.5px;'>📦 Total deals closed</div>"
            f"<div style='font-size:32px;font-weight:700;color:#1e3a8a;margin-top:4px;'>{_closed_count}</div>"
            f"<div style='font-size:12px;color:#475569;margin-top:2px;'>Combined value: ${(_paid_sum_top + _unpaid_sum_top):,.2f}</div>"
            "</div>"
            # Card 2 — Payment received
            "<div style='background:#ecfdf5;border:1px solid #86efac;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#166534;text-transform:uppercase;letter-spacing:0.5px;'>✅ Payment received</div>"
            f"<div style='font-size:32px;font-weight:700;color:#14532d;margin-top:4px;'>{_paid_count} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
            f"<div style='font-size:12px;color:#166534;margin-top:2px;'>Total: ${_paid_sum_top:,.2f}</div>"
            "</div>"
            # Card 3 — Payment NOT received
            "<div style='background:#fef2f2;border:1px solid #fca5a5;border-radius:10px;padding:14px 16px;'>"
            "<div style='font-size:12px;font-weight:600;color:#991b1b;text-transform:uppercase;letter-spacing:0.5px;'>❌ Payment not received</div>"
            f"<div style='font-size:32px;font-weight:700;color:#7f1d1d;margin-top:4px;'>{_unpaid_count} <span style='font-size:16px;font-weight:500;'>deal(s)</span></div>"
            f"<div style='font-size:12px;color:#991b1b;margin-top:2px;'>Excluded value: ${_unpaid_sum_top:,.2f}</div>"
            "</div>"
            "</div>",
            unsafe_allow_html=True,
        )

    if (comm.get("eligibility_text") or "").strip():
        st.caption(comm.get("eligibility_text"))
    st.markdown("---")

    # ---- Bullet bars ----
    parts: list[str] = [
        '<div class="goal-attainment-wrap">',
        '<div class="ga-bullet-section">',
        '<p class="ga-bullet-section-title">Performance vs goal</p>',
    ]
    for r in reps_top:
        nm = (r.get("name") or "").strip()
        ach = float(r.get("achievement_usd") or 0)
        tgt = float(r.get("target_usd") or 0)
        av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar(nm))}</div>'
        parts.append(
            _bullet_chart_row_html(nm, f"Individual target · Q{quarter} {year}", ach, tgt, initials_html=av)
        )
    parts.append(
        _bullet_chart_row_html(
            "Enterprise team",
            f"Q{quarter} {year} total",
            team_achievement,
            team_target,
        )
    )
    parts.extend(["</div>", "</div>"])
    st.markdown("".join(parts), unsafe_allow_html=True)

    # ---- Paid deals table ----
    paid_deals = pinned.get("paid_deals") or []
    if paid_deals:
        paid_total = sum(float(p.get("amount_usd") or 0) for p in paid_deals)
        paid_html_rows: list[str] = []
        for p in paid_deals:
            dn = html_module.escape((p.get("deal_name") or "").strip())
            amt = float(p.get("amount_usd") or 0)
            cd = html_module.escape((p.get("close_date") or "").strip())
            owner = html_module.escape((p.get("deal_owner") or "").strip())
            paid_html_rows.append(
                f"<tr style='background:#f0fdf4;'>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;'><strong>{dn}</strong></td>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;'>{owner}</td>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;'>{cd}</td>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #bbf7d0;text-align:right;color:#166534;'><strong>${amt:,.2f}</strong></td>"
                f"</tr>"
            )
        # Total row at the bottom
        paid_html_rows.append(
            f"<tr style='background:#bbf7d0;font-weight:700;'>"
            f"<td style='padding:10px 12px;'>Total Payment Received</td>"
            f"<td style='padding:10px 12px;'></td>"
            f"<td style='padding:10px 12px;'></td>"
            f"<td style='padding:10px 12px;text-align:right;color:#14532d;'>${paid_total:,.2f}</td>"
            f"</tr>"
        )
        st.markdown(
            "<div style='margin:16px 0 8px 0;padding:12px 14px;background:#ecfdf5;border:1px solid #86efac;border-radius:8px;color:#14532d;'>"
            f"<strong>✅ Paid Q1 deals (included in Payment Received):</strong> "
            f"<strong>{len(paid_deals)} deal(s)</strong> totaling "
            f"<strong>${paid_total:,.2f}</strong> — these are the Q1 deals whose payment was received at the time of the Q1 commission calculation."
            "</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #86efac;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            "<thead><tr style='background:#a7f3d0;color:#064e3b;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Deal Name</th>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Deal Owner</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Close date</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Payment Received</th>"
            f"</tr></thead><tbody>{''.join(paid_html_rows)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Unpaid deals notice ----
    unpaid = pinned.get("unpaid_deals_overrides") or []
    if unpaid:
        unpaid_total = sum(float(u.get("amount_usd") or 0) for u in unpaid)
        unpaid_html_rows: list[str] = []
        for u in unpaid:
            dn = html_module.escape((u.get("deal_name") or "").strip())
            amt = float(u.get("amount_usd") or 0)
            cd = html_module.escape((u.get("close_date") or "").strip())
            reason = html_module.escape((u.get("reason") or "").strip())
            unpaid_html_rows.append(
                f"<tr style='background:#fef2f2;'>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;'><strong>{dn}</strong>"
                + (f"<br/><span style='font-size:12px;color:#7f1d1d;'>{reason}</span>" if reason else "")
                + "</td>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;text-align:right;'>{cd}</td>"
                f"<td style='padding:8px 12px;border-bottom:1px solid #fecaca;text-align:right;color:#b91c1c;'><strong>${amt:,.2f}</strong></td>"
                f"</tr>"
            )
        # Total row at the bottom
        unpaid_html_rows.append(
            f"<tr style='background:#fee2e2;font-weight:700;'>"
            f"<td style='padding:10px 12px;'>Total unpaid</td>"
            f"<td style='padding:10px 12px;'></td>"
            f"<td style='padding:10px 12px;text-align:right;color:#991b1b;'>${unpaid_total:,.2f}</td>"
            f"</tr>"
        )
        st.markdown(
            "<div style='margin:16px 0 8px 0;padding:12px 14px;background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;color:#9a3412;'>"
            f"<strong>📌 Unpaid deals (excluded from Payment Received):</strong> "
            f"Total deal value <strong>${unpaid_total:,.2f}</strong>. "
            "Commission is calculated on Payment Received ($453,168.40 − unpaid = $254,293.90)."
            "</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #fecaca;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            "<thead><tr style='background:#fecaca;color:#7f1d1d;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Deal</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Close date</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Amount</th>"
            f"</tr></thead><tbody>{''.join(unpaid_html_rows)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Commission summary (mirrors SMB / AM Q1 layout) ----
    comm_reps = comm.get("reps") or []
    if comm_reps:
        st.markdown("#### Commission summary")
        st.caption("Commission is calculated on **Payment Received** at the ENT slab rate (0–75.99% → 7%).")
        rows_html: list[str] = []
        for c in comm_reps:
            nm = (c.get("name") or "").strip()
            grp = (c.get("group") or "").strip()
            quota = float(c.get("quota_usd") or 0)
            rev = float(c.get("revenue_achieved_usd") or 0)
            paid = float(c.get("payment_received_usd") or 0)
            ach_pct_v = float(c.get("eligible_pct") or 0)
            slab_pct_v = float(c.get("slab_pct") or 0)
            base = float(c.get("base_commission_usd") or 0)
            payout = float(c.get("total_payout_usd") or 0)
            payout_inr = round(payout * inr_rate, 2) if inr_rate else 0
            note = (c.get("note") or "").strip()
            rows_html.append(
                "<tr>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;'><strong>{html_module.escape(nm)}</strong></td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;'>{html_module.escape(grp)}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${quota:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${rev:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${paid:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{ach_pct_v:.0f}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>{slab_pct_v:.0f}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>${base:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;text-align:right;'>"
                f"<strong>${payout:,.2f}</strong>"
                + (f"<br/><span style='font-size:11px;color:#6b7280;'>≈ ₹{payout_inr:,.0f}</span>" if inr_rate else "")
                + "</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #f1f1ef;font-size:12px;color:#475569;min-width:360px;'>{html_module.escape(note)}</td>"
                "</tr>"
            )
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:max-content;min-width:100%;'>"
            "<thead><tr style='background:#f3e8ff;color:#5b21b6;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Rep. Name</th>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Group</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Quota</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Revenue Achieved</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Payment Received</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Achieved %</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Slab %</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Base Commission</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:600;'>Total Payout</th>"
            "<th style='text-align:left;padding:10px 12px;font-weight:600;'>Note</th>"
            f"</tr></thead><tbody>{''.join(rows_html)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Pending commissions (Q4 deals paid in Q1) ----
    mgr_inc = pinned.get("manager_incentive") or {}
    pending = mgr_inc.get("pending_commissions") or []
    if pending:
        pending_total = float(mgr_inc.get("total_pending_commission_usd") or sum(float(p.get("commission_usd") or 0) for p in pending))
        rows_h: list[str] = []
        for p in pending:
            dn = html_module.escape((p.get("deal_name") or "").strip())
            per = html_module.escape((p.get("period") or "").strip())
            deal_amt = float(p.get("deal_amount_usd") or 0)
            rate = float(p.get("rate_pct") or 0)
            comm = float(p.get("commission_usd") or 0)
            note = html_module.escape((p.get("note") or "").strip())
            rows_h.append(
                "<tr style='background:#ecfdf5;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;'><strong>{dn}</strong>"
                + (f"<br/><span style='font-size:12px;color:#14532d;'>{note}</span>" if note else "")
                + "</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;'>{per}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;'>${deal_amt:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;'>{rate:g}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #a7f3d0;text-align:right;color:#14532d;'><strong>+${comm:,.2f}</strong></td>"
                "</tr>"
            )
        rows_h.append(
            f"<tr style='background:#bbf7d0;font-weight:700;'>"
            f"<td colspan='4' style='padding:10px 12px;'>Total pending commission</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#14532d;'>+${pending_total:,.2f}</td>"
            "</tr>"
        )
        st.markdown("#### Pending commissions (Q4 deals paid in Q1)")
        st.caption("Q4 FY2025 deals whose payment was received in Q1 2026. Commission is paid at the Q4 rate (10%).")
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:1px solid #86efac;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            "<thead><tr style='background:#a7f3d0;color:#064e3b;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:700;'>Deal Name</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Period</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Paid Amount</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Q4 Rate</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Commission</th>"
            f"</tr></thead><tbody>{''.join(rows_h)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Refunds / Clawbacks (Q4 cancelled deals) ----
    clawbacks = mgr_inc.get("clawbacks") or []
    if clawbacks:
        cb_total = float(mgr_inc.get("total_clawback_usd") or sum(float(c.get("deduction_usd") or 0) for c in clawbacks))
        cb_rows: list[str] = []
        for cb in clawbacks:
            lbl = html_module.escape((cb.get("label") or "").strip())
            per = html_module.escape((cb.get("period") or "").strip())
            refund_amt = float(cb.get("deal_amount_usd") or 0)
            rate = float(cb.get("rate_pct") or 0)
            ded = float(cb.get("deduction_usd") or 0)
            note = html_module.escape((cb.get("note") or "").strip())
            cb_rows.append(
                "<tr style='background:#fff1f2;'>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;'><strong>{lbl}</strong>"
                + (f"<br/><span style='font-size:12px;color:#7f1d1d;'>{note}</span>" if note else "")
                + "</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;'>{per}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;'>${refund_amt:,.2f}</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;'>{rate:g}%</td>"
                f"<td style='padding:10px 12px;border-bottom:1px solid #fecdd3;text-align:right;color:#b91c1c;'><strong>${ded:,.2f}</strong></td>"
                "</tr>"
            )
        cb_rows.append(
            f"<tr style='background:#fee2e2;font-weight:700;'>"
            f"<td colspan='4' style='padding:10px 12px;'>Total commission to deduct</td>"
            f"<td style='padding:10px 12px;text-align:right;color:#991b1b;'>${cb_total:,.2f}</td>"
            "</tr>"
        )
        st.markdown("#### Refunds / cancelled Q4 deals (commission to deduct)")
        st.caption("Q4 FY2025 deals that were later cancelled / refunded. The previously-paid commission is clawed back.")
        st.markdown(
            "<div style='width:100%;overflow-x:auto;border:2px solid #fecaca;border-radius:8px;margin-bottom:18px;'>"
            "<table style='border-collapse:collapse;font-size:13px;width:100%;'>"
            "<thead><tr style='background:#fecaca;color:#7f1d1d;'>"
            "<th style='text-align:left;padding:10px 12px;font-weight:700;'>Deal Name</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Period</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Refund Amount</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Q4 Rate</th>"
            "<th style='text-align:right;padding:10px 12px;font-weight:700;'>Commission to deduct</th>"
            f"</tr></thead><tbody>{''.join(cb_rows)}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ---- Final payout banner ----
    final_amt = float(mgr_inc.get("final_payout_usd") or 0)
    if final_amt > 0:
        final_inr = round(final_amt * inr_rate, 2) if inr_rate else 0
        st.markdown("#### Final payout")
        inr_suffix = (
            f"<div style='font-size:14px;font-weight:500;color:#166534;margin-top:4px;'>≈ ₹{final_inr:,.0f} (INR @ ₹{inr_rate:g}/USD)</div>"
            if inr_rate
            else ""
        )
        st.markdown(
            f"<div style='background:#bbf7d0;border:1px solid #4ade80;border-radius:8px;padding:14px 18px;color:#14532d;'>"
            f"<div style='font-size:18px;font-weight:700;'>Total Q{quarter} FY{year} Payout: ${final_amt:,.2f}</div>"
            f"{inr_suffix}"
            f"</div>",
            unsafe_allow_html=True,
        )
        calc_note = (mgr_inc.get("calculation_note") or "").strip()
        if calc_note:
            st.caption("Calculation: " + calc_note.replace("$", "\\$"))

    # ---- Inline chat assistant ----
    render_pinned_chat_assistant(pinned, session_prefix="ent_q1")


def _render_enterprise_goal_attainment() -> None:
    """Enterprise: optional Anthony snapshot (policy) + bullet charts vs team goal + **hubspot_quota_usd** per rep."""
    from commission_policy import (
        ENTERPRISE_ANTHONY_PVG_ACHIEVED,
        ENTERPRISE_ANTHONY_PVG_ENABLED,
        ENTERPRISE_ANTHONY_PVG_SUBTITLE,
        ENTERPRISE_ANTHONY_PVG_TARGET,
        ENTERPRISE_TEAM_NAME,
        TEAM_QUARTERLY_TARGETS_USD,
    )

    show_anthony_pvg = bool(
        ENTERPRISE_ANTHONY_PVG_ENABLED and ENTERPRISE_ANTHONY_PVG_TARGET > 0
    )

    inv = get_rep_incentives() or []
    users = get_all_users_with_teams()

    start_d, end_d, range_label, year, quarter = _close_date_picker("ga_ent", default_preset="Last quarter")

    # If the selected range matches a quarter for which ENT has a pinned file, render
    # the bullet bars from that file and skip the live DB aggregation.
    _yr_pick = int(start_d.year)
    _qt_pick = _quarter_of_month(int(start_d.month))
    _qs, _qe = _quarter_start_end(_yr_pick, _qt_pick)
    if start_d == _qs and end_d == _qe:
        ent_q_pinned = _load_pinned_quarter("ent", _yr_pick, _qt_pick)
        if ent_q_pinned is not None:
            _render_ent_quarterly_pinned(ent_q_pinned)
            return

    period_labels = _quarter_month_labels(year, quarter)
    quota_period = period_labels[0]
    team_goal = float(TEAM_QUARTERLY_TARGETS_USD.get(ENTERPRISE_TEAM_NAME, 0) or 0)

    rep_rows: list[dict] = []
    for u in users:
        if (u.get("team_name") or "").strip() != ENTERPRISE_TEAM_NAME:
            continue
        if u.get("role") != "SALES_REP":
            continue
        name = (u.get("full_name") or "").strip() or "—"
        uid = int(u["user_id"])
        try:
            hq = u.get("hubspot_quota_usd")
            tgt = float(hq) if hq is not None and float(hq) > 0 else 0.0
        except (TypeError, ValueError):
            tgt = 0.0
        if tgt <= 0:
            continue
        rev = _sum_team_rep_revenue_for_periods(inv, uid, period_labels, ENTERPRISE_TEAM_NAME)
        rep_rows.append({"name": name, "target": tgt, "attained": rev})

    rep_rows.sort(key=lambda r: r["name"].lower())
    if show_anthony_pvg:
        rep_rows = [
            r
            for r in rep_rows
            if "anthony" not in (r.get("name") or "").lower()
        ]

    if not rep_rows and not show_anthony_pvg:
        st.info(
            "No Enterprise sales reps with **hubspot_quota_usd** set yet. Assign **Team** = Enterprise in User Management "
            "and set individual quotas."
        )
        return

    q_sub = f"Q{quarter} {year} total"
    parts = [
        '<div class="goal-attainment-wrap">',
        '<div class="ga-bullet-section">',
        '<p class="ga-bullet-section-title">Performance vs goal</p>',
    ]
    if show_anthony_pvg:
        av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar("Anthony"))}</div>'
        parts.append(
            _bullet_chart_row_html(
                "Anthony",
                ENTERPRISE_ANTHONY_PVG_SUBTITLE,
                ENTERPRISE_ANTHONY_PVG_ACHIEVED,
                ENTERPRISE_ANTHONY_PVG_TARGET,
                initials_html=av,
            )
        )
    if rep_rows:
        team_att = sum(r["attained"] for r in rep_rows)
        parts.append(
            _bullet_chart_row_html(
                "Enterprise team",
                f"Team revenue · {q_sub}",
                team_att,
                team_goal,
            )
        )
        for r in rep_rows:
            av = f'<div class="ga-avatar" aria-hidden="true">{html_module.escape(_initials_for_avatar(r["name"]))}</div>'
            parts.append(
                _bullet_chart_row_html(
                    r["name"],
                    "Individual quota",
                    r["attained"],
                    r["target"],
                    initials_html=av,
                )
            )
    parts.extend(["</div>", "</div>"])
    st.markdown("".join(parts), unsafe_allow_html=True)


def _prepare_upload_deals_display(df: pd.DataFrame) -> pd.DataFrame:
    """Align paid_amount with payment_status for the table; format close_date for display."""
    out = df.copy()
    if "paid_amount" in out.columns and "payment_status" in out.columns:
        out["paid_amount"] = out.apply(lambda row: effective_paid_amount_from_status(dict(row)), axis=1)
    if "close_date" in out.columns:
        def _fmt_cd(x):
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return "—"
            if hasattr(x, "isoformat"):
                return x.isoformat()[:10]
            return str(x)

        out["close_date"] = out["close_date"].apply(_fmt_cd)
    return out


def _chart_bar_with_labels(df: pd.DataFrame, x_col: str, y_col: str, height: int = 320):
    """Bar chart with data labels on each bar."""
    d = df[[x_col, y_col]].copy()
    d = d.rename(columns={x_col: "x", y_col: "y"})
    d["label"] = d["y"].apply(lambda v: f"{float(v):,.0f}" if pd.notna(v) else "")
    d = d.sort_values("y", ascending=False)
    cat_order = d["x"].astype(str).tolist()
    base = alt.Chart(d).encode(
        x=alt.X("x:N", title=x_col, sort=cat_order),
        y=alt.Y("y:Q", title=y_col),
        color=_chart_categorical_color("x:N", cat_order, legend=None),
    )
    bars = base.mark_bar(size=28)
    text = base.mark_text(dy=-8, align="center", fontSize=11, color="#0d47a1").encode(text="label:N")
    st.altair_chart(bars + text, use_container_width=True, theme=None)


def _chart_line_with_labels(df: pd.DataFrame, x_col: str, y_col: str, height: int = 320):
    """Line chart with data labels at each point."""
    d = df[[x_col, y_col]].copy()
    d = d.rename(columns={x_col: "x", y_col: "y"})
    d["label"] = d["y"].apply(lambda v: f"{float(v):,.0f}" if pd.notna(v) else "")
    cat_order = d["x"].astype(str).tolist()
    line = (
        alt.Chart(d)
        .mark_line(color="#1565c0", strokeWidth=2.5)
        .encode(x=alt.X("x:N", title=x_col), y=alt.Y("y:Q", title=y_col))
    )
    pts = (
        alt.Chart(d)
        .mark_point(size=70, filled=True, stroke="#0d47a1", strokeWidth=1)
        .encode(
            x=alt.X("x:N", title=x_col),
            y=alt.Y("y:Q", title=y_col),
            color=_chart_categorical_color("x:N", cat_order, legend=None),
        )
    )
    text = alt.Chart(d).mark_text(dy=-12, align="center", fontSize=10, color="#051a3a").encode(
        x="x:N", y="y:Q", text="label:N"
    )
    st.altair_chart(line + pts + text, use_container_width=True, theme=None)


def _chart_area_with_labels(df: pd.DataFrame, x_col: str, y_col: str, height: int = 320):
    """Area chart with data labels at top of area."""
    d = df[[x_col, y_col]].copy()
    d = d.rename(columns={x_col: "x", y_col: "y"})
    d["label"] = d["y"].apply(lambda v: f"{float(v):,.0f}" if pd.notna(v) else "")
    x_enc = alt.X("x:N", title=x_col)
    y_enc = alt.Y("y:Q", title=y_col)
    area = (
        alt.Chart(d)
        .mark_area(interpolate="monotone", color="#64b5f6", opacity=0.38)
        .encode(x=x_enc, y=y_enc)
    )
    line = alt.Chart(d).mark_line(color="#0d47a1", strokeWidth=2).encode(x=x_enc, y=y_enc)
    pts = (
        alt.Chart(d)
        .mark_point(size=55, filled=True, color="#1976d2", stroke="#1565c0", strokeWidth=1)
        .encode(x=x_enc, y=y_enc)
    )
    text = alt.Chart(d).mark_text(dy=-8, align="center", fontSize=11, color="#051a3a").encode(
        x="x:N", y="y:Q", text="label:N"
    )
    st.altair_chart(area + line + pts + text, use_container_width=True, theme=None)


def _fiesta_alt_configure(chart: alt.Chart, height: int = 280) -> alt.Chart:
    """Dark navy theme for analytics charts (blue-tinted axes and grid)."""
    return (
        chart.properties(height=height)
        .configure(background="#1a2332")
        .configure_axis(
            labelColor="#e3f2fd",
            titleColor="#90caf9",
            gridColor="#3d5a80",
            domainColor="#5c7cfa",
        )
        .configure_view(strokeWidth=0)
        .configure_legend(labelColor="#e3f2fd", titleColor="#90caf9")
    )


def _render_rep_incentives_fiesta_analytics(
    incentives: list,
    *,
    primary_team_name: str | None = None,
    analytics_title: str | None = None,
    period_key: str = "fiesta_rep_period",
) -> None:
    """
    Dark analytics board: team progress ring, rep bars for incentive $, achievement %, and revenue.
    Uses rep incentive rows from the database (same basis as the Rep Incentives table).
    ``primary_team_name``: which team’s quarterly policy target to use for the ring (default SMB).
    """
    from commission_policy import SMB_TEAM_NAME, TEAM_QUARTERLY_TARGETS_USD

    team_for_goal = (primary_team_name or SMB_TEAM_NAME).strip()
    title_main = analytics_title or (f"{team_for_goal} performance analytics")
    df = _enrich_rep_incentives_display(pd.DataFrame(incentives))
    if df.empty:
        return

    st.markdown(
        f"""
<style>
.fiesta-analytics-wrap {{
  font-family: "Segoe UI", system-ui, sans-serif;
  margin-bottom: 1rem;
}}
.fiesta-analytics-title {{
  font-size: 1.35rem;
  font-weight: 700;
  letter-spacing: 0.06em;
  margin: 0 0 0.25rem 0;
  background: linear-gradient(95deg, #80deea 0%, #b39ddb 40%, #90caf9 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  color: #e3f2fd !important;
}}
.fiesta-analytics-sub {{
  color: #b0bec5 !important;
  font-size: 0.85rem;
  margin: 0 0 1rem 0;
}}
.fiesta-kpi-row {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 1rem; }}
.fiesta-kpi {{
  background: linear-gradient(155deg, #2c3e65 0%, #1a2332 55%, #243044 100%);
  border: 1px solid #5c6bc0;
  border-radius: 12px;
  padding: 14px 18px;
  min-width: 140px;
  flex: 1 1 140px;
  box-shadow: 0 4px 14px rgba(92, 107, 192, 0.25);
}}
.fiesta-kpi:nth-child(1) {{ border-left: 4px solid #42a5f5; }}
.fiesta-kpi:nth-child(2) {{ border-left: 4px solid #ab47bc; }}
.fiesta-kpi:nth-child(3) {{ border-left: 4px solid #26c6da; }}
.fiesta-kpi:nth-child(4) {{ border-left: 4px solid #ff7043; }}
.fiesta-kpi-label {{ color: #b0bec5 !important; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 6px; }}
.fiesta-kpi-val {{
  font-size: 1.35rem;
  font-weight: 700;
  background: linear-gradient(90deg, #64b5f6, #ce93d8, #4dd0e1);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  color: #90caf9 !important;
}}
.fiesta-kpi-note {{ color: #78909c !important; font-size: 0.8rem; margin-top: 4px; }}
</style>
<div class="fiesta-analytics-wrap">
<p class="fiesta-analytics-title">{html_module.escape(title_main)}</p>
<p class="fiesta-analytics-sub">Rep incentive data — quota, achievement, and payouts</p>
</div>
""",
        unsafe_allow_html=True,
    )

    if "calculation_period" in df.columns and df["calculation_period"].notna().any():
        periods = sorted({str(x) for x in df["calculation_period"].dropna().unique()})
        if len(periods) > 1:
            pick = st.selectbox("Calculation period", periods, key=period_key)
            df = df[df["calculation_period"].astype(str) == pick]

    if "team_name" in df.columns:
        team_mask = df["team_name"].astype(str).str.strip() == team_for_goal
    else:
        team_mask = pd.Series([True] * len(df), index=df.index)
    df_team = df[team_mask].copy()
    work = df_team if not df_team.empty else df.copy()

    if work.empty:
        st.warning("No rows to chart for this selection.")
        return

    team_goal = float(TEAM_QUARTERLY_TARGETS_USD.get(team_for_goal, 0) or 0) if not work.empty else 0.0
    team_rev = float(pd.to_numeric(work["total_revenue"], errors="coerce").fillna(0).sum())
    total_inc = float(pd.to_numeric(work["incentive_amount"], errors="coerce").fillna(0).sum())
    total_paid = float(pd.to_numeric(work["total_paid_amount"], errors="coerce").fillna(0).sum())
    pct_team = (team_rev / team_goal * 100.0) if team_goal > 0 else 0.0

    kpi_html = f"""
<div class="fiesta-kpi-row">
<div class="fiesta-kpi"><div class="fiesta-kpi-label">Team revenue</div><div class="fiesta-kpi-val">${team_rev:,.0f}</div>
<div class="fiesta-kpi-note">vs goal ${team_goal:,.0f}</div></div>
<div class="fiesta-kpi"><div class="fiesta-kpi-label">Team attainment</div><div class="fiesta-kpi-val">{pct_team:.1f}%</div>
<div class="fiesta-kpi-note">Quarterly team target</div></div>
<div class="fiesta-kpi"><div class="fiesta-kpi-label">Total incentive</div><div class="fiesta-kpi-val">${total_inc:,.0f}</div>
<div class="fiesta-kpi-note">Paid-based commission</div></div>
<div class="fiesta-kpi"><div class="fiesta-kpi-label">Total paid (deals)</div><div class="fiesta-kpi-val">${total_paid:,.0f}</div>
<div class="fiesta-kpi-note">Basis for %</div></div>
</div>
"""
    st.markdown(kpi_html, unsafe_allow_html=True)

    col_g, col_b = st.columns([1, 1.2])
    with col_g:
        st.caption("**Team progress vs quarterly goal**")
        if team_goal > 0:
            p = min(1.0, max(0.0, team_rev / team_goal))
            ring = pd.DataFrame({"kind": ["Achieved", "Remaining"], "v": [p, max(0.0, 1.0 - p)]})
            base = (
                alt.Chart(ring)
                .mark_arc(innerRadius=58, outerRadius=92, padAngle=0.02)
                .encode(
                    theta=alt.Theta("v:Q", stack=True),
                    color=alt.Color(
                        "kind:N",
                        scale=alt.Scale(domain=["Achieved", "Remaining"], range=["#26c6da", "#7e57c2"]),
                        legend=alt.Legend(orient="bottom", labelColor="#e3f2fd", title=None),
                    ),
                )
            )
            st.altair_chart(_fiesta_alt_configure(base.properties(width=240, height=240), height=260), use_container_width=True)
            st.caption(f"{pct_team:.1f}% of ${team_goal:,.0f} team goal")
        else:
            st.caption("Set team goal in policy / **Teams** to show progress ring.")

    with col_b:
        st.caption("**Incentive payout by rep** (horizontal)")
        plot = work[["full_name", "incentive_amount"]].copy()
        plot["incentive_amount"] = pd.to_numeric(plot["incentive_amount"], errors="coerce").fillna(0)
        plot = plot.sort_values("incentive_amount", ascending=True)
        plot["label"] = plot["incentive_amount"].apply(lambda v: f"${float(v):,.0f}")
        _fn_order = plot["full_name"].astype(str).tolist()
        hb = (
            alt.Chart(plot)
            .mark_bar(cornerRadiusEnd=4)
            .encode(
                x=alt.X("incentive_amount:Q", title="Incentive ($)"),
                y=alt.Y("full_name:N", sort="-x", title=""),
                color=_chart_categorical_color("full_name:N", _fn_order, legend=None),
                tooltip=["full_name", alt.Tooltip("incentive_amount:Q", format=",.2f", title="Incentive")],
            )
        )
        ht = hb.mark_text(align="left", baseline="middle", dx=5, color="#f1f5f9").encode(
            x="incentive_amount:Q", y="full_name:N", text="label:N"
        )
        st.altair_chart(
            _fiesta_alt_configure(hb + ht, height=max(220, 28 * len(plot))),
            use_container_width=True,
        )

    st.caption("**Quota achievement % by rep** (50% minimum for payout)")
    qa = work[["full_name", "quota_achievement_pct"]].copy()
    qa["quota_achievement_pct"] = pd.to_numeric(qa["quota_achievement_pct"], errors="coerce").fillna(0)
    qa = qa.sort_values("quota_achievement_pct", ascending=True)
    xmax = max(120.0, float(qa["quota_achievement_pct"].max()) * 1.08, 50.0)
    qa["label"] = qa["quota_achievement_pct"].apply(lambda v: f"{float(v):.1f}%")
    _qa_names = qa["full_name"].astype(str).tolist()
    ach = (
        alt.Chart(qa)
        .mark_bar(cornerRadiusEnd=4)
        .encode(
            x=alt.X("quota_achievement_pct:Q", title="Achievement %", scale=alt.Scale(domain=[0, xmax])),
            y=alt.Y("full_name:N", sort="-x", title=""),
            color=_chart_categorical_color("full_name:N", _qa_names, legend=None),
            tooltip=["full_name", alt.Tooltip("quota_achievement_pct:Q", format=".2f", title="Achievement %")],
        )
    )
    tx = ach.mark_text(align="left", baseline="middle", dx=5, color="#f1f5f9").encode(
        x="quota_achievement_pct:Q", y="full_name:N", text="label:N"
    )
    st.altair_chart(
        _fiesta_alt_configure(ach + tx, height=max(240, 28 * len(qa))),
        use_container_width=True,
    )

    st.caption("**Total revenue by rep** (columns)")
    rev = work[["full_name", "total_revenue"]].copy()
    rev["total_revenue"] = pd.to_numeric(rev["total_revenue"], errors="coerce").fillna(0)
    rev = rev.sort_values("total_revenue", ascending=False)
    rev["label"] = rev["total_revenue"].apply(lambda v: f"{float(v):,.0f}")
    _rev_names = rev["full_name"].astype(str).tolist()
    vb = (
        alt.Chart(rev)
        .mark_bar(cornerRadiusEnd=4)
        .encode(
            x=alt.X("full_name:N", sort=_rev_names, title="", axis=alt.Axis(labelAngle=-35)),
            y=alt.Y("total_revenue:Q", title="Revenue ($)"),
            color=_chart_categorical_color("full_name:N", _rev_names, legend=None),
            tooltip=["full_name", alt.Tooltip("total_revenue:Q", format=",.2f", title="Revenue")],
        )
    )
    vtx = vb.mark_text(dy=-6, align="center", color="#f1f5f9").encode(text="label:N")
    st.altair_chart(_fiesta_alt_configure(vb + vtx, height=300), use_container_width=True)


def _chart_pie(df: pd.DataFrame, name_col: str, value_col: str, height: int = 320):
    """Pie chart with slice labels and tooltip; name_col = slice label, value_col = slice size."""
    d = df[[name_col, value_col]].copy()
    d = d.rename(columns={name_col: "name", value_col: "value"})
    d["value"] = d["value"].astype(float)
    total = d["value"].sum()
    if total <= 0:
        st.altair_chart(alt.Chart(pd.DataFrame()).mark_text(text="No data").properties(height=height), use_container_width=True, theme=None)
        return
    d["pct"] = d["value"] / total
    d["start_pct"] = d["pct"].cumsum().shift(1).fillna(0)
    d["mid_pct"] = d["start_pct"] + d["pct"] / 2
    d["mid_rad"] = d["mid_pct"] * 2 * math.pi
    d["lx"] = 0.45 * d["mid_rad"].apply(math.cos)
    d["ly"] = 0.45 * d["mid_rad"].apply(math.sin)
    d["label"] = d["name"].astype(str) + " (" + d["value"].apply(lambda v: f"{float(v):,.0f}") + ")"
    _pie_names = d["name"].astype(str).tolist()
    pie = alt.Chart(d).mark_arc(innerRadius=0).encode(
        theta=alt.Theta("value:Q", stack=True),
        color=_chart_categorical_color("name:N", _pie_names, legend=alt.Legend(title=name_col)),
        tooltip=[alt.Tooltip("name:N", title=name_col), alt.Tooltip("value:Q", title=value_col, format=",.0f")],
    ).properties(height=height)
    text = alt.Chart(d).mark_text(align="center", baseline="middle", fontSize=10, color="#0d47a1").encode(
        x=alt.X("lx:Q", scale=alt.Scale(domain=[-1, 1])),
        y=alt.Y("ly:Q", scale=alt.Scale(domain=[-1, 1])),
        text="label:N",
    ).properties(height=height)
    st.altair_chart(pie + text, use_container_width=True, theme=None)


def init_session():
    if "user" not in st.session_state:
        st.session_state.user = None
    # Run every rerun so existing DBs get Outbound even if session was initialized before this team existed.
    try:
        ensure_outbound_team()
    except Exception:
        pass
    if "initialized" not in st.session_state:
        st.session_state.initialized = True
        try:
            initialize_schema()
            ensure_admin_user()
        except Exception as e:
            st.error(f"Database setup failed: {e}")
            st.stop()


def _login_logo_path():
    """Path to CloudFuze logo (assets folder in project root). Uses first preferred name, else any PNG/JPG."""
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
    if not os.path.isdir(base):
        return None
    # Preferred names, including Cursor-saved image name when user attaches logo
    preferred = [
        "cloudfuze_logo.png",
        "cloudfuze_logo.jpg",
        "CloudFuze_Logo.png",
        "c__Users_SakshiPriya_AppData_Roaming_Cursor_User_workspaceStorage_1383515b60ceed97fcaa06006a47ee66_images_CloudFuze_Logo-a5165e69-7c32-41b3-a1c9-ec3b30efa757.png",
        "logo.png",
        "logo.jpg",
    ]
    for name in preferred:
        path = os.path.join(base, name)
        if os.path.isfile(path):
            return path
    for f in os.listdir(base):
        if f.lower().endswith((".png", ".jpg", ".jpeg")) and "cloudfuze" in f.lower():
            return os.path.join(base, f)
    for f in os.listdir(base):
        if f.lower().endswith((".png", ".jpg", ".jpeg")):
            return os.path.join(base, f)
    return None


def _login_page_css():
    """Login: split navy / light hero (personal-finance poster style), blue accents only."""
    return """
    <style>
    html, body, .stApp {
        font-family: "Amasis MT Pro", "Amasis MT", Georgia, serif !important;
    }
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > section,
    [data-testid="stAppViewContainer"] > section > div,
    section.main > div {
        background: #e8ecf2 !important;
        min-height: 100vh;
    }
    .stApp { overflow-x: hidden !important; }
    .main .block-container {
        padding-top: 1.25rem;
        max-width: min(1140px, 100%);
        margin-left: auto;
        margin-right: auto;
        background: transparent !important;
    }
    /* --- Login v2: split hero --- */
    .login-v2-wrap {
        width: 100vw;
        max-width: 100vw;
        position: relative;
        left: 50%;
        transform: translateX(-50%);
        margin-bottom: 1.75rem;
        min-height: min(480px, 82vh);
        background: linear-gradient(102deg,
            #051a32 0%,
            #0c2744 52%,
            #dfe6ee 52.2%,
            #eef1f7 72%,
            #e4eaf2 100%);
        border-radius: 0 0 22px 22px;
        box-shadow: 0 20px 50px rgba(5, 26, 50, 0.28);
        overflow: hidden;
    }
    .login-v2-split {
        display: grid;
        grid-template-columns: minmax(220px, min(400px, 38vw)) minmax(120px, 200px) minmax(200px, 1.15fr);
        gap: 0.75rem 1rem;
        align-items: center;
        padding: 2.25rem clamp(0.85rem, 2.5vw, 2.25rem) 1.75rem;
        max-width: 1180px;
        margin: 0 auto;
        min-height: 420px;
    }
    @media (max-width: 900px) {
        .login-v2-split {
            grid-template-columns: 1fr;
            min-height: auto;
        }
        .login-v2-left {
            border-radius: 14px;
            margin: 0 0 0.5rem 0;
            max-width: 100%;
        }
        .login-v2-kicker, .login-v2-bullets { max-width: 100%; }
        .login-v2-sheets { min-height: 320px !important; margin-top: 1rem; }
        .login-v2-sheet-a { right: 12% !important; }
        .login-v2-sheet-b { right: 4% !important; }
    }
    .login-v2-left {
        font-family: "Segoe UI", "Aptos", system-ui, sans-serif !important;
        color: #fff !important;
        z-index: 3;
        align-self: stretch;
        display: flex;
        flex-direction: column;
        justify-content: center;
        /* Opaque navy panel so white type never sits on the pale diagonal */
        background: linear-gradient(165deg, #061a33 0%, #0e3052 55%, #0a2540 100%);
        border-radius: 6px 22px 22px 6px;
        padding: 1.5rem 1.35rem 1.5rem 1.1rem;
        margin: 0.35rem 0.35rem 0.35rem 0;
        max-width: 100%;
        box-shadow: 8px 4px 28px rgba(4, 18, 40, 0.35);
        border: 1px solid rgba(100, 181, 246, 0.15);
    }
    .login-v2-kicker {
        margin: 0 0 1rem 0 !important;
        font-size: 0.88rem !important;
        font-style: italic !important;
        font-weight: 400 !important;
        color: rgba(255,255,255,0.88) !important;
        line-height: 1.45 !important;
        max-width: 22rem;
    }
    .login-v2-title {
        margin: 0 !important;
        padding: 0 !important;
        font-family: "Segoe UI", "Aptos", system-ui, sans-serif !important;
        font-size: clamp(1.35rem, 3.8vw, 2.05rem) !important;
        font-weight: 800 !important;
        line-height: 1.12 !important;
        letter-spacing: 0.06em !important;
        text-transform: uppercase !important;
        color: #ffffff !important;
        max-width: 21rem;
        word-wrap: break-word;
    }
    .login-v2-title strong { font-weight: 800 !important; color: #fff !important; }
    .login-v2-rule {
        height: 2px;
        width: 100%;
        max-width: 260px;
        margin: 1.1rem 0 1.15rem;
        background: #ffffff;
        opacity: 0.95;
        border-radius: 1px;
    }
    .login-v2-bullets {
        margin: 0 0 1.5rem 0 !important;
        padding: 0 !important;
        list-style: none !important;
        font-size: 0.92rem !important;
        line-height: 1.65 !important;
        color: rgba(255,255,255,0.95) !important;
        max-width: 22rem;
    }
    .login-v2-bullets li { margin-bottom: 0.35rem; padding-left: 0; word-wrap: break-word; }
    .login-v2-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 0.65rem;
        align-items: center;
    }
    .login-v2-badge {
        width: 44px;
        height: 44px;
        border-radius: 50%;
        background: #ffffff !important;
        color: #0a2540 !important;
        font-family: "Segoe UI", "Aptos", system-ui, sans-serif !important;
        font-size: 0.62rem !important;
        font-weight: 800 !important;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 4px 14px rgba(0,0,0,0.18);
        letter-spacing: 0.02em;
    }
    .login-v2-center {
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 4;
    }
    .login-v2-card {
        width: min(168px, 42vw);
        height: min(168px, 42vw);
        background: #ffffff !important;
        border-radius: 20px !important;
        box-shadow: 0 18px 50px rgba(0, 20, 50, 0.22), 0 0 0 1px rgba(255,255,255,0.8);
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 12px;
    }
    .login-v2-card svg { width: 100%; height: 100%; max-width: 132px; max-height: 132px; }
    .login-v2-sheets {
        position: relative;
        min-height: 400px;
        z-index: 1;
    }
    .login-v2-sheet {
        position: absolute;
        background: #ffffff !important;
        border-radius: 16px !important;
        padding: 14px 14px 16px;
        box-shadow: 0 14px 36px rgba(15, 40, 70, 0.14);
        border: 1px solid rgba(100, 181, 246, 0.35);
        width: min(268px, 72vw);
    }
    .login-v2-sheet-a {
        transform: rotate(-9deg);
        top: 12px;
        right: 18%;
        z-index: 1;
        opacity: 0.97;
    }
    .login-v2-sheet-b {
        transform: rotate(7deg);
        top: 108px;
        right: 4%;
        z-index: 2;
        border-color: rgba(25, 118, 210, 0.45);
        box-shadow: 0 20px 44px rgba(25, 118, 210, 0.12);
    }
    .login-v2-icon-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 10px;
    }
    .login-v2-grid-cell {
        width: 100%;
        aspect-ratio: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(145deg, #f5f8fc 0%, #e8f0f8 100%);
        border-radius: 10px;
        border: 1px solid rgba(13, 71, 161, 0.12);
    }
    .login-v2-sheet-b .login-v2-grid-cell {
        background: linear-gradient(145deg, #e3f2fd 0%, #bbdefb 100%);
        border-color: rgba(25, 118, 210, 0.22);
    }
    .login-v2-grid-cell svg { width: 24px; height: 24px; display: block; }
    .main .block-container .stForm button,
    .main .block-container .stForm button[kind="formSubmit"] {
        background: linear-gradient(135deg, #1976d2 0%, #0d47a1 100%) !important;
        color: #fff !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.5rem 1.5rem !important;
        box-shadow: 0 4px 14px rgba(13, 71, 161, 0.45) !important;
    }
    /* Form below hero sits on light gray — dark labels */
    .main .block-container .stTextInput label,
    .main .block-container [data-testid="stForm"] label,
    .main .block-container label,
    .main .block-container [data-testid="stTextInput"] label,
    .main .block-container p label {
        color: #0d2137 !important;
    }
    /* Input fields: dark text so typed email/password are visible (labels stay white above) */
    .main .block-container .stTextInput input,
    .main .block-container input[type="text"],
    .main .block-container input[type="password"],
    .main .block-container [data-testid="stTextInput"] input {
        background: #ffffff !important;
        border: 1px solid #30363d !important;
        color: #1a1a1a !important;
        -webkit-text-fill-color: #1a1a1a !important;
    }
    .main .block-container .stTextInput input::placeholder,
    .main .block-container input::placeholder {
        color: #666666 !important;
        opacity: 1;
    }
    .main .block-container input::-webkit-input-placeholder {
        color: #666666 !important;
    }
    .main .block-container input::-moz-placeholder {
        color: #666666 !important;
    }
    .stApp label[for], [data-testid="stAppViewContainer"] label {
        color: #0d2137 !important;
    }
    .stApp input[type="text"], .stApp input[type="password"],
    [data-testid="stAppViewContainer"] input[type="text"],
    [data-testid="stAppViewContainer"] input[type="password"] {
        background-color: #ffffff !important;
        background: #ffffff !important;
        border: 1px solid #30363d !important;
        color: #1a1a1a !important;
        -webkit-text-fill-color: #1a1a1a !important;
    }
    .stApp input[type="text"]::placeholder, .stApp input[type="password"]::placeholder {
        color: #666666 !important;
    }
    .login-title-custom {
        text-align: center;
        color: #ffffff !important;
        font-weight: 700;
        margin-bottom: 0.2rem !important;
    }
    .login-subtitle-custom {
        text-align: center;
        color: #1565c0 !important;
        font-size: 0.95rem !important;
        margin-bottom: 0.35rem !important;
        font-weight: 600 !important;
    }
    .login-caption-custom {
        text-align: center;
        color: #37474f !important;
        margin-bottom: 1rem !important;
        font-weight: 500;
        font-size: 0.9rem !important;
    }
    </style>
    """


# Legacy tagline (login hero now uses HTML headline in ``render_login``).
LOGIN_HEADING_MSG = (
    "You've unlocked the CloudFuze Compensation Tool — "
    "where commissions stop hiding and start multiplying. 💸🚀"
)
# Spoken voice message (Sara) on Play button and after Sign In
WELCOME_VOICE_MSG = "Welcome to the CloudFuze Compensation Tool."


def _get_sara_voice_js():
    """JS snippet to pick Sara (female) voice, else first female English voice."""
    return """
    function pickVoice(voices) {
        var sara = voices.find(function(v) { return v.name.toLowerCase().indexOf('sara') !== -1 && v.lang.startsWith('en'); });
        if (sara) return sara;
        var femaleEn = voices.find(function(v) { return v.lang.startsWith('en') && (v.name.toLowerCase().indexOf('female') !== -1 || v.name.toLowerCase().indexOf('sara') !== -1 || v.name.toLowerCase().indexOf('zira') !== -1); });
        if (femaleEn) return femaleEn;
        return voices.find(function(v) { return v.lang.startsWith('en'); }) || null;
    }
    """


def _login_voice_html(autoplay=False):
    """HTML + JS for welcome message: Sara (female) voice; button or auto-play on load."""
    msg = WELCOME_VOICE_MSG
    voice_js = _get_sara_voice_js()
    if autoplay:
        return f"""
        <script>
        (function() {{
            var msg = {repr(msg)};
            {voice_js}
            function speak() {{
                if (window.speechSynthesis) {{
                    window.speechSynthesis.cancel();
                    var u = new SpeechSynthesisUtterance(msg);
                    u.rate = 0.95;
                    u.pitch = 1;
                    var voices = speechSynthesis.getVoices();
                    if (voices.length) {{
                        var chosen = pickVoice(voices);
                        if (chosen) u.voice = chosen;
                    }}
                    speechSynthesis.speak(u);
                }}
            }}
            function run() {{
                var voices = speechSynthesis.getVoices();
                if (voices.length) speak();
                else speechSynthesis.addEventListener('voiceschanged', function() {{ speak(); }}, {{ once: true }});
            }}
            if (document.readyState === 'complete') setTimeout(run, 300);
            else window.addEventListener('load', function() {{ setTimeout(run, 400); }});
            if (window.speechSynthesis) speechSynthesis.getVoices();
        }})();
        </script>
        """
    return f"""
    <div style="text-align: center; margin: 1rem 0;">
        <button id="play-voice-btn" style="
            background: linear-gradient(135deg, #1976d2 0%, #0d47a1 100%);
            color: #fff;
            border: none;
            border-radius: 24px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 2px 12px rgba(13,71,161,0.45);
        " onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
            🔊 Play welcome message
        </button>
    </div>
    <script>
    (function() {{
        var btn = document.getElementById('play-voice-btn');
        var msg = {repr(msg)};
        {voice_js}
        function play() {{
            if (window.speechSynthesis) {{
                window.speechSynthesis.cancel();
                var u = new SpeechSynthesisUtterance(msg);
                u.rate = 0.95;
                u.pitch = 1;
                var voices = speechSynthesis.getVoices();
                if (voices.length) {{ var chosen = pickVoice(voices); if (chosen) u.voice = chosen; }}
                speechSynthesis.speak(u);
            }}
        }}
        btn.addEventListener('click', function() {{
            var voices = speechSynthesis.getVoices();
            if (voices.length) play();
            else speechSynthesis.addEventListener('voiceschanged', function() {{ play(); }}, {{ once: true }});
        }});
        if (window.speechSynthesis) speechSynthesis.getVoices();
    }})();
    </script>
    """


def _login_mini_icon_svgs() -> list[str]:
    """Small flat compensation-themed icons (24×24, two-tone blue)."""
    c1, c2, c3 = "#90caf9", "#42a5f5", "#0d47a1"
    return [
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="9" cy="18" r="3" fill="{c2}"/><circle cx="15" cy="16" r="3" fill="{c1}"/><circle cx="12" cy="11" r="3" fill="{c3}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="4" y="14" width="4" height="6" rx="1" fill="{c2}"/><rect x="10" y="10" width="4" height="10" rx="1" fill="{c1}"/><rect x="16" y="6" width="4" height="14" rx="1" fill="{c3}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><ellipse cx="12" cy="15" rx="8" ry="6" fill="{c2}"/><circle cx="8" cy="11" r="2" fill="{c1}"/><ellipse cx="14" cy="9" rx="3" ry="2" fill="{c3}"/><rect x="10" y="13" width="4" height="2" rx="0.5" fill="{c3}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M12 3 L18 8 L18 20 L6 20 L6 8 Z" fill="{c2}"/><path d="M12 3 L18 8 L6 8 Z" fill="{c1}"/><text x="12" y="16" text-anchor="middle" font-size="8" font-weight="700" fill="{c3}">$</text></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="5" y="8" width="14" height="11" rx="1" fill="{c2}"/><path d="M8 8 V6 Q12 4 16 6 V8" fill="none" stroke="{c3}" stroke-width="1.5"/><rect x="10" y="12" width="4" height="2" fill="{c1}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="12" cy="12" r="9" fill="none" stroke="{c1}" stroke-width="2"/><circle cx="12" cy="12" r="5" fill="none" stroke="{c2}" stroke-width="2"/><circle cx="12" cy="12" r="2" fill="{c3}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="6" y="10" width="12" height="9" rx="1" fill="{c2}"/><rect x="9" y="7" width="6" height="5" rx="1" fill="{c1}"/><path d="M12 7 V5" stroke="{c3}" stroke-width="1.5"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><circle cx="12" cy="12" r="9" fill="{c2}"/><text x="12" y="15" text-anchor="middle" font-size="9" font-weight="700" fill="{c3}">%</text></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M6 18 Q6 12 12 10 Q18 12 18 18" fill="none" stroke="{c2}" stroke-width="2"/><circle cx="12" cy="8" r="3" fill="{c1}"/><circle cx="12" cy="8" r="1.5" fill="{c3}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M5 19 L19 19 L19 8 L14 5 L10 5 L5 8 Z" fill="{c2}"/><rect x="8" y="11" width="8" height="5" fill="{c1}"/><path d="M10 5 V3 M14 5 V3" stroke="{c3}" stroke-width="1.5"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><rect x="4" y="6" width="16" height="12" rx="1" fill="{c2}"/><rect x="6" y="9" width="4" height="6" fill="{c1}"/><rect x="11" y="11" width="3" height="4" fill="{c3}"/><rect x="15" y="8" width="3" height="7" fill="{c1}"/></svg>',
        f'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M12 4 L15 10 L22 11 L17 16 L18 23 L12 19 L6 23 L7 16 L2 11 L9 10 Z" fill="{c2}"/><path d="M12 4 L15 10 L12 14 L9 10 Z" fill="{c1}"/></svg>',
    ]


def _login_v2_shield_card_svg() -> str:
    """White card graphic: navy shield + blue $ circle (reference layout, blue only)."""
    return """
<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <path d="M50 6 L88 23 V50 Q88 70 50 94 Q12 70 12 50 V23 Z" fill="none" stroke="#0d47a1" stroke-width="2.5" stroke-linejoin="round"/>
  <circle cx="64" cy="40" r="20" fill="#1976d2"/>
  <text x="64" y="46" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-size="19" font-weight="700" fill="#ffffff">$</text>
</svg>
"""


def _login_sheet_icon_svgs(n: int = 16) -> list[str]:
    icons = list(_login_mini_icon_svgs())
    while len(icons) < n:
        icons.extend(_login_mini_icon_svgs())
    return icons[:n]


def _login_v2_icon_grid_inner_html() -> str:
    """4×4 grid of mini icons; parent ``.login-v2-sheet-b`` styles accent cells."""
    cells = "".join(
        f'<div class="login-v2-grid-cell">{svg}</div>' for svg in _login_sheet_icon_svgs(16)
    )
    return f'<div class="login-v2-icon-grid">{cells}</div>'


def _login_v2_hero_html() -> str:
    card = _login_v2_shield_card_svg().strip()
    grid_html = _login_v2_icon_grid_inner_html()
    badges = "".join(f'<span class="login-v2-badge">{lbl}</span>' for lbl in ("SVG", "PNG", "AI", "EPS"))
    return f"""
<div class="login-v2-wrap">
  <div class="login-v2-split">
    <div class="login-v2-left">
      <p class="login-v2-kicker">Track attainment, calculate commissions, and export payouts — for CloudFuze sales teams.</p>
      <h1 class="login-v2-title"><strong>CloudFuze sales compensation</strong></h1>
      <div class="login-v2-rule" aria-hidden="true"></div>
      <ul class="login-v2-bullets">
        <li>• HubSpot &amp; Excel deal imports</li>
        <li>• Policy-based slabs &amp; commission engine</li>
        <li>• Rep, manager &amp; team incentive views</li>
      </ul>
      <div class="login-v2-badges" aria-hidden="true">{badges}</div>
    </div>
    <div class="login-v2-center">
      <div class="login-v2-card">{card}</div>
    </div>
    <div class="login-v2-sheets" aria-hidden="true">
      <div class="login-v2-sheet login-v2-sheet-a">{grid_html}</div>
      <div class="login-v2-sheet login-v2-sheet-b">{grid_html}</div>
    </div>
  </div>
</div>
"""


def render_login():
    st.markdown(_login_page_css(), unsafe_allow_html=True)

    st.markdown(_login_v2_hero_html(), unsafe_allow_html=True)

    logo_path = _login_logo_path()
    if logo_path:
        st.image(logo_path, width=180)

    st.markdown("<p class='login-subtitle-custom'>CloudFuze Migrate · Incentive Calculator</p>", unsafe_allow_html=True)
    st.markdown("<p class='login-caption-custom'>Sign in to access your dashboard</p>", unsafe_allow_html=True)

    with st.form("login_form"):
        email = st.text_input("Email", placeholder="admin@cloudfuze.com")
        password = st.text_input("Password", type="password", placeholder="••••••••")
        submitted = st.form_submit_button("Sign In")

        if submitted:
            if not email or not password:
                st.error("Please enter email and password")
            else:
                user = authenticate(email.strip(), password)
                if user:
                    st.session_state.user = user
                    log_audit("LOGIN", "user", performed_by=user["user_id"], entity_id=str(user["user_id"]))
                    st.rerun()
                else:
                    st.error("Invalid email or password")

    # Voice message removed per request — no TTS button on the login page either.


def render_branding_and_settings(user: dict) -> None:
    """Top-right: CloudFuze Migrate, logged-in line, gear icon → Sign out + Security → Change password."""
    name = html_module.escape(str(user.get("full_name", "")))
    role = html_module.escape(str(user.get("role", "")))
    st.markdown(
        f'<p style="text-align:right;margin:0 0 0.15rem 0;font-weight:700;font-size:1.15rem;font-family:Aptos,Segoe UI,sans-serif;">CloudFuze Migrate</p>'
        f'<p style="text-align:right;margin:0;color:#5f6368;font-size:0.88rem;">Logged in as {name} ({role})</p>',
        unsafe_allow_html=True,
    )
    _, gear_col = st.columns([1, 0.2])
    with gear_col:
        with st.popover("⚙️", help="Settings", type="tertiary", use_container_width=True):
            if st.button("Sign out", key="header_sign_out", use_container_width=True):
                uid = st.session_state.user.get("user_id")
                log_audit("LOGOUT", "user", performed_by=uid, entity_id=str(uid) if uid else None)
                st.session_state.user = None
                st.rerun()
            with st.expander("Security", expanded=False):
                st.caption("Change password")
                with st.form("change_password_form"):
                    current = st.text_input("Current password", type="password", key="cp_current")
                    new_pw = st.text_input("New password", type="password", key="cp_new")
                    confirm = st.text_input("Confirm new password", type="password", key="cp_confirm")
                    submitted = st.form_submit_button("Update Password")
                    if submitted:
                        if not current or not new_pw or not confirm:
                            st.error("Fill all fields")
                            return
                        if new_pw != confirm:
                            st.error("New password and confirmation don't match")
                            return
                        if len(new_pw) < 6:
                            st.error("New password must be at least 6 characters")
                            return
                        u = st.session_state.user
                        db_user = get_user_by_email(u["email"])
                        if not db_user or not verify_password(current, db_user["password_hash"]):
                            st.error("Current password is incorrect")
                            return
                        update_user_password(u["user_id"], hash_password(new_pw))
                        log_audit("CHANGE_PASSWORD", "user", performed_by=u["user_id"], entity_id=str(u["user_id"]))
                        st.success("Password updated. You may need to sign in again.")
                        st.rerun()


def _admin_team_key_prefix(team_name: str) -> str:
    return (team_name or "").strip().lower().replace(" ", "_")


def _filter_deals_by_team_name(deals: list, team_name: str) -> list:
    users = get_all_users_with_teams(active_only=False)
    uids = {int(u["user_id"]) for u in users if (u.get("team_name") or "").strip() == team_name}
    out: list = []
    for d in deals or []:
        oid = d.get("deal_owner_id")
        if oid is None:
            continue
        try:
            if int(oid) in uids:
                out.append(d)
        except (TypeError, ValueError):
            continue
    return out


def render_rep_incentives_admin(user_id: int, team_filter: str, key_prefix: str) -> None:
    """Rep incentives table and deletes for one sales team (e.g. SMB, Account Management, Enterprise)."""
    incentives = get_rep_incentives()
    if not incentives:
        st.info("No rep incentives yet. Finalize an upload to generate.")
        return
    filtered = [r for r in incentives if (r.get("team_name") or "").strip() == team_filter]
    if not filtered:
        st.info(f"No rep incentives for **{team_filter}** yet.")
        return
    st.caption(
        "These rows are **saved in the database** when you **Finalize** an upload. "
        "**HubSpot Fetch** only loads deals for review—it does **not** create or refresh this table."
    )
    if team_filter == "SMB":
        st.caption(
            "**SMB reps (A/B):** commission % from **policy/** tier table; achievement % = total revenue ÷ individual quota. "
            "Group from **user compensation_group** or quota list match. **Re-finalize** an upload after policy changes."
        )
    else:
        st.caption(f"**{team_filter}** — policy achievement bands (from last **Finalize** for each period).")
    df = pd.DataFrame(filtered)
    df = _enrich_rep_incentives_display(df)
    if "created_at" in df.columns:
        ts = pd.to_datetime(df["created_at"], errors="coerce")
        df["Stored at"] = ts.dt.strftime("%Y-%m-%d %H:%M").fillna("—")
    cols = [
        "full_name",
        "team_name",
        "quota",
        "quota_achievement_pct",
        "total_revenue",
        "total_paid_amount",
        "incentive_percentage",
        "incentive_amount",
        "incentive_eligibility",
        "calculation_period",
    ]
    if "Stored at" in df.columns:
        cols.append("Stored at")
    show = [c for c in cols if c in df.columns]
    base = df[show] if show else df
    df_display = _format_df_dollars(base, ["quota", "total_revenue", "total_paid_amount", "incentive_amount"])
    df_display = _format_pct_display(df_display, ["quota_achievement_pct", "incentive_percentage"])
    st.dataframe(df_display, use_container_width=True)
    st.divider()
    st.subheader("Delete rep incentive")
    options = {
        f"ID {r['rep_incentive_id']}: {r.get('full_name', '')} – {r.get('calculation_period', '')} – ${float(r.get('incentive_amount', 0)):,.2f}": r["rep_incentive_id"]
        for r in filtered
    }
    selected = st.selectbox(
        "Choose one to delete",
        options=[""] + list(options.keys()),
        key=f"del_rep_select_{key_prefix}",
    )
    if selected and selected in options:
        if st.button("Delete this rep incentive", key=f"del_rep_btn_{key_prefix}"):
            delete_rep_incentive(options[selected])
            log_audit("DELETE_REP_INCENTIVE", "rep_incentive", user_id, str(options[selected]), selected)
            st.success("Rep incentive deleted.")
            st.rerun()
    st.divider()
    confirm_del_rep = st.checkbox(
        "I confirm I want to delete all rep incentives for this team",
        key=f"confirm_del_all_rep_{key_prefix}",
    )
    if confirm_del_rep and st.button("Delete all rep incentives (this team)", type="secondary", key=f"del_all_rep_{key_prefix}"):
        to_del = [r["rep_incentive_id"] for r in filtered]
        for rid in to_del:
            delete_rep_incentive(rid)
        log_audit("DELETE_REP_INCENTIVES_TEAM", "rep_incentive", user_id, team_filter, str(len(to_del)))
        st.success(f"Deleted {len(to_del)} rep incentive row(s) for {team_filter}.")
        st.rerun()


def render_manager_incentive_admin(user_id: int, team_filter: str, key_prefix: str) -> None:
    """Team goal editor + manager incentive table for one team. Metrics calls this twice to show SMB and AM."""
    st.subheader("Team goal")
    teams = [t for t in get_all_teams() if (t.get("team_name") or "").strip() == team_filter]
    if not teams:
        st.info(f"No **{team_filter}** team found. Add it under User Management.")
    for t in teams:
        with st.container():
            col_name, col_goal, col_btn = st.columns([2, 2, 1])
            with col_name:
                st.write(f"**{t['team_name']}**")
            with col_goal:
                goal_key = f"team_goal_{key_prefix}_{t['team_id']}"
                current = float(t["team_goal"]) if t.get("team_goal") is not None else 0.0
                new_goal = st.number_input("Goal ($)", min_value=0.0, value=current, step=1000.0, key=goal_key)
            with col_btn:
                if st.button("Update", key=f"goal_btn_{key_prefix}_{t['team_id']}"):
                    update_team_goal(t["team_id"], new_goal if new_goal > 0 else None)
                    log_audit("UPDATE_TEAM_GOAL", "team", user_id, str(t["team_id"]), t["team_name"])
                    st.success("Updated.")
                    st.rerun()
    st.divider()
    st.subheader("Manager incentive")
    incentives = get_team_incentives()
    filtered = [r for r in (incentives or []) if (r.get("team_name") or "").strip() == team_filter]
    if not filtered:
        st.info("No manager incentive rows for this team yet. Set team goals and finalize an upload.")
        return
    st.caption(
        "Incentive = total_team_revenue × commission %. **SMB + Chitradip:** tiers vs **team goal** from policy JSON. "
        "**Other teams:** standard achievement bands."
    )
    df = _enrich_team_incentives_display(pd.DataFrame(filtered))
    cols = [
        "team_name",
        "team_lead_name",
        "total_team_revenue",
        "Team goal achievement %",
        "incentive_percentage",
        "incentive_amount",
        "Payout note",
    ]
    show = [c for c in cols if c in df.columns]
    base_df = df[show] if show else df
    dollar_cols = [c for c in ["total_team_revenue", "incentive_amount"] if c in base_df.columns]
    df_display = _format_df_dollars(base_df, dollar_cols)
    if "Team goal achievement %" in df_display.columns:
        df_display = df_display.copy()
        df_display["Team goal achievement %"] = pd.to_numeric(
            df_display["Team goal achievement %"], errors="coerce"
        ).fillna(0.0)
    _cfg = {k: v for k, v in _manager_incentive_column_config().items() if k in df_display.columns}
    _df_kw = {"use_container_width": True}
    if _cfg:
        _df_kw["column_config"] = _cfg
    st.dataframe(df_display, **_df_kw)
    st.divider()
    st.subheader("Delete manager incentive")
    team_options = {
        f"ID {r['team_incentive_id']}: {r.get('team_name', '')} – {r.get('calculation_period', '')} – ${float(r.get('incentive_amount', 0)):,.2f}": r["team_incentive_id"]
        for r in filtered
    }
    team_selected = st.selectbox(
        "Choose one to delete",
        options=[""] + list(team_options.keys()),
        key=f"del_team_select_{key_prefix}",
    )
    if team_selected and team_selected in team_options and st.button(
        "Delete this manager incentive", key=f"del_team_btn_{key_prefix}"
    ):
        delete_team_incentive(team_options[team_selected])
        log_audit("DELETE_TEAM_INCENTIVE", "team_incentive", user_id, str(team_options[team_selected]), team_selected)
        st.success("Manager incentive deleted.")
        st.rerun()


def _commission_policy_pdf_candidate_paths() -> list[Path]:
    """Resolve official commission deck: env override, then files under ``policy/``."""
    root = Path(__file__).resolve().parent
    out: list[Path] = []
    env_main = (os.environ.get("COMMISSION_POLICY_PDF") or "").strip()
    if env_main:
        out.append(Path(env_main))
    env_extra = (os.environ.get("COMMISSION_POLICY_PDF_EXTRA") or "").strip()
    if env_extra:
        out.append(Path(env_extra))
    out.extend(
        [
            root / "policy" / "Sales_Commission_Policy.pdf",
            root / "policy" / "Sales Commision PPT.pdf",
            root / "policy" / "Sales Commission PPT.pdf",
        ]
    )
    return out


def resolve_commission_policy_pdf_path() -> Path | None:
    for p in _commission_policy_pdf_candidate_paths():
        try:
            if p.is_file():
                return p
        except OSError:
            continue
    return None


def render_rules_and_policy_page() -> None:
    """Rules hub: structured team policy + optional file text extraction + reference PDF."""
    from policy_visual import render_commission_policy_page

    role = (st.session_state.user or {}).get("role", "")
    prefix = "rules_policy_admin" if role == "ADMIN" else "rules_policy_member"
    render_commission_policy_page(resolve_commission_policy_pdf_path(), key_prefix=prefix)


def render_admin_dashboard():
    """Admin: Metrics = upload + uploads list, then team tabs (SMB / AM / ENT) each with Rep / Manager / Team sub-tabs."""
    from commission_policy import ACCOUNT_MANAGEMENT_TEAM_NAME, ENTERPRISE_TEAM_NAME

    user_id = st.session_state.user["user_id"]
    user = st.session_state.user
    nav = st.session_state.get("admin_sidebar_nav", "Rules & Policy")

    with st.container(border=True):
        hdr_left, hdr_right = st.columns([2.6, 1])
        with hdr_left:
            st.markdown(
                '<h1 class="smb-metrics-dashboard-title">Compensation Tool</h1>',
                unsafe_allow_html=True,
            )
        with hdr_right:
            render_branding_and_settings(user)

    if nav == "Rules & Policy":
        with st.container(border=True):
            render_rules_and_policy_page()
        return

    if nav == "Upload & Deals":
        with st.container(border=True):
            st.caption(
                "**Performance vs goal**, **HubSpot**, and **Excel** use tabs: **SMB**, **Account Management**, and **Enterprise**."
            )
            render_upload_section(user_id, sales_target_team="SMB")
            # Uploads list is hidden — Q1 2026 data is locked. Re-enable with
            # COMP_TOOL_ENABLE_IMPORTS=1 in .env when imports are needed again.
            if (os.environ.get("COMP_TOOL_ENABLE_IMPORTS") or "").strip() in ("1", "true", "yes"):
                st.divider()
                render_uploads_list(user_id, admin=True)
        return
    if nav == "Outbound":
        with st.container(border=True):
            render_outbound_admin(user_id)
        return
    if nav == "User Management":
        with st.container(border=True):
            render_user_management(user_id)
        return
    if nav == "Export":
        with st.container(border=True):
            render_export_section()
        return
    if nav == "Assistant":
        with st.container(border=True):
            render_compensation_assistant_page()
        return

    _metrics_team_map = {
        "SMB": ("SMB", "smb"),
        "AM": (ACCOUNT_MANAGEMENT_TEAM_NAME, "am"),
        "ENT": (ENTERPRISE_TEAM_NAME, "ent"),
    }

    if nav in _metrics_team_map:
        tf_label, kp = _metrics_team_map[nav]
        with st.container():
            with st.container(border=True):
                render_upload_section(user_id, sales_target_team=tf_label)
                # Uploads list is hidden — Q1 2026 data is locked. Re-enable with
                # COMP_TOOL_ENABLE_IMPORTS=1 in .env when imports are needed again.
                if (os.environ.get("COMP_TOOL_ENABLE_IMPORTS") or "").strip() in ("1", "true", "yes"):
                    st.divider()
                    render_uploads_list(user_id, admin=True)
            st.divider()
            t_mgr, t_team = st.tabs(
                ["Manager incentive", "Team view"]
            )
            with t_mgr:
                with st.container(border=True):
                    handled = False
                    if nav in ("SMB", "AM"):
                        _picker_prefix = "ga_smb" if nav == "SMB" else "ga_am"
                        _team_prefix = "smb" if nav == "SMB" else "am"
                        # The picker is now Custom-only — read the From/To directly from session
                        # state. Fall back to Last quarter default if the picker hasn't been
                        # touched yet (first render before the Sales Target tab was opened).
                        _cs = st.session_state.get(f"{_picker_prefix}_close_date_from")
                        _ce = st.session_state.get(f"{_picker_prefix}_close_date_to")
                        _today = datetime.now().date()
                        if _cs is not None and _ce is not None:
                            _start, _end = _resolve_close_date_range(
                                "Custom date range", today=_today, custom_start=_cs, custom_end=_ce
                            )
                            _period_lbl = f"{_start:%b %d, %Y} – {_end:%b %d, %Y}"
                        else:
                            # No selection yet — default to Last quarter
                            _start, _end = _resolve_close_date_range("Last quarter", today=_today)
                            _period_lbl = f"Last quarter ({_start:%b %d, %Y} – {_end:%b %d, %Y})"
                        _month_n = _is_single_month_range(_start, _end)
                        st.caption(
                            f"Period: **{_period_lbl}**. "
                            f"Change the From/To dates on the {nav} Sales Target page to switch periods."
                        )

                        # 1) Single-month range → monthly manager view
                        if _month_n:
                            monthly_pinned = _load_pinned_monthly(_team_prefix, int(_start.year), _month_n)
                            if monthly_pinned is not None:
                                team_lbl_for_mgr = "SMB" if nav == "SMB" else "Account Management"
                                render_pinned_monthly_manager_view(monthly_pinned, team_lbl_for_mgr)
                                handled = True

                        # 2) Range matching a pinned quarter → quarterly manager view
                        if not handled:
                            _yr_pick = int(_start.year)
                            _qt_pick = _quarter_of_month(int(_start.month))
                            qstart, qend = _quarter_start_end(_yr_pick, _qt_pick)
                            if _start == qstart and _end == qend:
                                cand = _load_pinned_quarter(_team_prefix, _yr_pick, _qt_pick)
                                if cand and (cand.get("manager_incentive") or {}):
                                    render_pinned_smb_manager_view(cand)
                                    handled = True

                    if not handled:
                        # ENT does not have a separate manager incentive — Anthony is the only person on the team.
                        # His full commission (base + Q4 pending + Q4 refund clawback) is shown on the Sales Target page.
                        if nav == "ENT":
                            st.info(
                                "ℹ️ **No separate manager incentive for the Enterprise team.** "
                                "Anthony Raymond is the only person on the ENT team. His complete commission breakdown "
                                "— base commission, Q4 pending deals, and refund clawback — is displayed on the **Enterprise Sales Target** tab. "
                                "There is no additional manager-pool incentive."
                            )
                        else:
                            st.caption(f"Team goals and manager incentives — **{tf_label}**.")
                            render_manager_incentive_admin(user_id, tf_label, kp)
            with t_team:
                with st.container(border=True):
                    pinned_for_team = None
                    if nav == "SMB":
                        for _yr in (2026,):
                            for _qt in (1, 2, 3, 4):
                                cand = _load_pinned_smb_quarter(_yr, _qt)
                                if cand and (cand.get("reps") or []):
                                    pinned_for_team = cand
                                    break
                            if pinned_for_team:
                                break
                    if pinned_for_team is not None:
                        render_pinned_smb_team_view(pinned_for_team)
                    else:
                        st.caption(f"Charts and deals — **{tf_label}**.")
                        render_team_view_admin(team_filter=tf_label, key_prefix=kp)
        return


def render_outbound_admin(admin_user_id: int):
    """Log outbound meetings eligible under Q1 2026 policy (NAM + Western Europe)."""
    from commission_policy import (
        OUTBOUND_ELIGIBLE_REGIONS,
        OUTBOUND_MEETING_PAYOUT_NOTE,
        OUTBOUND_MEETING_PAYOUT_ROWS,
        OUTBOUND_MEETING_PAYOUT_TITLE,
        OUTBOUND_POLICY_ELIGIBILITY,
        OUTBOUND_POLICY_LABEL,
        display_label_for_outbound_region,
    )

    st.subheader("Outbound meeting incentives")
    st.caption(f"**{OUTBOUND_POLICY_LABEL}** — {OUTBOUND_POLICY_ELIGIBILITY}")
    if OUTBOUND_MEETING_PAYOUT_ROWS:
        st.markdown(f"**{OUTBOUND_MEETING_PAYOUT_TITLE}**")
        st.dataframe(
            pd.DataFrame(OUTBOUND_MEETING_PAYOUT_ROWS),
            use_container_width=True,
            hide_index=True,
            column_config={
                "meetings": st.column_config.TextColumn("Meetings (count)", width="large"),
                "payout": st.column_config.TextColumn("Payout", width="large"),
            },
        )
        st.caption(OUTBOUND_MEETING_PAYOUT_NOTE)
    st.caption("Enter the incentive amount for each meeting when you record a payout (match the tier above).")

    users = get_all_users_with_teams()
    sales_users = [u for u in users if u.get("role") in ("SALES_REP", "SALES_MANAGER")]

    with st.form("outbound_add_form", clear_on_submit=True):
        if not sales_users:
            st.warning("Add sales reps in User Management before logging outbound meetings.")
        rep_options = {f"{u['full_name']} ({u['email']})": u["user_id"] for u in sales_users}
        rep_label = st.selectbox("Rep", options=[""] + list(rep_options.keys()))
        region_label = st.selectbox(
            "Region",
            options=[lbl for _, lbl in OUTBOUND_ELIGIBLE_REGIONS],
        )
        meeting_date = st.date_input("Meeting date")
        incentive_amt = st.number_input("Incentive amount ($)", min_value=0.0, value=0.0, step=25.0)
        notes = st.text_area("Notes (optional)", height=68)
        submitted = st.form_submit_button("Add outbound meeting")
        if submitted:
            if not rep_label or rep_label not in rep_options:
                st.error("Choose a rep.")
            else:
                code = next(c for c, lbl in OUTBOUND_ELIGIBLE_REGIONS if lbl == region_label)
                oid = insert_outbound_meeting(
                    rep_options[rep_label],
                    code,
                    meeting_date,
                    incentive_amt,
                    notes or None,
                    admin_user_id,
                )
                log_audit("OUTBOUND_ADD", "outbound_meeting", admin_user_id, str(oid), rep_label)
                st.success("Outbound meeting recorded.")
                st.rerun()

    rows = get_all_outbound_meetings()
    if not rows:
        st.info("No outbound meetings logged yet.")
        return

    df = pd.DataFrame(rows)
    if "region" in df.columns:
        df["region"] = df["region"].apply(lambda c: display_label_for_outbound_region(str(c)) or c)
    show_cols = [c for c in ["rep_name", "rep_email", "region", "meeting_date", "incentive_amount", "notes", "created_at"] if c in df.columns]
    df_show = df[show_cols].copy()
    if "incentive_amount" in df_show.columns:
        df_show = _format_df_dollars(df_show, ["incentive_amount"])
    st.dataframe(df_show, use_container_width=True, hide_index=True)

    del_opts = {
        f"#{r['outbound_id']}: {r.get('rep_name', '')} — {r.get('meeting_date')} — {r.get('region', '')}": r[
            "outbound_id"
        ]
        for r in rows
    }
    pick = st.selectbox("Delete a record", options=[""] + list(del_opts.keys()), key="outbound_del_select")
    if pick and pick in del_opts and st.button("Delete selected", key="outbound_del_btn"):
        delete_outbound_meeting(del_opts[pick])
        log_audit("OUTBOUND_DELETE", "outbound_meeting", admin_user_id, str(del_opts[pick]), pick)
        st.success("Deleted.")
        st.rerun()


def render_team_view_admin(team_filter: str | None = None, key_prefix: str = "gv"):
    """Team view: Reps (charts + deal names per rep) and Managers (team charts only, no deal names)."""
    from commission_policy import SMB_TEAM_NAME

    tf = (team_filter or "").strip() or None
    title_suffix = f" — {tf}" if tf else ""
    st.subheader(f"Team view{title_suffix}")
    view_mode = st.radio(
        "View by",
        ["Reps", "Managers (Team)"],
        horizontal=True,
        key=f"team_view_mode_{key_prefix}",
    )

    if view_mode == "Reps":
        incentives = get_rep_incentives()
        if tf:
            incentives = [r for r in (incentives or []) if (r.get("team_name") or "").strip() == tf]
        if not incentives:
            st.info("No rep incentives yet. Finalize an upload to generate." if not tf else f"No rep incentives for **{tf}** yet.")
            return
        ptn = tf if tf else SMB_TEAM_NAME
        _render_rep_incentives_fiesta_analytics(
            incentives,
            primary_team_name=ptn,
            analytics_title=f"{ptn} performance analytics" if tf else None,
            period_key=f"fiesta_rep_period_{key_prefix}",
        )
        st.divider()
        st.subheader("Additional chart views")
        df = pd.DataFrame(incentives)
        # Aggregate by rep (full_name) in case of multiple periods
        agg = df.groupby("full_name", as_index=False).agg({
            "incentive_amount": "sum",
            "total_revenue": "sum",
            "total_paid_amount": "sum",
        }).rename(columns={"incentive_amount": "Incentive ($)", "total_revenue": "Revenue ($)", "total_paid_amount": "Paid ($)"})

        # Different graph types with data labels
        st.caption("**Incentive amount by rep** (bar chart)")
        _chart_bar_with_labels(agg, "full_name", "Incentive ($)", height=320)

        st.caption("**Total revenue by rep** (line chart)")
        _chart_line_with_labels(agg, "full_name", "Revenue ($)", height=320)

        st.caption("**Paid amount by rep** (area chart)")
        _chart_area_with_labels(agg, "full_name", "Paid ($)", height=320)

        # Per-rep view: each rep with metrics and chart with data labels
        st.divider()
        st.subheader("Per rep view")
        for i, row in agg.iterrows():
            rep_name = row["full_name"]
            inc_val = float(row["Incentive ($)"])
            rev_val = float(row["Revenue ($)"])
            paid_val = float(row["Paid ($)"])
            with st.expander(f"**{rep_name}** — Incentive: ${inc_val:,.2f} | Revenue: ${rev_val:,.2f}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.metric("Incentive ($)", _fmt_dollar(inc_val))
                    st.metric("Revenue ($)", _fmt_dollar(rev_val))
                    st.metric("Paid ($)", _fmt_dollar(paid_val))
                with c2:
                    mini_df = pd.DataFrame({"Metric": ["Incentive ($)", "Revenue ($)", "Paid ($)"], "Value": [inc_val, rev_val, paid_val]})
                    mini_df["Label"] = mini_df["Value"].apply(lambda v: f"{float(v):,.0f}")
                    _mord = mini_df["Metric"].astype(str).tolist()
                    base = alt.Chart(mini_df).encode(
                        x=alt.X("Metric:N", title="", sort=_mord),
                        y=alt.Y("Value:Q", title=""),
                        color=_chart_categorical_color("Metric:N", _mord, legend=None),
                    )
                    bars = base.mark_bar(size=36)
                    text = base.mark_text(dy=-8, align="center", fontSize=10, color="#0d47a1").encode(text="Label:N")
                    st.altair_chart(bars + text, use_container_width=True, theme=None)

        # Deals per rep (all deal names for each rep; close date from uploaded file)
        st.divider()
        st.subheader("Deals per rep")
        def _fmt_close_date(dt):
            if dt is None:
                return ""
            return dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)
        deals_raw = get_deals_from_finalized_uploads()
        if tf:
            deals_raw = _filter_deals_by_team_name(deals_raw, tf)
        deals_by_rep = defaultdict(list)
        for d in deals_raw:
            name = d.get("deal_owner_name") or f"User #{d.get('deal_owner_id')}"
            deals_by_rep[name].append({
                "Deal": d.get("deal_name"),
                "Amount": _fmt_dollar(float(d.get("amount") or 0)),
                "Paid": _fmt_dollar(float(d.get("paid_amount") or 0)),
                "Status": d.get("payment_status") or "",
                "Close date": _fmt_close_date(d.get("close_date")),
            })

        for rep_name in sorted(deals_by_rep.keys()):
            rep_deals = deals_by_rep[rep_name]
            with st.expander(f"**{rep_name}** — {len(rep_deals)} deal(s)"):
                st.dataframe(pd.DataFrame(rep_deals), use_container_width=True, hide_index=True)

    else:
        # Managers (Team) — different graph types, no deal names
        incentives = get_team_incentives()
        if tf:
            incentives = [r for r in (incentives or []) if (r.get("team_name") or "").strip() == tf]
        if not incentives:
            st.info(
                "No team incentives yet. Set team goals and finalize an upload to generate."
                if not tf
                else f"No manager incentive rows for **{tf}** yet."
            )
            return
        df = pd.DataFrame(incentives)
        agg = df.groupby("team_name", as_index=False).agg({
            "incentive_amount": "sum",
            "total_team_revenue": "sum",
        }).rename(columns={"incentive_amount": "Incentive ($)", "total_team_revenue": "Revenue ($)"})

        st.caption("**Chitradip's Manager incentive by team** (bar chart)")
        _chart_bar_with_labels(agg, "team_name", "Incentive ($)", height=320)
        st.caption("**Chitradip's Team revenue by team** (pie chart)")
        _chart_pie(agg, "team_name", "Revenue ($)", height=320)

        st.divider()
        st.subheader("Chitradip's Per team view")
        for _, row in agg.iterrows():
            team_name = row["team_name"]
            inc_val = float(row["Incentive ($)"])
            rev_val = float(row["Revenue ($)"])
            with st.expander(f"**{team_name}** — Incentive: {_fmt_dollar(inc_val)} | Revenue: {_fmt_dollar(rev_val)}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.metric("Incentive ($)", _fmt_dollar(inc_val))
                    st.metric("Revenue ($)", _fmt_dollar(rev_val))
                with c2:
                    mini_df = pd.DataFrame({"Metric": ["Incentive ($)", "Revenue ($)"], "Value": [inc_val, rev_val]})
                    mini_df["Label"] = mini_df["Value"].apply(lambda v: f"{float(v):,.0f}")
                    _mord2 = mini_df["Metric"].astype(str).tolist()
                    base = alt.Chart(mini_df).encode(
                        x=alt.X("Metric:N", title="", sort=_mord2),
                        y=alt.Y("Value:Q", title=""),
                        color=_chart_categorical_color("Metric:N", _mord2, legend=None),
                    )
                    bars = base.mark_bar(size=36)
                    text = base.mark_text(dy=-8, align="center", fontSize=10, color="#0d47a1").encode(text="Label:N")
                    st.altair_chart(bars + text, use_container_width=True, theme=None)
        st.caption("Chitradip's manager view: team-level metrics only. Deal names are not shown here.")


def render_member_dashboard():
    user = st.session_state.user
    nav = st.session_state.get("member_sidebar_nav", "Rules & Policy")
    with st.container(border=True):
        m_left, m_right = st.columns([2.6, 1])
        with m_left:
            if nav == "Rules & Policy":
                st.title("Rules and policy")
                st.caption("Commission documentation (read-only)")
            else:
                st.title("Incentive Report")
                st.caption("View your incentives (read-only)")
        with m_right:
            render_branding_and_settings(user)

    if nav == "Rules & Policy":
        with st.container(border=True):
            render_rules_and_policy_page()
        return

    user_id = user["user_id"]
    role = user.get("role", "")
    team_id = user.get("team_id")
    team_name = (user.get("team_name") or "").strip()

    # ---- My Sales Target: filtered pinned Q1 / April / May view for the current user ----
    if nav == "My Sales Target":
        with st.container(border=True):
            # Auto-detect the user's team by matching their name against the pinned JSON
            # rep lists. This works even if user.team_name isn't set correctly.
            _matchers = _get_user_rep_name_matchers(user) or []
            _detected_team = None  # "smb" | "am" | "ent"

            def _user_in_pinned(pinned: dict | None) -> bool:
                if not pinned or not _matchers:
                    return False
                # Check reps + paid_deals + unpaid_deals_overrides
                for r in (pinned.get("reps") or []):
                    if _rep_name_matches_user(r.get("name") or "", _matchers):
                        return True
                for d in (pinned.get("paid_deals") or []):
                    if _rep_name_matches_user(d.get("deal_owner") or "", _matchers):
                        return True
                for d in (pinned.get("unpaid_deals_overrides") or []):
                    if _rep_name_matches_user(d.get("deal_owner") or "", _matchers):
                        return True
                # Also check if the user is the manager
                mgr_name = (pinned.get("manager_incentive") or {}).get("manager_name") or ""
                if _rep_name_matches_user(mgr_name, _matchers):
                    return True
                return False

            # Try SMB (Q1 → April → May), then AM, then ENT.
            for team_prefix, team_key in (("smb", "smb"), ("am", "am"), ("ent", "ent")):
                # Q1 pinned
                q_pinned = _load_pinned_smb_quarter(2026, 1) if team_prefix == "smb" else (
                    _load_pinned_am_quarter(2026, 1) if team_prefix == "am" else None
                )
                if _user_in_pinned(q_pinned):
                    _detected_team = team_key
                    break
                # Monthly
                for _mo in (4, 5):
                    m_pinned = _load_pinned_monthly(team_prefix, 2026, _mo)
                    if _user_in_pinned(m_pinned):
                        _detected_team = team_key
                        break
                if _detected_team:
                    break

            # Fallback to user.team_name if we couldn't auto-detect from pinned files.
            if not _detected_team:
                from commission_policy import ACCOUNT_MANAGEMENT_TEAM_NAME, ENTERPRISE_TEAM_NAME, SMB_TEAM_NAME
                _team_low = (team_name or "").lower()
                if team_name == SMB_TEAM_NAME or _team_low == "smb":
                    _detected_team = "smb"
                elif team_name == ACCOUNT_MANAGEMENT_TEAM_NAME or "account" in _team_low or _team_low == "am":
                    _detected_team = "am"
                elif team_name == ENTERPRISE_TEAM_NAME or "enterprise" in _team_low or _team_low == "ent":
                    _detected_team = "ent"

            if _detected_team == "smb":
                st.markdown("### My SMB Sales Target")
                _render_smb_goal_attainment_table()
            elif _detected_team == "am":
                st.markdown("### My Account Management Sales Target")
                _render_account_management_goal_attainment()
            elif _detected_team == "ent":
                st.markdown("### My Enterprise Sales Target")
                _render_enterprise_goal_attainment()
            else:
                st.info(
                    f"Your team ({team_name or 'unknown'}) is not currently set up with pinned Sales Target data. "
                    "Please contact your administrator to assign you to the SMB, Account Management, or Enterprise roster."
                )
                st.caption(
                    f"Debug: logged-in user = {user.get('full_name') or user.get('email') or user.get('user_id')}, "
                    f"role = {user.get('role')}, team_name = {team_name!r}. "
                    f"Name matchers used for detection: {_matchers or '(none — no full_name set)'}."
                )
        return

    # Rep incentives (own) — with graphical view and deal names for reps
    st.subheader("My Rep Incentives")
    rep_inv = get_rep_incentives(user_id=user_id)
    if rep_inv:
        df = pd.DataFrame(rep_inv)
        df = _enrich_rep_incentives_display(df)
        cols = ["full_name", "team_name", "quota", "quota_achievement_pct", "total_revenue", "total_paid_amount", "incentive_percentage", "incentive_amount", "close_date", "incentive_eligibility", "calculation_period"]
        show = [c for c in cols if c in df.columns]
        base = df[show] if show else df
        df_display = _format_df_dollars(base, ["quota", "total_revenue", "total_paid_amount", "incentive_amount"])
        df_display = _format_pct_display(df_display, ["quota_achievement_pct", "incentive_percentage"])
        st.dataframe(df_display, use_container_width=True)
        # Graphical view for rep: bar of incentive/revenue + my deals
        st.divider()
        st.subheader("My view")
        inc_sum = df["incentive_amount"].astype(float).sum()
        rev_sum = df["total_revenue"].astype(float).sum()
        chart_df = pd.DataFrame({"Metric": ["Incentive ($)", "Revenue ($)"], "Value": [inc_sum, rev_sum]})
        chart_df["Label"] = chart_df["Value"].apply(lambda v: f"{float(v):,.0f}")
        _mv = chart_df["Metric"].astype(str).tolist()
        base = alt.Chart(chart_df).encode(
            x=alt.X("Metric:N", title="", sort=_mv),
            y=alt.Y("Value:Q", title=""),
            color=_chart_categorical_color("Metric:N", _mv, legend=None),
        )
        bars = base.mark_bar(size=40)
        text = base.mark_text(dy=-10, align="center", fontSize=11, color="#0d47a1").encode(text="Label:N")
        st.altair_chart(bars + text, use_container_width=True, theme=None)
        deals_raw = get_deals_from_finalized_uploads()
        def _fmt_date(dt):
            if dt is None:
                return ""
            return dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)
        my_deals = [{"Deal": d.get("deal_name"), "Amount": _fmt_dollar(float(d.get("amount") or 0)), "Paid": _fmt_dollar(float(d.get("paid_amount") or 0)), "Status": d.get("payment_status") or "", "Close date": _fmt_date(d.get("close_date"))} for d in deals_raw if d.get("deal_owner_id") == user_id]
        if my_deals:
            st.caption("My deals (from finalized uploads)")
            st.dataframe(pd.DataFrame(my_deals), use_container_width=True, hide_index=True)
    else:
        st.info("No rep incentives for you yet.")

    # Team incentives (Sales Manager sees their team) — graphical view only, no deal names
    if role == "SALES_MANAGER" and team_id:
        st.subheader("Manager Incentive")
        team_inv = get_team_incentives(team_lead_id=user_id)
        if team_inv:
            df = _enrich_team_incentives_display(pd.DataFrame(team_inv))
            cols = [
                "team_name",
                "total_team_revenue",
                "Team goal achievement %",
                "incentive_percentage",
                "incentive_amount",
                "Payout note",
                "calculation_period",
            ]
            show = [c for c in cols if c in df.columns]
            base_df = df[show] if show else df
            dollar_cols = [c for c in ["total_team_revenue", "incentive_amount"] if c in base_df.columns]
            df_display = _format_df_dollars(base_df, dollar_cols)
            if "Team goal achievement %" in df_display.columns:
                df_display = df_display.copy()
                df_display["Team goal achievement %"] = pd.to_numeric(
                    df_display["Team goal achievement %"], errors="coerce"
                ).fillna(0.0)
            _cfg = {k: v for k, v in _manager_incentive_column_config().items() if k in df_display.columns}
            _df_kw = {"use_container_width": True}
            if _cfg:
                _df_kw["column_config"] = _cfg
            st.dataframe(df_display, **_df_kw)
            st.divider()
            st.subheader("Chitradip's Manager view")
            agg = df.groupby("team_name", as_index=False).agg({"incentive_amount": "sum", "total_team_revenue": "sum"})
            agg = agg.rename(columns={"incentive_amount": "Incentive ($)", "total_team_revenue": "Revenue ($)"})
            _chart_bar_with_labels(agg, "team_name", "Incentive ($)", height=280)
            _chart_line_with_labels(agg, "team_name", "Revenue ($)", height=280)
            st.caption("Chitradip's manager view: team-level metrics. Deal names are not shown.")
        else:
            st.info("No team incentives for your team yet.")


_HUBSPOT_IMPORTS_BY_CTX_KEY = "_compensation_tool_hubspot_imports_by_ctx"


def _hubspot_legacy_keys(ctx: str) -> tuple[str, str, str, str]:
    return (
        f"hubspot_last_fetch_summary_{ctx}",
        f"hubspot_last_fetch_deals_df_{ctx}",
        f"hubspot_validated_deals_{ctx}",
        f"hubspot_upload_name_{ctx}",
    )


def _hubspot_imports_store() -> dict:
    if _HUBSPOT_IMPORTS_BY_CTX_KEY not in st.session_state:
        st.session_state[_HUBSPOT_IMPORTS_BY_CTX_KEY] = {}
    return st.session_state[_HUBSPOT_IMPORTS_BY_CTX_KEY]


def _hubspot_migrate_legacy_into_store() -> None:
    """Copy legacy flat HubSpot keys into the per-team dict so they survive tab switches."""
    store = _hubspot_imports_store()
    for ctx in ("smb", "am", "ent"):
        k_sum, k_df, k_val, k_name = _hubspot_legacy_keys(ctx)
        if ctx in store and store[ctx].get("df") is not None:
            continue
        df = st.session_state.get(k_df)
        if df is not None and hasattr(df, "empty") and not df.empty:
            store[ctx] = {
                "summary": st.session_state.get(k_sum),
                "df": df.copy(),
                "valid": st.session_state.get(k_val),
                "fname": st.session_state.get(k_name),
            }


def _hubspot_store_write(
    ctx: str,
    *,
    summary: str | None,
    df,
    valid=None,
    fname=None,
) -> None:
    """Persist one team's HubSpot fetch in session (separate from other teams' tabs)."""
    store = _hubspot_imports_store()
    df_stored = df.copy() if df is not None and hasattr(df, "copy") else df
    store[ctx] = {
        "summary": summary,
        "df": df_stored,
        "valid": valid,
        "fname": fname,
    }
    k_sum, k_df, k_val, k_name = _hubspot_legacy_keys(ctx)
    if summary is not None:
        st.session_state[k_sum] = summary
    if df_stored is not None:
        st.session_state[k_df] = df_stored
    else:
        st.session_state.pop(k_df, None)
    if valid is not None:
        st.session_state[k_val] = valid
    else:
        st.session_state.pop(k_val, None)
    if fname is not None:
        st.session_state[k_name] = fname
    else:
        st.session_state.pop(k_name, None)


def _hubspot_store_read(ctx: str) -> dict:
    _hubspot_migrate_legacy_into_store()
    store = _hubspot_imports_store()
    k_sum, k_df, k_val, k_name = _hubspot_legacy_keys(ctx)
    entry = dict(store.get(ctx) or {})
    if entry.get("df") is None:
        df = st.session_state.get(k_df)
        if df is not None and hasattr(df, "empty") and not df.empty:
            entry["df"] = df
            entry["summary"] = st.session_state.get(k_sum)
            entry["valid"] = st.session_state.get(k_val)
            entry["fname"] = st.session_state.get(k_name)
    if entry.get("df") is None:
        return {"summary": None, "df": None, "valid": None, "fname": None}
    entry.setdefault("summary", st.session_state.get(k_sum))
    entry.setdefault("valid", st.session_state.get(k_val))
    entry.setdefault("fname", st.session_state.get(k_name))
    return entry


def _hubspot_clear_pending_after_save(ctx: str) -> None:
    store = _hubspot_imports_store()
    if ctx in store:
        store[ctx]["valid"] = None
        store[ctx]["fname"] = None
    _, _, k_val, k_name = _hubspot_legacy_keys(ctx)
    st.session_state.pop(k_val, None)
    st.session_state.pop(k_name, None)


def render_upload_section(user_id: int, sales_target_team: str = "SMB"):
    """
    HubSpot import + Excel upload (shared).

    **Performance vs goal, HubSpot, Excel:** Three tabs — **SMB**, **Account Management**, and **Enterprise**.

    **HubSpot / Excel:** Separate session storage per tab (``_smb`` / ``_am`` / ``_ent`` keys).

    ``sales_target_team`` is kept for API compatibility; upload UI no longer depends on it.
    """
    from hubspot_service import (
        default_deal_year_quarter,
        fetch_and_map_hubspot_deals,
        get_access_token,
    )

    hubspot_auto_token = get_access_token()
    if not hubspot_auto_token:
        st.info(
            "For **HubSpot** (deal import), add `HUBSPOT_ACCESS_TOKEN` to `.env` or paste a token below. "
            "Excel upload works without HubSpot."
        )
        st.text_input(
            "HubSpot Private App access token",
            type="password",
            placeholder="Required for HubSpot deal import",
            key="hubspot_token_manual",
        )
    manual_only = (st.session_state.get("hubspot_token_manual") or "").strip()
    token = hubspot_auto_token or manual_only

    _team_to_ctx = {
        "SMB": "smb",
        "Account Management": "am",
        "Enterprise": "ent",
    }
    _ctx_to_label = {"smb": "SMB", "am": "Account Management", "ent": "Enterprise"}
    _all_contexts = [("smb", "SMB"), ("am", "Account Management"), ("ent", "Enterprise")]
    _selected_ctx = _team_to_ctx.get((sales_target_team or "").strip())
    if _selected_ctx is None:
        _contexts_to_render = _all_contexts
        _tabs = st.tabs([lbl for _c, lbl in _contexts_to_render])
    else:
        _contexts_to_render = [(c, lbl) for c, lbl in _all_contexts if c == _selected_ctx]
        _tabs = [contextlib.nullcontext() for _ in _contexts_to_render]
    for tab, (ctx, _ctx_label_outer) in zip(_tabs, _contexts_to_render):
        with tab:
            _ctx_label = {"smb": "SMB", "am": "Account Management", "ent": "Enterprise"}[ctx]
            if ctx == "smb":
                st.markdown("### SMB Sales Target")
                _render_smb_goal_attainment_table()
                st.caption(
                    "SMB: use **HubSpot** or **Excel** below in this tab."
                )
            elif ctx == "am":
                st.markdown("### Account Management Sales Target")
                _render_account_management_goal_attainment()
            else:
                st.markdown("### Enterprise Sales Target")
                _render_enterprise_goal_attainment()

            # HubSpot import / Excel upload UI is gated behind COMP_TOOL_ENABLE_IMPORTS.
            # When imports are disabled (default) we skip the entire section silently —
            # no banner is shown to keep the page clean.
            _imports_enabled = (os.environ.get("COMP_TOOL_ENABLE_IMPORTS") or "").strip() in ("1", "true", "yes")
            if not _imports_enabled:
                continue

            st.divider()
            st.markdown("### Connect to HubSpot (fetch deals)")
            st.markdown(f"#### {_ctx_label}")
            k_sum = f"hubspot_last_fetch_summary_{ctx}"
            k_df = f"hubspot_last_fetch_deals_df_{ctx}"
            k_val = f"hubspot_validated_deals_{ctx}"
            k_name = f"hubspot_upload_name_{ctx}"

            with st.expander(
                f"Import from HubSpot — {_ctx_label}",
                expanded=(ctx == "smb" and not bool(hubspot_auto_token)),
            ):
                st.caption(
                    "Set **close date** (year + quarter), optional **team** and **deal owner**, **deal stage**, and **payment status**, then **Fetch**. "
                    "**Team** narrows owners to a roster (name match in HubSpot); leave **deal owner** empty to fetch for that whole team. "
                    "**Any stage** = closed-won deals with close date in that quarter. "
                    "**A specific stage** = closed-won stages use close date; other stages use **created** date in that quarter. "
                    "**Payment status** is read from your HubSpot deal property (see `HUBSPOT_PAYMENT_STATUS_PROPERTY` in `.env.example`). "
                    "Only deals whose **HubSpot owner email** matches an **active user** are kept for **Save as draft**."
                )

                _def_y, _def_q = default_deal_year_quarter()
                col_hy, col_hq = st.columns(2)
                with col_hy:
                    hub_close_year = st.number_input(
                        "Close date — year",
                        min_value=2020,
                        max_value=2035,
                        value=int(_def_y),
                        step=1,
                        help="HubSpot deal property **Close date** must fall in this calendar year.",
                        key=f"hubspot_close_year_{ctx}",
                    )
                with col_hq:
                    hub_close_quarter = st.selectbox(
                        "Close date — quarter",
                        options=[1, 2, 3, 4],
                        index=max(0, min(3, int(_def_q) - 1)),
                        format_func=lambda q: {
                            1: "Q1 (Jan–Mar)",
                            2: "Q2 (Apr–Jun)",
                            3: "Q3 (Jul–Sep)",
                            4: "Q4 (Oct–Dec)",
                        }[q],
                        key=f"hubspot_close_quarter_{ctx}",
                    )

                owners_cache = []
                stages_cache = []
                if token:
                    try:
                        ck = _hubspot_cache_key(token)
                        owners_cache = _cached_hubspot_owners(ck, token)
                        stages_cache = _cached_hubspot_stages(ck, token)
                        _cached_hubspot_payment_options(ck, token)
                    except Exception as e:
                        st.warning(f"Could not load HubSpot owners or deal stages (check token scopes): {e}")

                owner_choices: list[str] = []
                owner_id_by_label: dict[str, str] = {}
                for o in owners_cache:
                    oid = str(o.get("id", ""))
                    em = (o.get("email") or "").strip()
                    fn = f"{(o.get('firstName') or '')} {(o.get('lastName') or '')}".strip()
                    lbl = (f"{fn} ({em})" if em else (fn or oid)).strip() or oid
                    owner_choices.append(lbl)
                    owner_id_by_label[lbl] = oid

                stage_labels = ["Any stage"]
                stage_id_by_index = [None]
                stage_is_won_by_index = [None]
                for s in stages_cache:
                    stage_labels.append(s.get("label") or s.get("stage_id", ""))
                    stage_id_by_index.append(s.get("stage_id"))
                    stage_is_won_by_index.append(s.get("is_closed_won_stage", True))

                payment_choices = [{"kind": "any", "label": "Any payment status", "hubspot_value": None, "norm": None}]
                if token:
                    try:
                        ckp = _hubspot_cache_key(token)
                        hs_pay = _cached_hubspot_payment_options(ckp, token)
                        for o in hs_pay:
                            payment_choices.append(
                                {
                                    "kind": "hubspot",
                                    "label": o.get("label") or o.get("value"),
                                    "hubspot_value": o.get("value"),
                                    "norm": None,
                                }
                            )
                    except Exception:
                        pass
                if len(payment_choices) == 1:
                    payment_choices.extend(
                        [
                            {"kind": "norm", "label": "PAID only", "hubspot_value": None, "norm": "PAID"},
                            {"kind": "norm", "label": "UNPAID only", "hubspot_value": None, "norm": "UNPAID"},
                            {"kind": "norm", "label": "PARTIALLY_PAID only", "hubspot_value": None, "norm": "PARTIALLY_PAID"},
                        ]
                    )

                hub_team_filter = st.selectbox(
                    "Team",
                    ["Any", "SMB", "Enterprise", "Account Management"],
                    key=f"hubspot_team_filter_{ctx}",
                    disabled=not token,
                    help="Restrict HubSpot owners to a team roster (substring match on owner name/email). **Any** = no team filter.",
                )
                filter_labels = filter_hubspot_owner_labels_for_team(
                    owner_choices,
                    hub_team_filter,
                    extra_tokens=_hubspot_team_extra_tokens_from_users(hub_team_filter),
                )
                if token and hub_team_filter != "Any" and not filter_labels:
                    st.warning(
                        "No HubSpot owners matched this team’s roster. Edit name tokens in `policy/hubspot_team_owner_filters.json` "
                        "so they appear in HubSpot owner labels, or pick **Any**."
                    )

                col_own, col_stg, col_pay = st.columns(3)
                with col_own:
                    picked_owners = st.multiselect(
                        "Deal owner",
                        options=filter_labels if filter_labels else owner_choices,
                        default=[],
                        key=f"hubspot_filter_owners_multi_{ctx}_{hub_team_filter}",
                        disabled=not token,
                        help="Select one or more HubSpot owners, or leave empty to use **all owners in the team** (if a team is selected) or **all owners** (if Team is **Any**).",
                    )
                with col_stg:
                    is_ = st.selectbox(
                        "Deal stage",
                        options=list(range(len(stage_labels))),
                        format_func=lambda i: stage_labels[i],
                        key=f"hubspot_filter_stage_idx_{ctx}",
                        disabled=not token,
                        help="**Any stage** = all closed-won in the quarter. **One stage** = that stage only; "
                        "open stages filter by **created** date in the quarter.",
                    )
                with col_pay:
                    ipay = st.selectbox(
                        "Payment status",
                        options=list(range(len(payment_choices))),
                        format_func=lambda i: payment_choices[i]["label"],
                        key=f"hubspot_filter_payment_idx_{ctx}",
                        disabled=not token,
                        help="When options load from HubSpot, filtering uses your property’s **stored values** (exact match). "
                        "Requires Private App scope **crm.schemas.read** (or equivalent) to read deal properties.",
                    )
                eff_labels = filter_labels if filter_labels else owner_choices
                if picked_owners:
                    filter_owner_ids = [owner_id_by_label[l] for l in picked_owners if l in owner_id_by_label]
                elif hub_team_filter != "Any":
                    if eff_labels:
                        filter_owner_ids = [owner_id_by_label[l] for l in eff_labels if l in owner_id_by_label]
                    else:
                        # Explicit empty list: no HubSpot owners matched the team roster — do not fetch all deals.
                        filter_owner_ids = []
                else:
                    filter_owner_ids = None
                filter_stage_id = stage_id_by_index[is_] if is_ is not None else None
                filter_stage_is_closed_won = stage_is_won_by_index[is_] if is_ is not None else None
                _pch = payment_choices[ipay] if ipay is not None and ipay < len(payment_choices) else payment_choices[0]
                filter_payment_hubspot_value = _pch.get("hubspot_value") if _pch.get("kind") == "hubspot" else None
                filter_payment_status = _pch.get("norm") if _pch.get("kind") == "norm" else None

                if st.button("Fetch deals from HubSpot", key=f"hubspot_fetch_btn_{ctx}"):
                    if not token:
                        st.error(
                            "Set `HUBSPOT_ACCESS_TOKEN` in a `.env` file (project folder), environment variable, or `.streamlit/secrets.toml`, "
                            "or enter your token in the field above."
                        )
                    else:
                        users = get_all_users_with_teams(active_only=False)
                        allowed_emails = {u["email"].strip().lower() for u in users if u.get("email")}
                        # Empty set {} is truthy but would exclude every deal — use None = no email filter
                        allowed_filter = allowed_emails if allowed_emails else None
                        if not allowed_emails:
                            st.info(
                                "No users in **User Management** yet — HubSpot deals still load below. "
                                "Add active users (matching HubSpot owner emails) to **save as draft**."
                            )
                        try:
                            raw, hub_stats, all_closed = fetch_and_map_hubspot_deals(
                                token,
                                allowed_emails=allowed_filter,
                                year=int(hub_close_year),
                                quarter=int(hub_close_quarter),
                                hubspot_owner_ids=filter_owner_ids,
                                deal_stage_id=filter_stage_id,
                                stage_is_closed_won=filter_stage_is_closed_won,
                                payment_status_filter=filter_payment_status,
                                payment_hubspot_value_filter=filter_payment_hubspot_value,
                            )
                        except Exception as e:
                            st.error(f"HubSpot request failed: {e}")
                        else:
                            # Count rows actually shown (after owner/stage mapping); total_in_hubspot is pre-map API count
                            n_fetched = len(all_closed) if all_closed else 0
                            if hub_stats.get("total_in_hubspot", 0) > 0 and n_fetched == 0:
                                st.warning(
                                    "HubSpot returned deals, but **none appear in the table** with the current filters. "
                                    "Try **clearing owner selection**, **Any payment status**, or confirm **User Management** emails match HubSpot owners; "
                                    "for **non–closed-won** stages, deals must have **created** date in the selected quarter."
                                )
                            if picked_owners:
                                if len(picked_owners) == 1:
                                    first_name = (picked_owners[0].split("(")[0].strip().split() or [""])[0]
                                    if first_name:
                                        summary_msg = f"**{first_name}** — **{n_fetched}** deals from HubSpot."
                                    else:
                                        summary_msg = f"Fetched **{n_fetched}** deals from HubSpot."
                                else:
                                    summary_msg = f"**{len(picked_owners)} owners** — **{n_fetched}** deals from HubSpot."
                            else:
                                if hub_team_filter != "Any":
                                    summary_msg = f"**{hub_team_filter}** — **{n_fetched}** deals from HubSpot."
                                else:
                                    summary_msg = f"Fetched **{n_fetched}** deals from HubSpot."
                            rows_hub = []
                            if all_closed:
                                for d in all_closed:
                                    cd = d.get("close_date")
                                    if cd is not None and hasattr(cd, "isoformat"):
                                        cd_str = cd.isoformat()
                                    else:
                                        cd_str = str(cd or "")
                                    rows_hub.append(
                                        {
                                            "Deal name": (d.get("deal_name") or "").strip() or "—",
                                            "Deal owner": (d.get("deal_owner") or "").strip() or "—",
                                            "Amount": d.get("amount", 0),
                                            "Close date": cd_str,
                                            "Payment status": (d.get("payment_status_label") or d.get("payment_status") or "UNPAID"),
                                        }
                                    )
                            df_out = pd.DataFrame(rows_hub) if rows_hub else None
                            email_to_team = {u["email"].strip().lower(): u["team_name"] for u in users}
                            parsed_deals = []
                            parse_errors = []
                            for i, d in enumerate(raw):
                                row_num = i + 2
                                owner = (d.get("deal_owner") or "").strip()
                                if not owner:
                                    parse_errors.append(ValidationError(row_num, "Deal Owner", "HubSpot deal has no owner"))
                                    continue
                                team = email_to_team.get(owner.lower()) if owner else None
                                if not team:
                                    parse_errors.append(ValidationError(row_num, "Deal Owner", f"Owner '{owner}' not found in Compensation Tool users. Add them in User Management (email must match HubSpot)."))
                                    continue
                                parsed_deals.append(
                                    ParsedDeal(
                                        deal_name=d.get("deal_name", ""),
                                        deal_owner=owner,
                                        amount=float(d.get("amount", 0) or 0),
                                        paid_amount=float(d.get("paid_amount", 0) or 0),
                                        payment_status=(d.get("payment_status") or "UNPAID").strip(),
                                        team=team,
                                        close_date=d.get("close_date"),
                                        incentive_eligibility=(d.get("incentive_eligibility") or "Eligible").strip(),
                                        license_resale_exclusion=False,
                                    )
                                )
                            parsed = ParseResult(deals=parsed_deals, errors=parse_errors)
                            valid, db_errors = validate_against_db(parsed)
                            all_errors = parsed.errors + db_errors
                            if all_errors:
                                st.error("Validation errors:")
                                for e in all_errors[:20]:
                                    st.write(f"Row {e.row}, {e.column}: {e.message}")
                                if len(all_errors) > 20:
                                    st.write(f"... and {len(all_errors) - 20} more")
                            valid_out = None
                            fname_out = None
                            if valid:
                                valid_out = valid
                                fname_out = f"HubSpot_import_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
                                st.success(f"Fetched and validated {len(valid)} deals from HubSpot. Click 'Save as draft' below to add them.")
                            _hubspot_store_write(
                                ctx,
                                summary=summary_msg,
                                df=df_out,
                                valid=valid_out,
                                fname=fname_out,
                            )

            # Last HubSpot fetch stays visible even when the deal import expander is collapsed
            _hub_ent = _hubspot_store_read(ctx)
            _hub_df = _hub_ent.get("df")
            if _hub_df is not None and not _hub_df.empty:
                if _hub_ent.get("summary"):
                    st.info(_hub_ent["summary"])
                st.caption(
                    f"Last HubSpot fetch for **{_ctx_label}** in this view. "
                    "All closed won deals from this fetch. Draft save only includes owners who exist in User Management."
                )
                st.dataframe(_hub_df, use_container_width=True, hide_index=True)

            if _hub_ent.get("valid"):
                valid = _hub_ent["valid"]
                fname = _hub_ent.get("fname") or "HubSpot_import.xlsx"
                st.info(f"Ready to save {len(valid)} deals from HubSpot as draft.")
                if st.button("Save as draft (HubSpot)", key=f"hubspot_save_draft_btn_{ctx}"):
                    upload_id = create_upload(user_id, fname, len(valid))
                    deals = [
                        (
                            d.deal_name,
                            d.user_id,
                            d.team_id,
                            d.amount,
                            d.payment_status,
                            upload_id,
                            d.paid_amount,
                            getattr(d, "close_date", None),
                            getattr(d, "incentive_eligibility", "Eligible"),
                            bool(getattr(d, "license_resale_exclusion", False)),
                        )
                        for d in valid
                    ]
                    insert_deals(deals)
                    log_audit("UPLOAD", "excel_upload", user_id, str(upload_id), f"HubSpot: {len(valid)} deals")
                    _hubspot_clear_pending_after_save(ctx)
                    if "confirm_del_all_uploads" in st.session_state:
                        del st.session_state["confirm_del_all_uploads"]
                    st.success(f"Saved as draft. Upload ID: {upload_id}. You can finalize it in the Uploads list below.")
                    st.rerun()

            st.divider()
            st.markdown("### Upload Excel from your computer")
            st.caption("Drafts are stored independently per tab.")
            st.markdown(f"#### {_ctx_label}")
            sample = create_sample_excel()
            st.download_button(
                "Download sample template",
                data=sample,
                file_name="deal_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"excel_sample_dl_{ctx}",
            )
            uploaded = st.file_uploader("Choose Excel file", type=["xlsx", "xls"], key=f"excel_file_uploader_{ctx}")
            if uploaded:
                content = uploaded.read()
                parsed = parse_excel(content, uploaded.name)
                valid, errors = validate_against_db(parsed)

                if errors:
                    st.error("Validation errors:")
                    for e in errors[:20]:
                        st.write(f"Row {e.row}, {e.column}: {e.message}")
                    if len(errors) > 20:
                        st.write(f"... and {len(errors) - 20} more")

                if valid:
                    st.success(f"Validated {len(valid)} deals. Ready to save as draft.")
                    if st.button("Save as Draft", key=f"upload_save_draft_btn_{ctx}"):
                        upload_id = create_upload(user_id, uploaded.name, len(valid))
                        deals = [
                            (
                                d.deal_name,
                                d.user_id,
                                d.team_id,
                                d.amount,
                                d.payment_status,
                                upload_id,
                                d.paid_amount,
                                getattr(d, "close_date", None),
                                getattr(d, "incentive_eligibility", "Eligible"),
                                bool(getattr(d, "license_resale_exclusion", False)),
                            )
                            for d in valid
                        ]
                        insert_deals(deals)
                        log_audit("UPLOAD", "excel_upload", user_id, str(upload_id), f"{len(valid)} deals")
                        if "confirm_del_all_uploads" in st.session_state:
                            del st.session_state["confirm_del_all_uploads"]
                        st.success(f"Saved as draft. Upload ID: {upload_id}. Previous uploads are kept.")
                        st.rerun()


def render_uploads_list(user_id: int, admin: bool = False):
    st.subheader("Uploads")
    uploads = get_uploads_for_user(user_id, admin=admin)
    if not uploads:
        st.info("No uploads yet.")
        return

    for u in uploads:
        with st.expander(f"Upload {u['upload_id']}: {u['file_name']} – {u['upload_status']}"):
            st.write(f"Records: {u.get('records_processed', 0)} | By: {u.get('uploaded_by_name', 'N/A')} | {u.get('uploaded_at')}")
            if u["upload_status"] == "DRAFT":
                deals = get_deals_by_upload(u["upload_id"])
                if deals:
                    df = pd.DataFrame(deals)
                    cols = [
                        "deal_name",
                        "deal_owner_name",
                        "team_name",
                        "amount",
                        "paid_amount",
                        "payment_status",
                        "close_date",
                    ]
                    df_show = df[[c for c in cols if c in df.columns]].copy()
                    df_show = _prepare_upload_deals_display(df_show)
                    df_show = _format_df_dollars(df_show, ["amount", "paid_amount"])
                    st.dataframe(df_show, use_container_width=True)
                col_fin, col_del = st.columns(2)
                with col_fin:
                    if st.button("Finalize", key=f"fin_{u['upload_id']}"):
                        from incentive_engine import compute_and_store_incentives

                        sync_deal_paid_amounts_from_status(u["upload_id"])
                        period = datetime.now().strftime("%b %Y")
                        update_upload_status(u["upload_id"], "FINALIZED", u.get("records_processed"))
                        n_rep, n_team = compute_and_store_incentives(u["upload_id"], period)
                        log_audit("FINALIZE", "excel_upload", user_id, str(u["upload_id"]), f"Rep: {n_rep}, Team: {n_team}")
                        st.success(f"Finalized! Generated {n_rep} rep incentives, {n_team} team incentives")
                        st.rerun()
                with col_del:
                    if st.button("Delete upload", key=f"del_{u['upload_id']}"):
                        delete_upload(u["upload_id"])
                        log_audit("DELETE_UPLOAD", "excel_upload", user_id, str(u["upload_id"]), u.get("file_name"))
                        st.success("Upload deleted.")
                        st.rerun()
            else:
                deals = get_deals_by_upload(u["upload_id"])
                if deals:
                    df = pd.DataFrame(deals)
                    cols = [
                        "deal_name",
                        "deal_owner_name",
                        "team_name",
                        "amount",
                        "paid_amount",
                        "payment_status",
                        "close_date",
                    ]
                    df_show = df[[c for c in cols if c in df.columns]].copy()
                    df_show = _prepare_upload_deals_display(df_show)
                    df_show = _format_df_dollars(df_show, ["amount", "paid_amount"])
                    st.dataframe(df_show, use_container_width=True)
                if st.button("Delete upload", key=f"del_{u['upload_id']}"):
                    delete_upload(u["upload_id"])
                    log_audit("DELETE_UPLOAD", "excel_upload", user_id, str(u["upload_id"]), u.get("file_name"))
                    st.success("Upload deleted. Note: Rep/Team incentives already generated from this upload are not removed.")
                    st.rerun()

    st.divider()
    st.caption("Delete all uploads and their deal data. Rep/Team incentives are not removed.")
    confirm_del_all = st.checkbox("I confirm I want to delete all uploads and their deals", key="confirm_del_all_uploads")
    if confirm_del_all and st.button("Delete all uploads", type="secondary", key="del_all_uploads"):
        n = delete_all_uploads()
        log_audit("DELETE_ALL_UPLOADS", "excel_upload", user_id, None, str(n))
        st.success(f"Deleted all uploads ({n}) and their deals.")
        st.rerun()


def _clear_add_user_form():
    """Clear User Management form fields from session state so they reset after add."""
    for key in ("um_full_name", "um_email", "um_password", "um_role", "um_team"):
        if key in st.session_state:
            del st.session_state[key]


def render_user_management(admin_id: int):
    """Admin: Add Sales Reps and Sales Managers."""
    st.subheader("User Management")
    st.caption(
        "Add users for Deal Owner mapping in Excel uploads. "
        "Assign **Team** and **Role**; compensation subgroups (if any) are managed in the database or policy."
    )

    teams = get_all_teams()
    team_options = {t["team_name"]: t["team_id"] for t in teams}
    team_list = [""] + list(team_options.keys())

    with st.form("add_user_form"):
        full_name = st.text_input("Full Name", placeholder="John Doe", key="um_full_name")
        email = st.text_input("Email", placeholder="john@cloudfuze.com", key="um_email")
        password = st.text_input("Password", type="password", key="um_password")
        role = st.selectbox("Role", ["SALES_REP", "SALES_MANAGER"], key="um_role")
        team_name = st.selectbox("Team", options=team_list, key="um_team")
        submitted = st.form_submit_button("Add User")
        if submitted:
            if not all([full_name, email, password]):
                st.error("Fill all fields")
                return
            team_id = team_options.get(team_name) if team_name else None
            try:
                pw_hash = hash_password(password)
                create_user(full_name, email, pw_hash, role, team_id, compensation_group=None)
                log_audit("CREATE_USER", "user", admin_id, None, full_name)
                _clear_add_user_form()
                st.success(f"Added {full_name}")
                st.rerun()
            except Exception as e:
                st.error(f"Failed: {e}")

    st.divider()
    st.subheader("Existing Users")
    users = get_all_users_with_teams(active_only=False)
    if users:
        df = pd.DataFrame(users)
        display_cols = ["user_id", "full_name", "email", "role", "team_name", "is_active"]
        if "compensation_group" in df.columns:
            display_cols.append("compensation_group")
        for _hq in ("hubspot_quota_usd", "hubspot_quota_period"):
            if _hq in df.columns:
                display_cols.append(_hq)
        _um_df = df[[c for c in display_cols if c in df.columns]].copy()
        if "hubspot_quota_usd" in _um_df.columns:
            _um_df = _format_df_dollars(_um_df, ["hubspot_quota_usd"])
        from commission_policy import (
            ACCOUNT_MANAGEMENT_TEAM_NAME,
            AM_QUOTA_ACHIEVEMENT_ENABLED,
            SMB_TEAM_NAME,
            am_individual_quota_usd_for_rep,
            smb_individual_quota_usd_for_rep,
        )

        period_um = datetime.now().strftime("%b %Y")
        eff_map = {}
        eff_am_map = {}
        for u in users:
            uid = u.get("user_id")
            if (u.get("team_name") or "").strip() != SMB_TEAM_NAME or u.get("role") != "SALES_REP":
                eff_map[uid] = None
            elif (u.get("compensation_group") or "").strip().upper() == "SMB_CHITRADIP":
                eff_map[uid] = None
            else:
                v = smb_individual_quota_usd_for_rep(
                    u.get("compensation_group"),
                    u.get("hubspot_quota_usd"),
                    full_name=u.get("full_name"),
                    email=u.get("email"),
                    calculation_period=period_um,
                )
                eff_map[uid] = v if v > 0 else None
            if (
                (u.get("team_name") or "").strip() == ACCOUNT_MANAGEMENT_TEAM_NAME
                and u.get("role") == "SALES_REP"
                and AM_QUOTA_ACHIEVEMENT_ENABLED
            ):
                av = am_individual_quota_usd_for_rep(
                    u.get("hubspot_quota_usd"),
                    full_name=u.get("full_name"),
                    email=u.get("email"),
                    calculation_period=period_um,
                )
                eff_am_map[uid] = av if av > 0 else None
            else:
                eff_am_map[uid] = None
        _um_df["effective_smb_target_usd"] = _um_df["user_id"].map(lambda x: eff_map.get(x))
        _um_df["effective_am_target_usd"] = _um_df["user_id"].map(lambda x: eff_am_map.get(x))
        _um_df = _format_df_dollars(_um_df, ["effective_smb_target_usd", "effective_am_target_usd"])
        st.caption(
            f"**Effective SMB target** = same resolution as incentives for **{period_um}** (SMB quota list in policy when enabled, else HubSpot / policy Group A/B defaults). "
            f"**Effective AM target** = policy-named quotas / overrides / HubSpot when AM quota rules are enabled."
        )
        st.dataframe(_um_df, use_container_width=True)

        st.divider()

        st.subheader("Edit user")
        edit_labels = {f"{u['full_name']} ({u['email']}) — ID {u['user_id']}": u for u in users}
        edit_pick = st.selectbox("Select user to edit", options=[""] + list(edit_labels.keys()), key="um_edit_pick")
        if edit_pick and edit_pick in edit_labels:
            eu = edit_labels[edit_pick]
            euid = int(eu["user_id"])
            _role_opts = ["ADMIN", "SALES_REP", "SALES_MANAGER"]
            _role_i = _role_opts.index(eu["role"]) if eu.get("role") in _role_opts else 0
            _team_i = team_list.index(eu["team_name"]) if eu.get("team_name") in team_list else 0
            with st.form(f"edit_user_form_{euid}"):
                e_name = st.text_input("Full name", value=eu.get("full_name") or "", key=f"e_name_{euid}")
                e_email = st.text_input("Email", value=eu.get("email") or "", key=f"e_email_{euid}")
                e_role = st.selectbox("Role", _role_opts, index=_role_i, key=f"e_role_{euid}")
                e_team = st.selectbox("Team", team_list, index=_team_i, key=f"e_team_{euid}")
                e_pw = st.text_input("New password (optional)", type="password", key=f"e_pw_{euid}", help="Leave blank to keep the current password.")
                save_edit = st.form_submit_button("Save changes")
                if save_edit:
                    if not e_name or not e_email:
                        st.error("Full name and email are required.")
                    elif not validate_email(e_email.strip()):
                        st.error("Invalid email format.")
                    elif e_pw and len(e_pw) < 6:
                        st.error("Password must be at least 6 characters if set.")
                    else:
                        tid = team_options.get(e_team) if e_team else None
                        if (e_team or "").strip() == "SMB":
                            smb_cg = eu.get("compensation_group")
                        else:
                            smb_cg = None
                        try:
                            update_user_profile(
                                euid,
                                e_name.strip(),
                                e_email.strip(),
                                e_role,
                                team_id=tid,
                                compensation_group=smb_cg,
                                password_hash=hash_password(e_pw) if e_pw else None,
                            )
                            log_audit("UPDATE_USER", "user", admin_id, str(euid), e_email.strip())
                            st.success("User updated.")
                            st.rerun()
                        except Exception as ex:
                            st.error(str(ex))

    else:
        st.info("No users yet.")


def render_export_section():
    st.subheader("Export Reports")
    rep_inv = get_rep_incentives()
    team_inv = get_team_incentives()
    if rep_inv or team_inv:
        format_ = st.radio("Format", ["Excel", "CSV"])
        if rep_inv:
            df_rep = _enrich_rep_incentives_display(pd.DataFrame(rep_inv))
            if format_ == "Excel":
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    df_rep.to_excel(w, sheet_name="Rep Incentives", index=False)
                    if team_inv:
                        _enrich_team_incentives_display(pd.DataFrame(team_inv)).to_excel(
                            w, sheet_name="Manager Incentive", index=False
                        )
                st.download_button("Download Excel", data=buf.getvalue(), file_name="incentives.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.download_button("Download Rep CSV", data=df_rep.to_csv(index=False), file_name="rep_incentives.csv", mime="text/csv")
        if team_inv and format_ == "CSV":
            st.download_button(
                "Download Manager Incentive CSV",
                data=_enrich_team_incentives_display(pd.DataFrame(team_inv)).to_csv(index=False),
                file_name="manager_incentive.csv",
                mime="text/csv",
            )
    else:
        st.info("No incentives to export. Finalize an upload first.")


def _dashboard_theme_css():
    """Light theme for all pages after login: white background, black text, Amasis MT Pro font (admin + member)."""
    return """
    <style>
    h1.smb-metrics-dashboard-title {
        font-family: "Aptos", "Segoe UI", "Segoe UI Variable", system-ui, sans-serif !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
        margin: 0 0 0.35rem 0 !important;
        padding: 0 !important;
        line-height: 1.2 !important;
        background: linear-gradient(110deg, #0d47a1 0%, #6a1b9a 45%, #00838f 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        color: #1565c0 !important;
    }
    html, body, .stApp, .stApp .main, section.main, [data-testid="stAppViewContainer"] {
        font-family: "Amasis MT Pro", "Amasis MT", Georgia, "Times New Roman", serif !important;
    }
    .goal-attainment-wrap, .goal-attainment-wrap * {
        font-family: "Aptos", "Segoe UI", "Segoe UI Variable", system-ui, sans-serif !important;
        font-size: 14px !important;
    }
    .goal-attainment-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        margin-top: 0.75rem;
    }
    .goal-attainment-table th {
        text-align: left;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
        padding: 14px 18px;
        border-bottom: 3px solid #7e57c2;
        background: linear-gradient(180deg, #e3f2fd 0%, #e1bee7 100%);
        color: #4a148c !important;
    }
    .goal-attainment-table th.ga-th-goal { background: linear-gradient(180deg, #fff8e1 0%, #ffe0b2 100%); color: #e65100 !important; }
    .goal-attainment-table td {
        padding: 22px 18px 26px 18px;
        border-bottom: 1px solid #e0e7ea;
        vertical-align: top;
    }
    .ga-team-title { font-weight: 700; color: #1a1a1a !important; }
    .ga-team-sub { color: #546e7a !important; font-weight: 400; margin-top: 6px; line-height: 1.45; }
    .ga-rep-line { display: flex; align-items: center; gap: 0.85rem; }
    .ga-avatar {
        width: 38px;
        height: 38px;
        border-radius: 50%;
        background: linear-gradient(145deg, #80deea 0%, #ce93d8 55%, #90caf9 100%);
        color: #4a148c !important;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        flex-shrink: 0;
        box-shadow: 0 2px 8px rgba(126, 87, 194, 0.25);
    }
    .ga-rep-name { color: #6a1b9a !important; font-weight: 500; }
    .ga-bar-row {
        display: flex;
        align-items: center;
        gap: 1.1rem;
        margin-bottom: 10px;
    }
    .ga-bar-track {
        flex: 1;
        min-width: 100px;
        height: 12px;
        background: #e8e8e8;
        border-radius: 6px;
        overflow: hidden;
    }
    .ga-bar-fill { height: 100%; background: linear-gradient(90deg, #26c6da, #7e57c2); border-radius: 6px; }
    .ga-pct {
        min-width: 3.25rem;
        text-align: right;
        font-weight: 600;
        color: #1a1a1a !important;
        padding-left: 10px;
        flex-shrink: 0;
    }
    .ga-subline {
        margin-top: 0;
        padding-top: 2px;
        color: #455a64 !important;
        font-weight: 400;
        line-height: 1.5;
    }
    .ga-bullet-section { margin-bottom: 0.5rem; }
    .ga-bullet-section-title {
        font-weight: 600 !important;
        font-size: 15px !important;
        color: #1a1a1a !important;
        margin: 0 0 1rem 0 !important;
    }
    .ga-bullet-row {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem 1.5rem;
        align-items: flex-start;
        margin-bottom: 1.6rem;
    }
    .ga-bullet-left { flex: 0 0 min(240px, 100%); max-width: 300px; }
    .ga-bullet-left-rep .ga-rep-line { align-items: flex-start !important; }
    .ga-bullet-title { font-weight: 700 !important; font-size: 15px !important; color: #1a1a1a !important; }
    .ga-bullet-sub { font-size: 13px !important; color: #546e7a !important; margin-top: 4px !important; line-height: 1.45 !important; }
    .ga-bullet-right { flex: 1 1 260px; min-width: min(100%, 220px); }
    .ga-bullet-legend {
        display: flex;
        justify-content: space-between;
        gap: 8px;
        font-size: 11px !important;
        color: #78909c !important;
        margin-bottom: 6px;
    }
    .ga-leg { flex: 1; text-align: center; }
    .ga-bullet-track-wrap { position: relative; height: 42px; margin-top: 2px; width: 100%; }
    .ga-bullet-zones {
        position: absolute;
        left: 0;
        right: 0;
        top: 10px;
        height: 22px;
        display: flex;
        border-radius: 5px;
        overflow: hidden;
    }
    .ga-bz { display: block; height: 100%; min-width: 0; }
    .ga-bz-a { background: rgba(13, 71, 161, 0.42); }
    .ga-bz-b { background: rgba(25, 118, 210, 0.36); }
    .ga-bz-c { background: rgba(100, 181, 246, 0.4); }
    .ga-bz-d { background: rgba(187, 222, 251, 0.5); }
    .ga-bullet-bar {
        position: absolute;
        left: 0;
        top: 12px;
        height: 18px;
        background: #0d47a1;
        border-radius: 4px;
        z-index: 2;
        min-width: 2px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.15);
    }
    .ga-bullet-marker {
        position: absolute;
        top: 5px;
        height: 32px;
        width: 3px;
        background: #1565c0;
        margin-left: -1.5px;
        z-index: 3;
        border-radius: 1px;
    }
    .ga-bullet-footer {
        display: flex;
        align-items: baseline;
        gap: 12px;
        margin-top: 8px;
        flex-wrap: wrap;
    }
    .ga-bullet-pct { font-weight: 700 !important; font-size: 15px !important; color: #1a1a1a !important; }
    .ga-bullet-detail { font-size: 13px !important; color: #455a64 !important; }
    .main .goal-attainment-wrap span.ga-rep-name { color: #6a1b9a !important; }
    .main .goal-attainment-wrap .ga-avatar { color: #4a148c !important; }
    .main .goal-attainment-wrap .ga-subline { color: #455a64 !important; }
    .main .goal-attainment-wrap .ga-team-sub { color: #546e7a !important; }
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > section,
    [data-testid="stAppViewContainer"] > section > div {
        background: linear-gradient(165deg, #eef7ff 0%, #faf5ff 38%, #fff8f0 72%, #f0fff8 100%) !important;
        background-color: #f3f8ff !important;
    }
    section[data-testid="stSidebar"] > div, [data-testid="stSidebar"] {
        background: linear-gradient(195deg, #ede7f6 0%, #e3f2fd 48%, #e0f7fa 100%) !important;
        border-right: 4px solid #7e57c2 !important;
        box-shadow: 2px 0 16px rgba(126, 87, 194, 0.12) !important;
    }
    section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stCaption {
        color: #000000 !important;
    }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2 {
        color: #000000 !important;
    }
    .main .block-container, .main .stMarkdown, .main p, .main label,
    .main .stCaption {
        color: #000000 !important;
    }
    .main h1 { color: #0d47a1 !important; }
    .main h2 { color: #6a1b9a !important; }
    .main h3 { color: #00695c !important; }
    /* Do not force color on .main div/span globally — breaks Streamlit st.dataframe (Glide Data Grid / canvas). */
    .main .stMarkdown div, .main .stMarkdown span, .main .stMarkdown p,
    .main .element-container p, .main .element-container span { color: #000000 !important; }
    /* st.dataframe: only background; let the grid control text (forced * { color } breaks cell rendering). */
    [data-testid="stDataFrame"] { background: #ffffff !important; }
    table td, table th, table span, table div { color: #000000 !important; }
    /* Metrics (numeric values) */
    [data-testid="stMetric"] label, [data-testid="stMetric"] div, [data-testid="stMetric"] span,
    [data-testid="stMetric"] * { color: #000000 !important; }
    .stMetric label, .stMetric div, .stMetric span { color: #000000 !important; }
    /* Caption / legend text (e.g. Slab %, PAID, N/A) */
    .main [data-testid="stCaption"] { color: #000000 !important; }
    .main [data-testid="stCaption"] * { color: #000000 !important; }
    .stExpander { background: linear-gradient(180deg, #ffffff 0%, #f8f6ff 100%) !important; border-color: #b39ddb !important; }
    .stExpander label, .stExpander p, .stExpander div, .stExpander span { color: #000000 !important; }
    div[data-testid="stExpander"] {
        border: 1px solid #b39ddb !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 12px rgba(179, 157, 219, 0.2) !important;
    }
    div[data-testid="stExpander"] * { color: #000000 !important; }
    .stButton > button {
        background: linear-gradient(135deg, #5c6bc0 0%, #3949ab 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
        box-shadow: 0 3px 12px rgba(57, 73, 171, 0.35) !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #7986cb 0%, #5c6bc0 100%) !important;
        box-shadow: 0 4px 16px rgba(57, 73, 171, 0.45) !important;
    }
    .stSelectbox label, .stRadio label, .stMultiSelect label { color: #000000 !important; }
    .stAlert { background: linear-gradient(90deg, #e8f5e9 0%, #e3f2fd 100%) !important; border: 1px solid #81c784 !important; border-left: 5px solid #43a047 !important; color: #000000 !important; }
    .stAlert * { color: #000000 !important; }
    .main input, .main .stTextInput input { background: #ffffff !important; color: #000000 !important; border: 1px solid #b39ddb !important; border-radius: 8px !important; }
    .main input:focus, .main .stTextInput input:focus {
        border-color: #7e57c2 !important;
        box-shadow: 0 0 0 3px rgba(126, 87, 194, 0.22) !important;
        outline: none !important;
    }
    .main .stTextInput label { color: #000000 !important; }
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] span, .stTabs label { color: #000000 !important; }
    .stTabs [data-baseweb="tab"] { border-radius: 10px !important; }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #5c6bc0, #7e57c2) !important;
        color: #ffffff !important;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] p, .stTabs [data-baseweb="tab"][aria-selected="true"] span {
        color: #ffffff !important;
    }
    .stTabs [role="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #5c6bc0, #7e57c2) !important;
        color: #ffffff !important;
    }
    /* Aptos for SMB goal attainment table */
    .main .goal-attainment-wrap, .main .goal-attainment-wrap * {
        font-family: "Aptos", "Segoe UI", "Segoe UI Variable", system-ui, sans-serif !important;
    }
    /* Manager Incentive: vivid gradient fill for team goal achievement progress bars */
    [data-testid="stDataFrame"] [role="progressbar"] > div {
        background: linear-gradient(90deg, #00bcd4, #7e57c2) !important;
    }
    /* --- App shell: layout / view only (same data & behavior) --- */
    .main .block-container {
        padding-left: clamp(1rem, 2.5vw, 2.25rem) !important;
        padding-right: clamp(1rem, 2.5vw, 2.25rem) !important;
        max-width: 1680px !important;
        margin-left: auto !important;
        margin-right: auto !important;
    }
    section[data-testid="stSidebar"] > div:first-child {
        padding-top: 0.75rem !important;
        padding-left: 0.65rem !important;
        padding-right: 0.65rem !important;
    }
    .admin-metric-nav-col {
        border: 2px solid #b39ddb !important;
        padding: 0.65rem 0.85rem 1rem 0.55rem !important;
        min-height: 280px;
        background: linear-gradient(160deg, #f3e5f5 0%, #e8eaf6 40%, #e1f5fe 100%) !important;
        border-radius: 14px !important;
        box-shadow: 0 4px 18px rgba(126, 87, 194, 0.15) !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        background: linear-gradient(90deg, #e8eaf6 0%, #e1f5fe 50%, #f3e5f5 100%) !important;
        border-radius: 14px !important;
        padding: 0.5rem 0.65rem !important;
        gap: 0.45rem !important;
        border: 2px solid #c5cae9 !important;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.8) !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border: 2px solid #90caf9 !important;
        border-radius: 14px !important;
        background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(227, 242, 253, 0.35) 100%) !important;
        box-shadow: 0 4px 22px rgba(33, 150, 243, 0.12) !important;
    }
    [data-testid="stMetricContainer"] {
        background: rgba(255, 255, 255, 0.85) !important;
        border-radius: 12px !important;
        border-left: 4px solid #42a5f5 !important;
        box-shadow: 0 2px 14px rgba(66, 165, 245, 0.18) !important;
        padding: 0.35rem 0.5rem !important;
    }
    </style>
    """


def _admin_sidebar_rail_css() -> str:
    """Vertical icon rail: blue gradient sidebar, larger icons/labels (admin only)."""
    return """
    <style>
    section[data-testid="stSidebar"] > div,
    section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        background: transparent !important;
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(188deg, #082a4a 0%, #0d47a1 35%, #1565c0 62%, #0c4a7d 100%) !important;
        border-right: 1px solid rgba(144, 202, 249, 0.45) !important;
        box-shadow: 6px 0 32px rgba(13, 71, 161, 0.28) !important;
    }
    /* Nav tiles: smaller than brand title; compact vertical series */
    section[data-testid="stSidebar"] [data-testid="element-container"]:has(.stButton) {
        margin-bottom: 2px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] .stButton > button {
        width: 100% !important;
        min-height: 72px !important;
        border-radius: 10px !important;
        border: none !important;
        box-shadow: none !important;
        font-family: "Segoe UI", "Aptos", system-ui, sans-serif !important;
        font-size: 1.18rem !important;
        font-weight: 600 !important;
        line-height: 1.12 !important;
        padding: 10px 6px 8px !important;
        white-space: pre-line !important;
        color: #ffffff !important;
        background: transparent !important;
        transition: background 0.15s ease !important;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] .stButton > button:hover {
        background: rgba(255,255,255,0.12) !important;
        color: #ffffff !important;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] .stButton > button[data-testid="baseButton-primary"] {
        background: rgba(255,255,255,0.18) !important;
        color: #ffffff !important;
        border-left: 4px solid #b3e5fc !important;
        border-radius: 10px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] .stButton > button[data-testid="baseButton-primary"]:hover {
        background: rgba(255,255,255,0.26) !important;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] .stButton > button p {
        font-size: 1.18rem !important;
        line-height: 1.12 !important;
        color: #ffffff !important;
        margin: 0 !important;
    }
    /* Brand block: clearly larger than nav series below */
    section[data-testid="stSidebar"] .cfz-sidebar-brand-wrap {
        text-align: center !important;
        padding: 0.65rem 0.4rem 0.35rem !important;
    }
    section[data-testid="stSidebar"] .cfz-sidebar-brand-title {
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: clamp(1.55rem, 2.6vw, 1.95rem) !important;
        letter-spacing: 0.04em !important;
        font-family: "Segoe UI", "Aptos", system-ui, sans-serif !important;
        line-height: 1.15 !important;
    }
    section[data-testid="stSidebar"] .cfz-sidebar-brand-sub {
        color: rgba(255, 255, 255, 0.95) !important;
        font-weight: 700 !important;
        font-size: clamp(0.95rem, 1.9vw, 1.22rem) !important;
        margin-top: 10px !important;
        font-family: "Segoe UI", "Aptos", system-ui, sans-serif !important;
        line-height: 1.3 !important;
        letter-spacing: 0.06em !important;
    }
    </style>
    """


# (session_value, short_label, icon_line) — matches ``render_admin_dashboard`` nav branches.
_ADMIN_SIDEBAR_RAIL_ITEMS: tuple[tuple[str, str, str], ...] = (
    ("Rules & Policy", "Policy", "📜"),
    ("SMB", "SMB", "📊"),
    ("AM", "AM", "📊"),
    ("ENT", "ENT", "📊"),
    ("Outbound", "Outbound", "📣"),
    ("Upload & Deals", "Upload", "📤"),
    ("User Management", "Users", "👥"),
    ("Export", "Export", "📥"),
    ("Assistant", "Assistant", "💬"),
)

_MEMBER_SIDEBAR_RAIL_ITEMS: tuple[tuple[str, str, str], ...] = (
    ("Rules & Policy", "Policy", "📜"),
    ("My Sales Target", "Target", "📊"),
    ("My report", "Report", "📋"),
)


def _admin_sidebar_nav_button_key(session_value: str) -> str:
    return "cfznav_" + session_value.replace("&", "and").replace(" ", "_")


def render_admin_sidebar_rail() -> None:
    """Vertical icon-style nav (snapshot-style rail) using existing admin sections."""
    st.markdown(_admin_sidebar_rail_css(), unsafe_allow_html=True)
    if "admin_sidebar_nav" not in st.session_state:
        st.session_state.admin_sidebar_nav = "Rules & Policy"
    current = st.session_state.admin_sidebar_nav
    st.sidebar.markdown(
        """
        <div class="cfz-sidebar-brand-wrap">
        <div class="cfz-sidebar-brand-title">CloudFuze</div>
        <div class="cfz-sidebar-brand-sub">Sales compensation</div>
        </div>
        <hr style="border:none;border-top:1px solid rgba(255,255,255,0.14);margin:14px 8px 16px;"/>
        """,
        unsafe_allow_html=True,
    )
    for session_value, short_label, icon in _ADMIN_SIDEBAR_RAIL_ITEMS:
        label = f"{icon}\n{short_label}"
        is_active = current == session_value
        if st.sidebar.button(
            label,
            key=_admin_sidebar_nav_button_key(session_value),
            use_container_width=True,
            type=("primary" if is_active else "secondary"),
        ):
            st.session_state.admin_sidebar_nav = session_value
            st.rerun()


def render_member_sidebar_rail() -> None:
    """Left sidebar for sales users: report vs rules & policy."""
    st.markdown(_admin_sidebar_rail_css(), unsafe_allow_html=True)
    if "member_sidebar_nav" not in st.session_state:
        st.session_state.member_sidebar_nav = "Rules & Policy"
    current = st.session_state.member_sidebar_nav
    st.sidebar.markdown(
        """
        <div class="cfz-sidebar-brand-wrap">
        <div class="cfz-sidebar-brand-title">CloudFuze</div>
        <div class="cfz-sidebar-brand-sub">Sales compensation</div>
        </div>
        <hr style="border:none;border-top:1px solid rgba(255,255,255,0.14);margin:14px 8px 16px;"/>
        """,
        unsafe_allow_html=True,
    )
    for session_value, short_label, icon in _MEMBER_SIDEBAR_RAIL_ITEMS:
        label = f"{icon}\n{short_label}"
        is_active = current == session_value
        key = "cfzmem_" + session_value.replace("&", "and").replace(" ", "_")
        if st.sidebar.button(
            label,
            key=key,
            use_container_width=True,
            type=("primary" if is_active else "secondary"),
        ):
            st.session_state.member_sidebar_nav = session_value
            st.rerun()


def main():
    # Load `.env` from the project folder (same folder as `app.py`), not only the process cwd — fixes HubSpot token when Streamlit starts from another directory.
    _root = os.path.dirname(os.path.abspath(__file__))
    load_dotenv(os.path.join(_root, ".env"))
    st.set_page_config(page_title="CloudFuze Migrate – Incentive Calculator", layout="wide")
    init_session()

    if st.session_state.user is None:
        render_login()
        return

    user = st.session_state.user
    st.markdown(_dashboard_theme_css(), unsafe_allow_html=True)

    # Welcome voice removed per request — no audio plays after sign-in.

    if user.get("role") == "ADMIN":
        render_admin_sidebar_rail()
    else:
        render_member_sidebar_rail()

    if user.get("role") == "ADMIN":
        render_admin_dashboard()
    else:
        render_member_dashboard()


if __name__ == "__main__":
    main()
