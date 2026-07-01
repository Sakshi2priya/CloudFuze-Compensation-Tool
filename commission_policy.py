"""
Q1 2026 Sales Commission & Incentive Policy — parameters for the Compensation Tool.

**Source of truth:** ``policy/commission_policy.json`` (transcribe values from your official PDF).
If that file is missing, built-in defaults below are used. Override path: env ``COMPENSATION_POLICY_JSON``.

Notes
-----
- SMB and AM **revenue commission tier tables** may appear as graphics in the PDF; transcribe
  bands and % into ``rep_slab_rows`` in the JSON file.
- ENT **quarterly individual** target is stored as the Enterprise **team_goal** for team math.
  **Enterprise named reps** (e.g. Anthony): ``enterprise_quota_achievement`` in JSON — achievement vs **hubspot_quota_usd** and tiered commission %.
- **50% individual quota** eligibility: enforce via deal **Incentive eligibility** on upload
  or **eligible_for_compensation** on the user.
- **Microsoft / Google license resale** exclusion: flag deals in data or extend the product.
- **SMB Group A / Group B**: PDF defines separate rep commission tiers. Configure in
  ``smb_subgroups`` in the JSON; assign each SMB user to **Group A** or **Group B** in
  User Management. SMB **team** goal = Group A target + Group B target (must match the SMB
  total in ``team_quarterly_targets_usd``).
- **Account Management**: ``account_management_quota_achievement`` — named individual quotas (e.g. Joy/Vivin/Arundhati),
  **Team** vs **Joy** commission columns, 50% minimum; team pool uses Team Commission tiers vs quarterly team target.
- **Outbound meetings**: NAM and Western Europe only; see ``OUTBOUND_*`` constants and Admin → Outbound.
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SMB_TEAM_NAME = "SMB"
ACCOUNT_MANAGEMENT_TEAM_NAME = "Account Management"
ENTERPRISE_TEAM_NAME = "Enterprise"

# --- Built-in defaults (used if policy/commission_policy.json is absent) ---

_DEFAULT_POLICY_NAME = "Q1 2026 Sales Commission & Incentive Policy"

_DEFAULT_TEAM_QUARTERLY_TARGETS_USD: dict[str, float] = {
    "SMB": 850_000.0,
    "Account Management": 250_000.0,
    "Enterprise": 1_000_000.0,
}

_DEFAULT_TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT: List[Tuple[float, float]] = [
    (100.0, 3.0),
    (75.0, 2.0),
    (50.0, 1.0),
]

_DEFAULT_REP_SLAB_ROWS_FOR_DB: List[Tuple[float, float | None, float]] = [
    (0, 29_999.99, 0.0),
    (30_000.0, 74_999.99, 1.0),
    (75_000.0, 99_999.99, 2.0),
    (100_000.0, None, 4.0),
]

_DEFAULT_MIN_INDIVIDUAL_QUOTA_ACHIEVEMENT_PCT = 50.0


def _policy_json_path() -> Path:
    env = (os.environ.get("COMPENSATION_POLICY_JSON") or "").strip()
    if env:
        return Path(env)
    return Path(__file__).resolve().parent / "policy" / "commission_policy.json"


def _smb_quota_snapshot_path() -> Path:
    return Path(__file__).resolve().parent / "policy" / "smb_quota_snapshot.json"


def _load_smb_quota_snapshot_raw() -> Optional[dict[str, Any]]:
    p = _smb_quota_snapshot_path()
    if not p.is_file():
        return None
    with p.open(encoding="utf-8") as f:
        data = json.load(f)
    return data if isinstance(data, dict) else None


def _load_policy_json() -> Optional[dict[str, Any]]:
    path = _policy_json_path()
    if not path.is_file():
        return None
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def _parse_rep_slab_rows(raw: list) -> List[Tuple[float, float | None, float]]:
    out: List[Tuple[float, float | None, float]] = []
    for row in raw:
        if not isinstance(row, (list, tuple)) or len(row) != 3:
            continue
        lo, hi, pct = row[0], row[1], row[2]
        max_rev: float | None
        if hi is None:
            max_rev = None
        else:
            max_rev = float(hi)
        out.append((float(lo), max_rev, float(pct)))
    return out


def _parse_team_thresholds(raw: list) -> List[Tuple[float, float]]:
    out: List[Tuple[float, float]] = []
    for row in raw:
        if isinstance(row, (list, tuple)) and len(row) == 2:
            out.append((float(row[0]), float(row[1])))
    return out


def _default_slab_thresholds_from_rows(
    rows: List[Tuple[float, float | None, float]],
) -> List[Tuple[float, float]]:
    """Calculator fallback: highest min_revenue first, same semantics as before."""
    sorted_rows = sorted(rows, key=lambda r: -r[0])
    return [(r[0], r[2]) for r in sorted_rows]


_j = _load_policy_json()

POLICY_NAME = str((_j or {}).get("policy_name") or _DEFAULT_POLICY_NAME)

TEAM_QUARTERLY_TARGETS_USD: dict[str, float] = dict(_DEFAULT_TEAM_QUARTERLY_TARGETS_USD)
if _j and isinstance(_j.get("team_quarterly_targets_usd"), dict):
    TEAM_QUARTERLY_TARGETS_USD.update(
        {k: float(v) for k, v in _j["team_quarterly_targets_usd"].items()}
    )

if _j and isinstance(_j.get("rep_slab_rows"), list) and _j["rep_slab_rows"]:
    REP_SLAB_ROWS_FOR_DB = _parse_rep_slab_rows(_j["rep_slab_rows"])
else:
    REP_SLAB_ROWS_FOR_DB = list(_DEFAULT_REP_SLAB_ROWS_FOR_DB)

if _j and isinstance(_j.get("team_achievement_commission_thresholds_pct"), list):
    _th = _parse_team_thresholds(_j["team_achievement_commission_thresholds_pct"])
    TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT = _th if _th else list(
        _DEFAULT_TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT
    )
else:
    TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT = list(_DEFAULT_TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT)

DEFAULT_SLAB_THRESHOLDS: List[Tuple[float, float]] = _default_slab_thresholds_from_rows(
    REP_SLAB_ROWS_FOR_DB
)

MIN_INDIVIDUAL_QUOTA_ACHIEVEMENT_PCT = float(
    (_j or {}).get("min_individual_quota_achievement_pct")
    or _DEFAULT_MIN_INDIVIDUAL_QUOTA_ACHIEVEMENT_PCT
)

# SMB Q1 2026: achievement % vs individual quota (Group A / B / Chitradip); replaces revenue slabs for SMB when enabled
_smb_qa = (_j or {}).get("smb_quota_achievement") if isinstance((_j or {}).get("smb_quota_achievement"), dict) else {}
SMB_QUOTA_ACHIEVEMENT_ENABLED: bool = bool(_smb_qa.get("enabled"))
SMB_ACHIEVEMENT_TIERS: List[Dict[str, Any]] = list(_smb_qa.get("achievement_tiers") or [])
SMB_INDIVIDUAL_QUOTAS_USD: Dict[str, float] = {}
if isinstance(_smb_qa.get("individual_quotas_usd"), dict):
    SMB_INDIVIDUAL_QUOTAS_USD = {
        str(k).strip().upper(): float(v) for k, v in _smb_qa["individual_quotas_usd"].items()
    }
SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT: float = float(
    _smb_qa.get("min_achievement_pct_for_commission") or MIN_INDIVIDUAL_QUOTA_ACHIEVEMENT_PCT
)


def smb_rep_incentive_ineligible_display() -> str:
    """Rep incentives / UI text when payout does not apply (min achievement, flags, or SMB group mismatch)."""
    return f"Below {SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:.0f}% quota achievement"


# Chitradip (SMB manager): team revenue / team goal → commission % (separate from rep tiers)
SMB_MANAGER_CHITRADIP_TEAM_TIERS: List[Dict[str, Any]] = list(
    _smb_qa.get("manager_chitradip_team_tiers") or []
)

# Account Management: individual quota achievement; Team vs Joy commission columns (policy JSON)
_am_qa = (
    (_j or {}).get("account_management_quota_achievement")
    if isinstance((_j or {}).get("account_management_quota_achievement"), dict)
    else {}
)
AM_QUOTA_ACHIEVEMENT_ENABLED: bool = bool(_am_qa.get("enabled"))
AM_ACHIEVEMENT_TIERS: List[Dict[str, Any]] = list(_am_qa.get("achievement_tiers") or [])
AM_INDIVIDUAL_QUOTA_ROWS: List[Dict[str, Any]] = [
    x for x in (_am_qa.get("individual_quotas") or []) if isinstance(x, dict)
]
AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT: float = float(
    _am_qa.get("min_achievement_pct_for_commission") or MIN_INDIVIDUAL_QUOTA_ACHIEVEMENT_PCT
)

# Enterprise: named reps with individual quota achievement tiers (policy JSON)
_ent_qa = (
    (_j or {}).get("enterprise_quota_achievement")
    if isinstance((_j or {}).get("enterprise_quota_achievement"), dict)
    else {}
)
ENTERPRISE_QUOTA_ACHIEVEMENT_ENABLED: bool = bool(_ent_qa.get("enabled"))
ENT_ENTERPRISE_REP_TIERS: List[Dict[str, Any]] = [
    x for x in (_ent_qa.get("reps") or []) if isinstance(x, dict)
]

# Enterprise upload UI: optional Anthony "Performance vs goal" snapshot (policy / deck numbers)
_ent_pvg = (
    (_j or {}).get("enterprise_anthony_performance_vs_goal")
    if isinstance((_j or {}).get("enterprise_anthony_performance_vs_goal"), dict)
    else {}
)
ENTERPRISE_ANTHONY_PVG_ENABLED: bool = bool(_ent_pvg.get("enabled"))
try:
    ENTERPRISE_ANTHONY_PVG_TARGET: float = float(_ent_pvg.get("target_usd") or 0.0)
except (TypeError, ValueError):
    ENTERPRISE_ANTHONY_PVG_TARGET = 0.0
try:
    ENTERPRISE_ANTHONY_PVG_ACHIEVED: float = float(_ent_pvg.get("achieved_usd") or 0.0)
except (TypeError, ValueError):
    ENTERPRISE_ANTHONY_PVG_ACHIEVED = 0.0
ENTERPRISE_ANTHONY_PVG_SUBTITLE: str = str(_ent_pvg.get("subtitle") or "Individual target · Q1 2026")

# Optional per-rep overrides (year/quarter); used only when snapshot mode is off or no snapshot match
INDIVIDUAL_QUOTA_OVERRIDES: List[Dict[str, Any]] = []
if _j and isinstance(_j.get("individual_quota_overrides"), list):
    INDIVIDUAL_QUOTA_OVERRIDES = [x for x in _j["individual_quota_overrides"] if isinstance(x, dict)]

# SMB: authoritative per-rep quotas from policy/smb_quota_snapshot.json (Group A/B); HubSpot forecast ignored when enabled
SMB_USE_QUOTA_SNAPSHOT: bool = bool((_j or {}).get("smb_use_quota_snapshot", False))
SMB_QUOTA_SNAPSHOT_RAW: Optional[dict[str, Any]] = _load_smb_quota_snapshot_raw() if SMB_USE_QUOTA_SNAPSHOT else None

# SMB Group A / B — legacy rep slab sets (DB keys: SMB_A, SMB_B); skipped when SMB_QUOTA_ACHIEVEMENT_ENABLED
SMB_SUBGROUP_SLAB_ROWS: Dict[str, List[Tuple[float, float | None, float]]] = {}
SMB_SUBGROUP_LABELS: Dict[str, str] = {}
SMB_SUBGROUP_TEAM_TARGETS_USD: Dict[str, float] = {}
if _j and isinstance(_j.get("smb_subgroups"), dict):
    _sg = _j["smb_subgroups"]
    for letter, key in (("A", "SMB_A"), ("B", "SMB_B")):
        block = _sg.get(letter)
        if isinstance(block, dict) and block.get("rep_slab_rows"):
            SMB_SUBGROUP_SLAB_ROWS[key] = _parse_rep_slab_rows(block["rep_slab_rows"])
            lbl = (block.get("label") or f"SMB Group {letter}").strip()
            SMB_SUBGROUP_LABELS[key] = lbl
            if block.get("team_quarterly_target_usd") is not None:
                SMB_SUBGROUP_TEAM_TARGETS_USD[letter] = float(block["team_quarterly_target_usd"])


def smb_team_goal_from_subgroups() -> Optional[float]:
    """If PDF defines A and B quarterly targets, return A+B for the single SMB team_goal row."""
    if len(SMB_SUBGROUP_TEAM_TARGETS_USD) < 2:
        return None
    return float(sum(SMB_SUBGROUP_TEAM_TARGETS_USD.values()))


def all_slab_sets_for_db() -> Dict[str, List[Tuple[float, float | None, float]]]:
    """All named slab sets to seed in ``incentive_slabs``. SMB quota mode seeds only DEFAULT."""
    out: Dict[str, List[Tuple[float, float | None, float]]] = {
        "DEFAULT": list(REP_SLAB_ROWS_FOR_DB),
    }
    if SMB_QUOTA_ACHIEVEMENT_ENABLED:
        return out
    out.update(SMB_SUBGROUP_SLAB_ROWS)
    return out


def resolve_slab_set(team_name: str | None, compensation_group: str | None) -> str:
    """
    Map deal owner team + user.compensation_group to a slab set name.

    When SMB uses quota-vs-achievement (policy JSON), SMB reps still resolve to DEFAULT here;
    the engine applies ``smb_commission_pct_from_quota_achievement`` instead of DB slabs.
    """
    if (team_name or "").strip() != SMB_TEAM_NAME:
        return "DEFAULT"
    if SMB_QUOTA_ACHIEVEMENT_ENABLED:
        return "DEFAULT"
    g = (compensation_group or "").strip().upper()
    if g in ("SMB_A", "A"):
        return "SMB_A" if "SMB_A" in SMB_SUBGROUP_SLAB_ROWS else "DEFAULT"
    if g in ("SMB_B", "B"):
        return "SMB_B" if "SMB_B" in SMB_SUBGROUP_SLAB_ROWS else "DEFAULT"
    if g in ("SMB_CHITRADIP", "CHITRADIP"):
        return "SMB_CHITRADIP" if "SMB_CHITRADIP" in SMB_SUBGROUP_SLAB_ROWS else "DEFAULT"
    return "DEFAULT"


def default_slab_thresholds_for_slab_set(slab_set: str) -> List[Tuple[float, float]]:
    """Calculator fallback thresholds for a slab set (when DB unavailable)."""
    rows = all_slab_sets_for_db().get(slab_set) or REP_SLAB_ROWS_FOR_DB
    return _default_slab_thresholds_from_rows(rows)


# --- Outbound meeting incentives (policy deck) ---
OUTBOUND_POLICY_LABEL = "Outbound Meeting Incentives Structure-2026"
OUTBOUND_POLICY_ELIGIBILITY = (
    "Applies strictly to meetings booked in NAM (North America) and Western Europe regions."
)
OUTBOUND_ELIGIBLE_REGIONS: List[Tuple[str, str]] = [
    ("NAM", "NAM (North America)"),
    ("WE", "Western Europe"),
]

_ob_payout = (
    (_j or {}).get("outbound_meeting_payout_tiers")
    if isinstance((_j or {}).get("outbound_meeting_payout_tiers"), dict)
    else {}
)
OUTBOUND_MEETING_PAYOUT_TITLE: str = str(_ob_payout.get("title") or "Outbound Meetings Payout")
_out_rows_raw = _ob_payout.get("rows") if isinstance(_ob_payout.get("rows"), list) else []
OUTBOUND_MEETING_PAYOUT_ROWS: List[Dict[str, Any]] = []
for _r in _out_rows_raw:
    if isinstance(_r, dict) and _r.get("meetings") is not None and _r.get("payout") is not None:
        OUTBOUND_MEETING_PAYOUT_ROWS.append({"meetings": str(_r["meetings"]), "payout": str(_r["payout"])})
if not OUTBOUND_MEETING_PAYOUT_ROWS:
    OUTBOUND_MEETING_PAYOUT_ROWS = [
        {"meetings": "0 – 5 meetings", "payout": "No payout"},
        {"meetings": "6 – 10 meetings", "payout": "₹3,000 per meeting"},
        {"meetings": "11 – 15 meetings", "payout": "₹5,000 per meeting"},
    ]
OUTBOUND_MEETING_PAYOUT_NOTE: str = str(
    _ob_payout.get("note")
    or (
        "e.g. 8 meetings → 8 × ₹3,000 when in the 6–10 band"
    )
)


def outbound_region_code_is_eligible(region_code: str) -> bool:
    """True if region code is one of the PPT-eligible outbound regions."""
    c = (region_code or "").strip().upper()
    return c in ("NAM", "WE")


def display_label_for_outbound_region(region_code: str) -> Optional[str]:
    for code, label in OUTBOUND_ELIGIBLE_REGIONS:
        if code.upper() == (region_code or "").strip().upper():
            return label
    return None


def team_commission_pct_from_achievement(achievement_pct: float) -> float:
    """
    Team incentive commission % from (team revenue / team goal) × 100.

    First matching tier: achievement must be *strictly greater* than the threshold
    (matches: >50% → 1%, >75% → 2%, >100% → 3%).
    """
    for threshold, pct in TEAM_ACHIEVEMENT_COMMISSION_THRESHOLDS_PCT:
        if achievement_pct > threshold:
            return pct
    return 0.0


def ent_team_commission_pct_from_achievement(achievement_pct: float) -> float:
    """ENT-specific manager incentive % — uses the Enterprise rep tier bands.

    Per the Enterprise policy:
      • 0% – 75.99%   → 7%
      • 76% – 100%    → 9%
      • 101% – 125%   → 11%
      • 126%+         → 13%
    Minimum achievement for ENT commission is 0% (any positive revenue qualifies).
    """
    x = float(achievement_pct or 0)
    if x >= 126:
        return 13.0
    if x >= 101:
        return 11.0
    if x >= 76:
        return 9.0
    # Any achievement below 76% (including 0) lands in the 0–75.99% band.
    return 7.0


def policy_config_path_resolved() -> str:
    """Absolute path to the JSON policy file in use (for UI / debugging)."""
    return str(_policy_json_path().resolve())


def smb_individual_quota_usd(compensation_group: str | None) -> float:
    """Individual quarterly quota (USD) for SMB_A / SMB_B from policy JSON (not Chitradip — manager is team-based)."""
    key = (compensation_group or "").strip().upper()
    return float(SMB_INDIVIDUAL_QUOTAS_USD.get(key, 0.0))


def period_string_to_year_quarter(calculation_period: str | None) -> Optional[Tuple[int, int]]:
    """Map ``calculation_period`` like ``Mar 2026`` or ``January 2026`` to ``(year, quarter 1–4)``."""
    if not calculation_period or not str(calculation_period).strip():
        return None
    s = str(calculation_period).strip()
    for fmt in ("%b %Y", "%B %Y"):
        try:
            dt = datetime.strptime(s, fmt)
            m = dt.month
            y = dt.year
            q = (m - 1) // 3 + 1
            return (y, q)
        except ValueError:
            continue
    return None


def individual_quota_override_usd(
    full_name: str | None,
    email: str | None,
    year: int,
    quarter: int,
) -> Optional[float]:
    """
    Match ``policy/commission_policy.json`` ``individual_quota_overrides`` by year/quarter and name tokens
    (substring of full name or email local-part).
    """
    if not INDIVIDUAL_QUOTA_OVERRIDES:
        return None
    fn = (full_name or "").lower()
    local = (email or "").split("@")[0].lower() if email else ""
    for row in INDIVIDUAL_QUOTA_OVERRIDES:
        try:
            y = int(row.get("year"))
            q = int(row.get("quarter"))
        except (TypeError, ValueError):
            continue
        if y != year or q != quarter:
            continue
        tokens = row.get("name_tokens") or []
        if not isinstance(tokens, list):
            continue
        for t in tokens:
            t = str(t).lower().strip()
            if not t:
                continue
            if t in fn or (local and t in local):
                try:
                    v = float(row.get("quota_usd", 0))
                except (TypeError, ValueError):
                    continue
                if v > 0:
                    return v
    return None


def _name_matches_any_token(full_name: str | None, email: str | None, tokens: Any) -> bool:
    if not isinstance(tokens, list):
        return False
    fn = (full_name or "").lower()
    local = (email or "").split("@")[0].lower() if email else ""
    for t in tokens:
        t = str(t).lower().strip()
        if not t:
            continue
        if t in fn or (local and t in local):
            return True
    return False


def smb_resolve_group_for_rep(
    compensation_group: str | None,
    full_name: str | None,
    email: str | None,
) -> Optional[str]:
    """
    Return ``SMB_A`` or ``SMB_B`` for rep commission tiers.

    Prefer **User Management** ``compensation_group`` when set; otherwise infer from the same
    quota snapshot name/email rules as ``smb_individual_quota_usd_for_rep`` so incentive % is not
    zero when quota was already resolved from the snapshot but the DB group was not assigned.
    """
    g = (compensation_group or "").strip().upper()
    if g in ("SMB_A", "SMB_B"):
        return g
    sq = smb_snapshot_subgroup_and_quota(full_name, email)
    if sq is not None:
        return sq[1]
    return None


def smb_snapshot_subgroup_and_quota(
    full_name: str | None,
    email: str | None,
) -> Optional[Tuple[float, str]]:
    """
    If ``policy/smb_quota_snapshot.json`` is loaded, return ``(quota_usd, 'SMB_A'|'SMB_B')`` when the rep
    matches Group A or B name tokens; otherwise ``None``.
    """
    if not SMB_USE_QUOTA_SNAPSHOT or not SMB_QUOTA_SNAPSHOT_RAW:
        return None
    snap = SMB_QUOTA_SNAPSHOT_RAW
    ga = snap.get("group_a") if isinstance(snap.get("group_a"), dict) else {}
    gb = snap.get("group_b") if isinstance(snap.get("group_b"), dict) else {}
    if _name_matches_any_token(full_name, email, ga.get("name_tokens")):
        try:
            q = float(ga.get("quota_usd_each", 0))
        except (TypeError, ValueError):
            q = 0.0
        return (q, "SMB_A") if q > 0 else None
    if _name_matches_any_token(full_name, email, gb.get("name_tokens")):
        try:
            q = float(gb.get("quota_usd_each", 0))
        except (TypeError, ValueError):
            q = 0.0
        return (q, "SMB_B") if q > 0 else None
    return None


def smb_rep_chart_group(
    full_name: str | None,
    email: str | None,
    compensation_group: str | None,
) -> str:
    """UI label for charts: Group A / Group B / Other — snapshot names first, then ``compensation_group``."""
    sq = smb_snapshot_subgroup_and_quota(full_name, email)
    if sq is not None:
        return "Group A" if sq[1] == "SMB_A" else "Group B"
    cg = (compensation_group or "").strip().upper()
    if cg == "SMB_A":
        return "Group A"
    if cg == "SMB_B":
        return "Group B"
    return "Other"


def smb_individual_quota_usd_for_rep(
    compensation_group: str | None,
    hubspot_quota_usd: Any,
    *,
    full_name: str | None = None,
    email: str | None = None,
    calculation_period: str | None = None,
) -> float:
    """
    Individual quota for SMB reps.

    When ``smb_use_quota_snapshot`` is true and ``smb_quota_snapshot.json`` exists: snapshot Group A/B
    quotas by name (HubSpot forecast is not used). Unmatched SMB reps use Group A/B defaults from JSON.

    When snapshot mode is off: optional ``individual_quota_overrides`` for the period, then HubSpot, then
    Group A/B defaults.
    """
    if SMB_USE_QUOTA_SNAPSHOT and SMB_QUOTA_SNAPSHOT_RAW:
        sq = smb_snapshot_subgroup_and_quota(full_name, email)
        if sq is not None:
            return float(sq[0])
        return smb_individual_quota_usd(compensation_group)

    yq = period_string_to_year_quarter(calculation_period)
    if yq:
        ov = individual_quota_override_usd(full_name, email, yq[0], yq[1])
        if ov is not None and ov > 0:
            return float(ov)
    try:
        if hubspot_quota_usd is not None and str(hubspot_quota_usd).strip() != "":
            v = float(hubspot_quota_usd)
            if v > 0:
                return v
    except (TypeError, ValueError):
        pass
    return smb_individual_quota_usd(compensation_group)


def _match_smb_deck_tier_band(achievement_pct: float, lo: float, hi: Optional[float]) -> bool:
    """
    Map achievement % to policy PDF bands without gaps.

    - Closed bands (50–75, 76–100): ``lo <= x <= hi``.
    - **101–125** row in the deck: treat as **strictly above 100%** through **125%** inclusive,
      i.e. ``100 < x <= 125`` (so e.g. 100.68% counts as the 101–125 tier, not 76–100).
    - **126% & above** (``hi`` is None, ``lo`` typically 126): ``x > 125``.
    """
    x = float(achievement_pct)
    if hi is None:
        if lo >= 126:
            return x > 125.0
        return x >= lo
    hi_f = float(hi)
    if lo >= 101.0 and hi_f <= 125.0:
        return 100.0 < x <= hi_f
    return lo <= x <= hi_f


def smb_manager_chitradip_team_commission_pct(team_achievement_pct: float) -> float:
    """
    Manager (Chitradip) commission % from **team** achievement: (team revenue / team goal) × 100.

    Uses ``manager_chitradip_team_tiers`` in policy JSON. Below 50% team achievement → 0%.
    """
    if not SMB_MANAGER_CHITRADIP_TEAM_TIERS:
        return 0.0
    x = float(team_achievement_pct)
    if x < SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:
        return 0.0
    for tier in SMB_MANAGER_CHITRADIP_TEAM_TIERS:
        lo = float(tier["min_pct"])
        hi = tier.get("max_pct")
        pct = float(tier["commission_pct"])
        if _match_smb_deck_tier_band(x, lo, hi):
            return pct
    return 0.0


def smb_commission_pct_from_quota_achievement(
    achievement_pct: float,
    compensation_group: str | None,
) -> float:
    """
    Commission % from (total revenue / individual quota) × 100, using SMB tier table.

    Returns 0 if achievement is below ``SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT`` (default 50%)
    or group is not SMB_A / SMB_B. (Chitradip manager comp is team-based, not this function.)
    """
    if not SMB_QUOTA_ACHIEVEMENT_ENABLED or not SMB_ACHIEVEMENT_TIERS:
        return 0.0
    x = float(achievement_pct)
    if x < SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:
        return 0.0
    g = (compensation_group or "").strip().upper()
    pct_field = {
        "SMB_A": "group_a_pct",
        "SMB_B": "group_b_pct",
    }.get(g)
    if not pct_field:
        return 0.0
    for tier in SMB_ACHIEVEMENT_TIERS:
        lo = float(tier["min_pct"])
        hi = tier.get("max_pct")
        if _match_smb_deck_tier_band(x, lo, hi):
            return float(tier[pct_field])
    return 0.0


def am_rep_incentive_ineligible_display() -> str:
    return f"Below {AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:.0f}% quota achievement"


def am_rep_is_joy(full_name: str | None, email: str | None) -> bool:
    """Joy uses the Joy Commission column; match first name or email local-part."""
    fn = (full_name or "").strip().lower()
    parts = fn.split()
    first = (parts[0].rstrip(".,;") if parts else "")
    if first == "joy":
        return True
    local = (email or "").split("@")[0].lower() if email else ""
    if local == "joy" or local.startswith("joy."):
        return True
    return False


def am_individual_quota_from_policy_names(full_name: str | None, email: str | None) -> float:
    for row in AM_INDIVIDUAL_QUOTA_ROWS:
        tokens = row.get("name_tokens") or []
        if not isinstance(tokens, list):
            continue
        if not _name_matches_any_token(full_name, email, tokens):
            continue
        try:
            q = float(row.get("quota_usd", 0))
        except (TypeError, ValueError):
            continue
        if q > 0:
            return q
    return 0.0


def am_individual_quota_usd_for_rep(
    hubspot_quota_usd: Any,
    *,
    full_name: str | None = None,
    email: str | None = None,
    calculation_period: str | None = None,
) -> float:
    """
    Account Management individual quarterly quota: policy name list, then period overrides, then HubSpot.
    """
    qn = am_individual_quota_from_policy_names(full_name, email)
    if qn > 0:
        return qn
    yq = period_string_to_year_quarter(calculation_period)
    if yq:
        ov = individual_quota_override_usd(full_name, email, yq[0], yq[1])
        if ov is not None and ov > 0:
            return float(ov)
    try:
        if hubspot_quota_usd is not None and str(hubspot_quota_usd).strip() != "":
            v = float(hubspot_quota_usd)
            if v > 0:
                return v
    except (TypeError, ValueError):
        pass
    return 0.0


def am_commission_pct_from_quota_achievement(
    achievement_pct: float,
    full_name: str | None,
    email: str | None,
) -> float:
    """
    AM rep commission % from individual achievement vs quota.

    Uses **team_pct** for Vivin, Arundhati, and other AM reps; **joy_pct** for Joy.
    """
    if not AM_QUOTA_ACHIEVEMENT_ENABLED or not AM_ACHIEVEMENT_TIERS:
        return 0.0
    x = float(achievement_pct)
    if x < AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:
        return 0.0
    pct_field = "joy_pct" if am_rep_is_joy(full_name, email) else "team_pct"
    for tier in AM_ACHIEVEMENT_TIERS:
        lo = float(tier["min_pct"])
        hi = tier.get("max_pct")
        if pct_field not in tier:
            continue
        if _match_smb_deck_tier_band(x, lo, hi):
            return float(tier[pct_field])
    return 0.0


def am_team_commission_pct_from_team_achievement(team_achievement_pct: float) -> float:
    """
    AM team incentive pool: team revenue vs quarterly team target, using **team_pct** tiers only.
    """
    if not AM_QUOTA_ACHIEVEMENT_ENABLED or not AM_ACHIEVEMENT_TIERS:
        return 0.0
    x = float(team_achievement_pct)
    if x < AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:
        return 0.0
    for tier in AM_ACHIEVEMENT_TIERS:
        lo = float(tier["min_pct"])
        hi = tier.get("max_pct")
        if _match_smb_deck_tier_band(x, lo, hi):
            return float(tier["team_pct"])
    return 0.0


def ent_rep_config_for_name(full_name: str | None, email: str | None) -> Optional[Dict[str, Any]]:
    """Return the ``enterprise_quota_achievement.reps`` block when **name_tokens** match."""
    if not ENT_ENTERPRISE_REP_TIERS:
        return None
    for rep in ENT_ENTERPRISE_REP_TIERS:
        tokens = rep.get("name_tokens") or []
        if isinstance(tokens, list) and _name_matches_any_token(full_name, email, tokens):
            return rep
    return None


def ent_individual_quota_usd_for_rep(
    hubspot_quota_usd: Any,
    *,
    full_name: str | None = None,
    email: str | None = None,
    calculation_period: str | None = None,
) -> float:
    """Enterprise named-rep mode: **hubspot_quota_usd** + optional period overrides (same as AM overrides path)."""
    yq = period_string_to_year_quarter(calculation_period)
    if yq:
        ov = individual_quota_override_usd(full_name, email, yq[0], yq[1])
        if ov is not None and ov > 0:
            return float(ov)
    try:
        if hubspot_quota_usd is not None and str(hubspot_quota_usd).strip() != "":
            v = float(hubspot_quota_usd)
            if v > 0:
                return v
    except (TypeError, ValueError):
        pass
    return 0.0


def ent_commission_pct_from_quota_achievement(
    achievement_pct: float,
    full_name: str | None,
    email: str | None,
) -> float:
    """
    Enterprise named reps: commission % from (revenue / individual quota) × 100, using **achievement_tiers** in JSON.
    """
    rep = ent_rep_config_for_name(full_name, email)
    if not rep:
        return 0.0
    tiers = [t for t in (rep.get("achievement_tiers") or []) if isinstance(t, dict)]
    if not tiers:
        return 0.0
    raw_min = rep.get("min_achievement_pct_for_commission")
    if raw_min is None:
        min_comm = 75.0
    else:
        try:
            min_comm = float(raw_min)
        except (TypeError, ValueError):
            min_comm = 75.0
    x = float(achievement_pct)
    if x < min_comm:
        return 0.0
    for tier in tiers:
        if "commission_pct" not in tier:
            continue
        lo = float(tier["min_pct"])
        hi = tier.get("max_pct")
        if hi is not None:
            try:
                hi = float(hi)
            except (TypeError, ValueError):
                hi = None
        if _match_smb_deck_tier_band(x, lo, hi):
            return float(tier["commission_pct"])
    return 0.0


def ent_min_achievement_pct_for_rep(full_name: str | None, email: str | None) -> float:
    rep = ent_rep_config_for_name(full_name, email)
    if not rep:
        return 75.0
    raw_min = rep.get("min_achievement_pct_for_commission")
    if raw_min is None:
        return 75.0
    try:
        return float(raw_min)
    except (TypeError, ValueError):
        return 75.0


def ent_rep_incentive_ineligible_display(full_name: str | None, email: str | None) -> str:
    m = ent_min_achievement_pct_for_rep(full_name, email)
    if m <= 0:
        return "Not eligible for commission"
    return f"Below {m:.0f}% quota achievement"
