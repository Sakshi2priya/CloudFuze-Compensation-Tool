"""
Incentive calculation engine for CloudFuze Migrate.

Rep incentives:
- **Non-SMB** (except AM quota mode): Slab % from total Amount (revenue bands). Incentive $ = paid amount × slab %.
- **SMB reps (Group A / B)**: Commission % from **individual quota** achievement (policy tiers). Min **50%**.
  Excludes **license resale** deals. Incentive $ = paid × commission %.
- **Account Management** (when enabled in policy): Individual quota (policy names / overrides / HubSpot) and **Team vs Joy**
  tier columns. Min **50%** of individual quota. Incentive $ = paid × commission %.
- **Enterprise** (when enabled in policy): Named reps (e.g. Anthony) use **enterprise_quota_achievement** tiers vs individual quota (**hubspot_quota_usd**). Other Enterprise reps use revenue slabs.
- **SMB Chitradip** (manager, ``SMB_CHITRADIP``): no rep row — compensation is **manager incentive** only.

Team incentives:
- **SMB + manager marked SMB_CHITRADIP**: commission % from **team** revenue vs team goal (Chitradip tier table).
- **Account Management** (quota mode): team pool uses **Team Commission** tier bands vs team quarterly target.
- **Other teams**: total_team_revenue × % from standard team achievement bands.

On Finalize: aggregates deals, applies rules, generates rep and team incentives.
"""

from collections import defaultdict
from datetime import datetime
from typing import List, Tuple

from calculator import calculate_incentive_from_paid
from commission_policy import (
    ACCOUNT_MANAGEMENT_TEAM_NAME,
    AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
    AM_QUOTA_ACHIEVEMENT_ENABLED,
    ENTERPRISE_QUOTA_ACHIEVEMENT_ENABLED,
    ENTERPRISE_TEAM_NAME,
    SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT,
    SMB_QUOTA_ACHIEVEMENT_ENABLED,
    SMB_TEAM_NAME,
    am_commission_pct_from_quota_achievement,
    am_individual_quota_usd_for_rep,
    am_rep_incentive_ineligible_display,
    am_team_commission_pct_from_team_achievement,
    ent_commission_pct_from_quota_achievement,
    ent_individual_quota_usd_for_rep,
    ent_min_achievement_pct_for_rep,
    ent_rep_config_for_name,
    ent_rep_incentive_ineligible_display,
    resolve_slab_set,
    smb_commission_pct_from_quota_achievement,
    smb_individual_quota_usd_for_rep,
    smb_manager_chitradip_team_commission_pct,
    smb_rep_incentive_ineligible_display,
    smb_resolve_group_for_rep,
    team_commission_pct_from_achievement,
)
from excel_service import effective_paid_amount_from_status, is_license_resale_excluded_deal
from database import (
    delete_rep_incentives_by_period,
    delete_team_incentives_by_period,
    get_active_slabs,
    get_all_teams,
    get_deals_by_upload,
    get_upload_by_id,
    get_user_by_id,
    get_user_ids_not_eligible_for_compensation,
    get_users_for_team,
    insert_rep_incentives,
    insert_team_incentives,
)


def _deal_row_marked_ineligible(incentive_eligibility: str | None) -> bool:
    """True when uploaded deal says the rep should not earn commission on this row."""
    u = (incentive_eligibility or "").strip().upper()
    if u in ("NOT ELIGIBLE", "N/A", "NO"):
        return True
    if "BELOW" in u and "QUOTA" in u and "ACHIEVEMENT" in u:
        return True
    return False


def _get_team_lead(team_id: int) -> int | None:
    """Get the first SALES_MANAGER in the team as team lead."""
    users = get_users_for_team(team_id)
    for u in users:
        if u.get("role") == "SALES_MANAGER":
            return u["user_id"]
    return None


def compute_and_store_incentives(upload_id: int, period: str | None = None) -> Tuple[int, int]:
    """
    Aggregate deals from a finalized upload, calculate rep and team incentives, store them.

    Args:
        upload_id: The finalized upload.
        period: Calculation period (e.g. "Feb 2025"). Defaults to current month.

    Returns:
        (count_rep_incentives, count_team_incentives)
    """
    if not period:
        period = datetime.now().strftime("%b %Y")

    upload = get_upload_by_id(upload_id)
    close_date_fallback = None
    if upload and upload.get("finalized_at"):
        close_date_fallback = upload["finalized_at"].date() if hasattr(upload["finalized_at"], "date") else None
    if close_date_fallback is None:
        close_date_fallback = datetime.now().date()

    delete_rep_incentives_by_period(period)
    delete_team_incentives_by_period(period)

    deals = get_deals_by_upload(upload_id)
    slabs_default = get_active_slabs("DEFAULT")

    paid_deals = [d for d in deals if (d.get("payment_status") or "").upper() == "PAID"]
    paid_or_partial_deals = [d for d in deals if (d.get("payment_status") or "").upper() in ("PAID", "PARTIALLY_PAID")]

    rep_total_amount_all: dict[Tuple[int, int], float] = defaultdict(float)
    team_total_amount_all: dict[int, float] = defaultdict(float)
    for d in deals:
        amt = float(d["amount"])
        tn = (d.get("team_name") or "").strip()
        if tn == SMB_TEAM_NAME and is_license_resale_excluded_deal(d):
            continue
        rep_total_amount_all[(d["deal_owner_id"], d["team_id"])] += amt
        team_total_amount_all[d["team_id"]] += amt

    rep_agg: dict[Tuple[int, int], List[dict]] = defaultdict(list)
    for d in paid_or_partial_deals:
        key = (d["deal_owner_id"], d["team_id"])
        rep_agg[key].append(d)

    team_agg: dict[int, List[dict]] = defaultdict(list)
    for d in paid_deals:
        team_agg[d["team_id"]].append(d)

    ineligible_rep_ids = set(get_user_ids_not_eligible_for_compensation())
    rep_slabs_cache: dict[str, list] = {}

    def _slabs_for(slab_set: str) -> list:
        if slab_set not in rep_slabs_cache:
            rep_slabs_cache[slab_set] = get_active_slabs(slab_set)
        return rep_slabs_cache[slab_set]

    rep_rows: list = []
    for (user_id, team_id), _total_revenue_raw in rep_total_amount_all.items():
        deal_list = rep_agg.get((user_id, team_id), [])
        all_deals_for_rep = [d for d in deals if d["deal_owner_id"] == user_id and d["team_id"] == team_id]
        sample = all_deals_for_rep[0] if all_deals_for_rep else None
        team_name = (sample.get("team_name") or "").strip() if sample else None
        owner_cg = sample.get("owner_compensation_group") if sample else None

        any_not_eligible_from_file = any(
            _deal_row_marked_ineligible(d.get("incentive_eligibility")) for d in all_deals_for_rep
        )
        not_eligible = any_not_eligible_from_file or (user_id in ineligible_rep_ids)

        eligible_for_smb = [
            d
            for d in all_deals_for_rep
            if not is_license_resale_excluded_deal(d)
        ]
        deal_list_smb = [
            d
            for d in deal_list
            if not is_license_resale_excluded_deal(d)
        ]

        if team_name == SMB_TEAM_NAME and SMB_QUOTA_ACHIEVEMENT_ENABLED:
            if (owner_cg or "").strip().upper() == "SMB_CHITRADIP":
                # Chitradip is SMB manager: team-based pay only (manager incentive row), not rep quota.
                continue

            total_revenue = sum(float(d["amount"]) for d in eligible_for_smb)
            total_paid_amount = sum(effective_paid_amount_from_status(x) for x in deal_list_smb)
            hubspot_q = sample.get("owner_hubspot_quota_usd") if sample else None
            deal_name = sample.get("deal_owner_name") if sample else None
            deal_email = sample.get("deal_owner_email") if sample else None
            resolved_smb_group = smb_resolve_group_for_rep(owner_cg, deal_name, deal_email)
            quota = smb_individual_quota_usd_for_rep(
                owner_cg,
                hubspot_q,
                full_name=deal_name,
                email=deal_email,
                calculation_period=period,
            )
            if quota <= 0:
                achievement_pct = 0.0
                incentive_pct = 0.0
            else:
                achievement_pct = (total_revenue / quota) * 100.0
                incentive_pct = smb_commission_pct_from_quota_achievement(achievement_pct, resolved_smb_group)
            res_slab_id = slabs_default[-1]["slab_id"] if slabs_default else 1

            _inel = smb_rep_incentive_ineligible_display()
            if not_eligible:
                incentive_eligibility = _inel
                incentive_amount = 0.0
            elif quota <= 0 or (resolved_smb_group or "").strip().upper() not in ("SMB_A", "SMB_B"):
                incentive_eligibility = _inel
                incentive_amount = 0.0
            elif achievement_pct < SMB_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:
                incentive_eligibility = _inel
                incentive_amount = 0.0
            else:
                incentive_eligibility = "Eligible"
                incentive_amount = round(total_paid_amount * (incentive_pct / 100.0), 2)

            incentive_pct = round(float(incentive_pct), 2)

            any_partial = any((x.get("payment_status") or "").upper() == "PARTIALLY_PAID" for x in deal_list)
            payment_status = "PARTIALLY_PAID" if any_partial else "PAID"
            deal_close_dates = [d["close_date"] for d in deal_list if d.get("close_date")]
            row_close_date = max(deal_close_dates) if deal_close_dates else close_date_fallback

            rep_rows.append(
                (
                    user_id,
                    team_id,
                    len(deal_list),
                    total_revenue,
                    total_paid_amount,
                    res_slab_id,
                    incentive_pct,
                    incentive_amount,
                    payment_status,
                    period,
                    row_close_date,
                    incentive_eligibility,
                    quota if quota > 0 else None,
                )
            )
            continue

        if team_name == ACCOUNT_MANAGEMENT_TEAM_NAME and AM_QUOTA_ACHIEVEMENT_ENABLED:
            total_revenue = sum(float(d["amount"]) for d in all_deals_for_rep)
            total_paid_amount = sum(effective_paid_amount_from_status(x) for x in deal_list)
            hubspot_q = sample.get("owner_hubspot_quota_usd") if sample else None
            deal_name = sample.get("deal_owner_name") if sample else None
            deal_email = sample.get("deal_owner_email") if sample else None
            quota = am_individual_quota_usd_for_rep(
                hubspot_q,
                full_name=deal_name,
                email=deal_email,
                calculation_period=period,
            )
            if quota <= 0:
                achievement_pct = 0.0
                incentive_pct = 0.0
            else:
                achievement_pct = (total_revenue / quota) * 100.0
                incentive_pct = am_commission_pct_from_quota_achievement(
                    achievement_pct, deal_name, deal_email
                )
            res_slab_id = slabs_default[-1]["slab_id"] if slabs_default else 1
            _inel_am = am_rep_incentive_ineligible_display()
            if not_eligible:
                incentive_eligibility = _inel_am
                incentive_amount = 0.0
            elif quota <= 0:
                incentive_eligibility = "No individual quota resolved"
                incentive_amount = 0.0
            elif achievement_pct < AM_MIN_ACHIEVEMENT_FOR_COMMISSION_PCT:
                incentive_eligibility = _inel_am
                incentive_amount = 0.0
            else:
                incentive_eligibility = "Eligible"
                incentive_amount = round(total_paid_amount * (incentive_pct / 100.0), 2)

            incentive_pct = round(float(incentive_pct), 2)

            any_partial = any((x.get("payment_status") or "").upper() == "PARTIALLY_PAID" for x in deal_list)
            payment_status = "PARTIALLY_PAID" if any_partial else "PAID"
            deal_close_dates = [d["close_date"] for d in deal_list if d.get("close_date")]
            row_close_date = max(deal_close_dates) if deal_close_dates else close_date_fallback

            rep_rows.append(
                (
                    user_id,
                    team_id,
                    len(deal_list),
                    total_revenue,
                    total_paid_amount,
                    res_slab_id,
                    incentive_pct,
                    incentive_amount,
                    payment_status,
                    period,
                    row_close_date,
                    incentive_eligibility,
                    quota if quota > 0 else None,
                )
            )
            continue

        if (
            team_name == ENTERPRISE_TEAM_NAME
            and ENTERPRISE_QUOTA_ACHIEVEMENT_ENABLED
            and ent_rep_config_for_name(
                sample.get("deal_owner_name") if sample else None,
                sample.get("deal_owner_email") if sample else None,
            )
        ):
            total_revenue = sum(float(d["amount"]) for d in all_deals_for_rep)
            total_paid_amount = sum(effective_paid_amount_from_status(x) for x in deal_list)
            hubspot_q = sample.get("owner_hubspot_quota_usd") if sample else None
            deal_name = sample.get("deal_owner_name") if sample else None
            deal_email = sample.get("deal_owner_email") if sample else None
            quota = ent_individual_quota_usd_for_rep(
                hubspot_q,
                full_name=deal_name,
                email=deal_email,
                calculation_period=period,
            )
            if quota <= 0:
                achievement_pct = 0.0
                incentive_pct = 0.0
            else:
                achievement_pct = (total_revenue / quota) * 100.0
                incentive_pct = ent_commission_pct_from_quota_achievement(
                    achievement_pct, deal_name, deal_email
                )
            res_slab_id = slabs_default[-1]["slab_id"] if slabs_default else 1
            _inel_ent = ent_rep_incentive_ineligible_display(deal_name, deal_email)
            min_ent = ent_min_achievement_pct_for_rep(deal_name, deal_email)
            if not_eligible:
                incentive_eligibility = _inel_ent
                incentive_amount = 0.0
            elif quota <= 0:
                incentive_eligibility = "No individual quota resolved"
                incentive_amount = 0.0
            elif achievement_pct < min_ent:
                incentive_eligibility = _inel_ent
                incentive_amount = 0.0
            else:
                incentive_eligibility = "Eligible"
                incentive_amount = round(total_paid_amount * (incentive_pct / 100.0), 2)

            incentive_pct = round(float(incentive_pct), 2)

            any_partial = any((x.get("payment_status") or "").upper() == "PARTIALLY_PAID" for x in deal_list)
            payment_status = "PARTIALLY_PAID" if any_partial else "PAID"
            deal_close_dates = [d["close_date"] for d in deal_list if d.get("close_date")]
            row_close_date = max(deal_close_dates) if deal_close_dates else close_date_fallback

            rep_rows.append(
                (
                    user_id,
                    team_id,
                    len(deal_list),
                    total_revenue,
                    total_paid_amount,
                    res_slab_id,
                    incentive_pct,
                    incentive_amount,
                    payment_status,
                    period,
                    row_close_date,
                    incentive_eligibility,
                    quota if quota > 0 else None,
                )
            )
            continue

        # --- Non-SMB or legacy SMB slab ---
        slab_set = resolve_slab_set(team_name, owner_cg)
        slabs = _slabs_for(slab_set)
        total_revenue = sum(float(d["amount"]) for d in all_deals_for_rep)
        total_paid_amount = sum(effective_paid_amount_from_status(x) for x in deal_list)
        res = calculate_incentive_from_paid(total_revenue, total_paid_amount, slabs)
        any_partial = any((x.get("payment_status") or "").upper() == "PARTIALLY_PAID" for x in deal_list)
        payment_status = "PARTIALLY_PAID" if any_partial else "PAID"
        if not_eligible:
            incentive_amount = 0.0
            incentive_eligibility = smb_rep_incentive_ineligible_display()
        else:
            incentive_amount = res.incentive_amount
            incentive_eligibility = "Eligible"
        deal_close_dates = [d["close_date"] for d in deal_list if d.get("close_date")]
        row_close_date = max(deal_close_dates) if deal_close_dates else close_date_fallback
        rep_rows.append(
            (
                user_id,
                team_id,
                len(deal_list),
                res.total_revenue,
                total_paid_amount,
                res.slab_id or 0,
                res.incentive_percentage,
                incentive_amount,
                payment_status,
                period,
                row_close_date,
                incentive_eligibility,
                None,
            )
        )

    teams_with_goals = {t["team_id"]: (float(t["team_goal"]) if t.get("team_goal") is not None else None) for t in get_all_teams(active_only=False)}
    team_meta = {t["team_id"]: t for t in get_all_teams(active_only=False)}
    default_slab_team = slabs_default[0]["slab_id"] if slabs_default else 1
    team_rows = []
    for team_id, total_team_revenue in team_total_amount_all.items():
        lead_id = _get_team_lead(team_id)
        if not lead_id:
            continue
        team_goal = teams_with_goals.get(team_id)
        achievement_pct = (total_team_revenue / team_goal * 100.0) if team_goal and team_goal > 0 else 0.0
        tname = (team_meta.get(team_id) or {}).get("team_name") or ""
        lead_user = get_user_by_id(lead_id)
        lead_cg = (lead_user.get("compensation_group") or "").strip().upper() if lead_user else ""
        if (
            SMB_QUOTA_ACHIEVEMENT_ENABLED
            and tname.strip() == SMB_TEAM_NAME
            and lead_cg == "SMB_CHITRADIP"
        ):
            commission_pct = smb_manager_chitradip_team_commission_pct(achievement_pct)
        elif AM_QUOTA_ACHIEVEMENT_ENABLED and tname.strip() == ACCOUNT_MANAGEMENT_TEAM_NAME:
            commission_pct = am_team_commission_pct_from_team_achievement(achievement_pct)
        else:
            commission_pct = team_commission_pct_from_achievement(achievement_pct)
        incentive_amount = total_team_revenue * (commission_pct / 100.0)
        deal_count = sum(1 for d in deals if d["team_id"] == team_id)
        team_rows.append((
            team_id,
            lead_id,
            deal_count,
            total_team_revenue,
            0.0,
            default_slab_team,
            commission_pct,
            round(incentive_amount, 2),
            "UNPAID",
            period,
        ))

    default_slab = slabs_default[-1]["slab_id"] if slabs_default else 1
    rep_rows = [
        (r[0], r[1], r[2], r[3], r[4], r[5] if r[5] else default_slab, r[6], r[7], r[8], r[9], r[10], r[11], r[12])
        for r in rep_rows
    ]
    team_rows = [
        (r[0], r[1], r[2], r[3], r[4], r[5] if r[5] else default_slab, r[6], r[7], r[8], r[9])
        for r in team_rows
    ]

    insert_rep_incentives(rep_rows)
    insert_team_incentives(team_rows)
    return len(rep_rows), len(team_rows)
