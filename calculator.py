"""
Incentive calculation logic for CloudFuze Migrate.

Revenue-based slabs: see ``commission_policy.DEFAULT_SLAB_THRESHOLDS`` (Q1 2026 policy defaults).

Formula: Incentive Amount = Total Revenue × (Incentive % / 100)
Rep incentives use slab from total amount; incentive $ applies paid amount × slab %.
"""

from dataclasses import dataclass
from typing import List, Optional, Tuple

from commission_policy import DEFAULT_SLAB_THRESHOLDS


@dataclass
class IncentiveResult:
    """Result of incentive calculation."""

    total_revenue: float
    incentive_percentage: float
    incentive_amount: float
    slab_id: Optional[int] = None


# Fallback when DB slabs unavailable (min_revenue threshold -> incentive %)
DEFAULT_SLABS = DEFAULT_SLAB_THRESHOLDS


def get_incentive_percentage(total_revenue: float, slabs: Optional[List[dict]] = None) -> Tuple[float, Optional[int]]:
    """
    Determine incentive percentage for given total revenue.

    Uses highest applicable slab (revenue >= min_revenue).

    Args:
        total_revenue: Sum of deal amounts (float or Decimal from DB).
        slabs: List of dicts with min_revenue, incentive_percentage, slab_id.
               If None, uses DEFAULT_SLABS (slab_id will be None).

    Returns:
        (incentive_percentage, slab_id or None)
    """
    total_revenue = float(total_revenue)
    if slabs:
        for s in slabs:
            min_rev = float(s.get("min_revenue", 0))
            max_rev = s.get("max_revenue")
            if total_revenue >= min_rev:
                if max_rev is None or total_revenue <= float(max_rev):
                    return float(s["incentive_percentage"]), s.get("slab_id")
        return 0.0, None

    for threshold, pct in DEFAULT_SLABS:
        if total_revenue >= threshold:
            return float(pct), None
    return 0.0, None


def calculate_incentive(
    total_revenue: float,
    slabs: Optional[List[dict]] = None,
) -> IncentiveResult:
    """
    Calculate incentive amount from total revenue.

    Incentive Amount = Total Revenue × (Incentive % / 100)

    Args:
        total_revenue: Sum of deal amounts (float or Decimal from DB).
        slabs: Optional slab definitions from DB.

    Returns:
        IncentiveResult with percentage and amount.
    """
    total_revenue = float(total_revenue)
    pct, slab_id = get_incentive_percentage(total_revenue, slabs)
    amount = total_revenue * (pct / 100.0)
    return IncentiveResult(
        total_revenue=round(total_revenue, 2),
        incentive_percentage=pct,
        incentive_amount=round(amount, 2),
        slab_id=slab_id,
    )


def calculate_incentive_from_paid(
    total_revenue: float,
    total_paid_amount: float,
    slabs: Optional[List[dict]] = None,
) -> IncentiveResult:
    """
    Slab % from total Amount (irrespective of payment status).
    Incentive $ from Paid Amount only (in line with payment status).

    Slab tier = f(total_revenue). Incentive amount = total_paid_amount × (slab % / 100).
    """
    total_revenue = float(total_revenue)
    total_paid_amount = float(total_paid_amount)
    pct, slab_id = get_incentive_percentage(total_revenue, slabs)
    amount = total_paid_amount * (pct / 100.0)
    return IncentiveResult(
        total_revenue=round(total_revenue, 2),
        incentive_percentage=pct,
        incentive_amount=round(amount, 2),
        slab_id=slab_id,
    )
