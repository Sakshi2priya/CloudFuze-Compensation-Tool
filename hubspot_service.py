"""
HubSpot integration for CloudFuze Compensation Tool.

Fetches deals and owners from HubSpot API and returns data in a format
that can be validated and saved as a draft upload (same as Excel upload).
"""

import json
import os
from calendar import monthrange
from functools import lru_cache
from pathlib import Path
from datetime import date, datetime, timezone
from typing import Any, Dict, List, Optional, Set, Tuple

import requests

from excel_service import effective_paid_amount_from_status

HUBSPOT_ACCESS_TOKEN_ENV = "HUBSPOT_ACCESS_TOKEN"
# Internal name of your HubSpot deal property for payment (Settings → Properties → Deals).
HUBSPOT_PAYMENT_STATUS_PROPERTY_ENV = "HUBSPOT_PAYMENT_STATUS_PROPERTY"
HUBSPOT_DEAL_YEAR_ENV = "HUBSPOT_DEAL_YEAR"
HUBSPOT_DEAL_QUARTER_ENV = "HUBSPOT_DEAL_QUARTER"
HUBSPOT_API_BASE = "https://api.hubapi.com"


def quarter_month_span(quarter: int) -> Tuple[int, int]:
    """Calendar quarter: Q1 Jan–Mar, Q2 Apr–Jun, Q3 Jul–Sep, Q4 Oct–Dec. Returns (start_month, end_month)."""
    if quarter == 1:
        return (1, 3)
    if quarter == 2:
        return (4, 6)
    if quarter == 3:
        return (7, 9)
    if quarter == 4:
        return (10, 12)
    raise ValueError(f"quarter must be 1–4, got {quarter}")


def quarter_close_date_bounds_ms(year: int, quarter: int) -> Tuple[int, int]:
    """UTC millisecond range for HubSpot `closedate` filter (inclusive)."""
    sm, em = quarter_month_span(quarter)
    start = datetime(year, sm, 1, 0, 0, 0, tzinfo=timezone.utc)
    last_day = monthrange(year, em)[1]
    end = datetime(year, em, last_day, 23, 59, 59, 999000, tzinfo=timezone.utc)
    return int(start.timestamp() * 1000), int(end.timestamp() * 1000)


def close_date_in_quarter(closedate: Optional[date], year: int, quarter: int) -> bool:
    if closedate is None:
        return False
    if closedate.year != year:
        return False
    sm, em = quarter_month_span(quarter)
    return sm <= closedate.month <= em


def default_deal_year_quarter() -> Tuple[int, int]:
    """Defaults from env or calendar year + Q4."""
    try:
        y = int(os.getenv(HUBSPOT_DEAL_YEAR_ENV, str(datetime.now().year)))
    except ValueError:
        y = datetime.now().year
    try:
        q = int(os.getenv(HUBSPOT_DEAL_QUARTER_ENV, "4"))
    except ValueError:
        q = 4
    if q not in (1, 2, 3, 4):
        q = 4
    return y, q


def get_access_token() -> Optional[str]:
    """Get HubSpot Private App token from env, .env (via load_dotenv in app), or Streamlit secrets."""
    token = os.getenv(HUBSPOT_ACCESS_TOKEN_ENV)
    if token and token.strip():
        return token.strip()
    try:
        import streamlit as st

        t = str(st.secrets[HUBSPOT_ACCESS_TOKEN_ENV]).strip()
        return t or None
    except Exception:
        pass
    return None


def _payment_status_property_name() -> Optional[str]:
    """
    HubSpot deal property **internal name** for payment status (often a custom property).
    If env is set to empty / none / -, the property is not requested (all deals import as UNPAID for status).
    If env is unset, defaults to `payment_status`.
    """
    if HUBSPOT_PAYMENT_STATUS_PROPERTY_ENV not in os.environ:
        return "payment_status"
    v = os.environ.get(HUBSPOT_PAYMENT_STATUS_PROPERTY_ENV, "").strip()
    if not v or v.lower() in ("none", "false", "-"):
        return None
    return v


def deal_properties_csv() -> str:
    """Comma-separated deal properties for HubSpot API (includes optional payment field)."""
    parts = [
        "dealname",
        "amount",
        "closedate",
        "createdate",
        "hubspot_owner_id",
        "hs_is_closed_won",
        "dealstage",
    ]
    pk = _payment_status_property_name()
    if pk:
        parts.append(pk)
    seen: Set[str] = set()
    out: List[str] = []
    for p in parts:
        p = p.strip()
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return ",".join(out)


def fetch_deal_enumeration_options(access_token: str, property_name: str) -> List[Dict[str, str]]:
    """
    Load select/enumeration options from HubSpot (GET /crm/v3/properties/deals/{name}).

    Returns rows with **value** (stored on deals) and **label** (UI text). Skips hidden options.
    """
    if not (property_name or "").strip():
        return []
    url = f"{HUBSPOT_API_BASE}/crm/v3/properties/deals/{property_name.strip()}"
    try:
        r = requests.get(url, headers=_headers(access_token), timeout=30)
        r.raise_for_status()
    except requests.RequestException as e:
        raise RuntimeError(f"HubSpot property '{property_name}' request failed: {e}") from e
    data = r.json()
    out: List[Dict[str, str]] = []
    for o in data.get("options") or []:
        if o.get("hidden"):
            continue
        val = o.get("value")
        if val is None or str(val).strip() == "":
            continue
        out.append(
            {
                "value": str(val).strip(),
                "label": str(o.get("label") if o.get("label") is not None else val).strip(),
            }
        )
    return out


def _resolve_payment_label_from_options(raw: Any, options: Optional[List[Dict[str, str]]]) -> Optional[str]:
    """Match HubSpot stored value to option label (case-insensitive on value)."""
    if raw is None or options is None:
        return None
    rs = str(raw).strip()
    if not rs:
        return None
    for o in options:
        if o.get("value", "").strip().lower() == rs.lower():
            return o.get("label") or o.get("value")
    return None


def _label_text_to_db_payment_status(label: str) -> str:
    """Map HubSpot option **label** (or free text) to DB payment_status."""
    u = str(label).strip().upper().replace("  ", " ")
    if u in ("PAID", "UNPAID", "PARTIALLY_PAID", "REFUNDED", "PARTIALLY_REFUNDED"):
        return u
    # Common HubSpot labels (incl. refunds)
    synonyms = {
        "PAID": "PAID",
        "PARTIALLY PAID": "PARTIALLY_PAID",
        "PARTIALLY_PAID": "PARTIALLY_PAID",
        "NOT PAID": "UNPAID",
        "NOT_PAID": "UNPAID",
        "REFUNDED": "REFUNDED",
        "PARTIALLY REFUNDED": "PARTIALLY_REFUNDED",
        "PARTIALLY_REFUNDED": "PARTIALLY_REFUNDED",
    }
    if u in synonyms:
        return synonyms[u]
    if "REFUND" in u:
        return "PARTIALLY_REFUNDED" if "PARTIAL" in u else "REFUNDED"
    if "PARTIAL" in u and "PAID" in u:
        return "PARTIALLY_PAID"
    if "NOT" in u and "PAID" in u:
        return "UNPAID"
    if u == "PAID" or (u.endswith(" PAID") and "PARTIAL" not in u and "REFUND" not in u):
        return "PAID"
    if "UNPAID" in u or u == "UNPAID":
        return "UNPAID"
    return "UNPAID"


def _normalize_payment_status_for_db(raw: Any, options: Optional[List[Dict[str, str]]] = None) -> str:
    """
    Map HubSpot stored value / labels to PAID, UNPAID, PARTIALLY_PAID, REFUNDED, PARTIALLY_REFUNDED.

    When ``options`` is provided (from ``fetch_deal_enumeration_options``), match by **value** first,
    then map the option **label** to the DB enum.
    """
    if raw is None or (isinstance(raw, str) and not str(raw).strip()):
        return "UNPAID"
    raw_s = str(raw).strip()
    label: Optional[str] = None
    if options:
        label = _resolve_payment_label_from_options(raw_s, options)
    if label:
        return _label_text_to_db_payment_status(label)

    s = raw_s.upper()
    if s in ("PAID", "UNPAID", "PARTIALLY_PAID", "REFUNDED", "PARTIALLY_REFUNDED"):
        return s
    if "REFUND" in s:
        return "PARTIALLY_REFUNDED" if "PARTIAL" in s else "REFUNDED"
    if "PARTIAL" in s and "PAID" in s:
        return "PARTIALLY_PAID"
    if s in ("YES", "COMPLETE", "COMPLETED", "TRUE", "1", "FULLY PAID", "FULLY_PAID"):
        return "PAID"
    if s in ("NO", "PENDING", "FALSE", "0", "NOT PAID", "NOT_PAID", "UN-PAID"):
        return "UNPAID"
    if "UNPAID" in s or "NOT PAID" in s:
        return "UNPAID"
    return "UNPAID"


def _headers(access_token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }


def _normalize_owner_id(value: Any) -> str:
    """
    HubSpot returns hubspot_owner_id as string/int; compare using one canonical form.
    """
    if value is None or value == "":
        return ""
    s = str(value).strip()
    try:
        return str(int(float(s)))
    except (ValueError, TypeError, OverflowError):
        return s


def _coerce_hubspot_owner_ids(hubspot_owner_ids: Optional[List[str]]) -> List[str]:
    """Deduplicate and normalize owner id list for filters."""
    if not hubspot_owner_ids:
        return []
    seen: Set[str] = set()
    out: List[str] = []
    for x in hubspot_owner_ids:
        if x is None or (isinstance(x, str) and not str(x).strip()):
            continue
        n = _normalize_owner_id(x) or str(x).strip()
        if n and n not in seen:
            seen.add(n)
            out.append(n)
    return out


def fetch_owners(access_token: str, include_archived: bool = True) -> List[Dict[str, Any]]:
    """Fetch all HubSpot owners (for mapping deal owner id to email/name). Paginates when present.

    When ``include_archived`` is True (default), also pulls archived (deactivated) owners by
    re-querying with ``?archived=true``. Each archived owner row gets ``_archived: True``
    so downstream UI can label them. Without this, deals belonging to former employees
    whose HubSpot accounts were deactivated would not resolve back to an owner email/name.
    """
    url = f"{HUBSPOT_API_BASE}/crm/v3/owners"
    all_rows: List[Dict[str, Any]] = []

    def _page(extra_params: Dict[str, Any], mark_archived: bool = False) -> None:
        after: Optional[str] = None
        while True:
            params: Dict[str, Any] = {"limit": 100, **extra_params}
            if after:
                params["after"] = after
            try:
                r = requests.get(url, headers=_headers(access_token), params=params, timeout=30)
                r.raise_for_status()
            except requests.RequestException as e:
                raise RuntimeError(f"HubSpot owners request failed: {e}") from e
            data = r.json()
            batch = data.get("results", []) or []
            for o in batch:
                if mark_archived:
                    o["_archived"] = True
            all_rows.extend(batch)
            paging = data.get("paging") or {}
            nxt = (paging.get("next") or {}).get("after")
            if not nxt or not batch:
                break
            after = str(nxt)

    _page({})  # active owners
    if include_archived:
        try:
            _page({"archived": "true"}, mark_archived=True)
        except RuntimeError:
            # Don't fail the whole call if the archived query is blocked by scope; just skip.
            pass

    # De-duplicate by id (archived endpoint can sometimes return overlap on some accounts).
    seen: set = set()
    out: List[Dict[str, Any]] = []
    for o in all_rows:
        oid = str(o.get("id") or "")
        if oid in seen:
            continue
        seen.add(oid)
        out.append(o)
    return out


def _stage_is_closed_won_from_pipeline(stg: Dict[str, Any]) -> bool:
    """
    HubSpot stage metadata: closed-won typically has isClosed true and probability 1.0.
    Closed-lost has isClosed true and probability 0.0.
    """
    meta = stg.get("metadata") or {}
    if str(meta.get("isClosed", "")).lower() not in ("true", "1", "yes"):
        return False
    try:
        prob = float(meta.get("probability") if meta.get("probability") is not None else 0)
    except (TypeError, ValueError):
        prob = 0.0
    return prob >= 0.999


def fetch_deal_pipeline_stages(access_token: str) -> List[Dict[str, Any]]:
    """
    All deal stages across pipelines: id, label (pipeline — stage), is_closed_won_stage for fetch logic.
    """
    url = f"{HUBSPOT_API_BASE}/crm/v3/pipelines/deals"
    try:
        r = requests.get(url, headers=_headers(access_token), timeout=30)
        r.raise_for_status()
        data = r.json()
        out: List[Dict[str, Any]] = []
        for pipe in data.get("results", []):
            plabel = (pipe.get("label") or pipe.get("id") or "").strip()
            for stg in pipe.get("stages", []) or []:
                sid = stg.get("id")
                if sid is None:
                    continue
                slabel = (stg.get("label") or str(sid)).strip()
                out.append(
                    {
                        "stage_id": str(sid),
                        "label": f"{plabel} — {slabel}" if plabel else slabel,
                        "is_closed_won_stage": _stage_is_closed_won_from_pipeline(stg),
                    }
                )
        return out
    except requests.RequestException as e:
        raise RuntimeError(f"HubSpot pipelines request failed: {e}") from e


def fetch_deals(
    access_token: str,
    limit: int = 100,
    after: Optional[str] = None,
) -> tuple[List[Dict[str, Any]], Optional[str]]:
    """
    Fetch deals from HubSpot. Returns (list of deal objects, next_page_after cursor or None).
    Each deal has properties: dealname, amount, closedate, hubspot_owner_id, and optionally
    custom properties for payment_status / team if you use them.
    """
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/deals"
    params: Dict[str, Any] = {
        "limit": limit,
        "properties": deal_properties_csv(),
    }
    if after:
        params["after"] = after
    try:
        r = requests.get(url, headers=_headers(access_token), params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        results = data.get("results", [])
        paging = data.get("paging", {})
        next_cursor = None
        if "next" in paging and "after" in paging["next"]:
            next_cursor = paging["next"]["after"]
        return results, next_cursor
    except requests.RequestException as e:
        raise RuntimeError(f"HubSpot deals request failed: {e}") from e


def fetch_all_deals(access_token: str, max_deals: int = 500) -> List[Dict[str, Any]]:
    """Fetch all deals (paginated) up to max_deals."""
    all_deals: List[Dict[str, Any]] = []
    after: Optional[str] = None
    while len(all_deals) < max_deals:
        batch, after = fetch_deals(access_token, limit=100, after=after)
        all_deals.extend(batch)
        if not after or len(batch) == 0:
            break
    return all_deals[:max_deals]


def _is_closed_won(props: Dict[str, Any]) -> bool:
    """HubSpot returns hs_is_closed_won as string 'true'/'false' or boolean."""
    v = props.get("hs_is_closed_won")
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("true", "1", "yes")


def _append_hubspot_owner_filters(filters: List[Dict[str, Any]], owner_ids: List[str]) -> None:
    """Narrow search to one owner (EQ) or several (IN)."""
    if len(owner_ids) == 1:
        filters.append(
            {
                "propertyName": "hubspot_owner_id",
                "operator": "EQ",
                "value": owner_ids[0],
            }
        )
    elif len(owner_ids) > 1:
        filters.append(
            {
                "propertyName": "hubspot_owner_id",
                "operator": "IN",
                "values": owner_ids,
            }
        )


def search_closed_won_deals(
    access_token: str,
    max_deals: int = 2000,
    year: Optional[int] = None,
    quarter: Optional[int] = None,
    hubspot_owner_ids: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Fetch only Closed Won deals via CRM search (efficient; does not download open/lost deals).
    Optionally restrict by HubSpot close date to a calendar quarter (closedate in that range).
    Optional ``hubspot_owner_ids`` narrows to one or more owners (``IN`` when multiple).

    **Deal stage is not sent to the search API** — filtering by stage is applied in Python only.
    Including dealstage in the API query often returns zero rows (ID mismatch across pipelines).
    """
    y, q = (year, quarter) if year is not None and quarter is not None else default_deal_year_quarter()
    ms_lo, ms_hi = quarter_close_date_bounds_ms(y, q)
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/deals/search"
    oids = _coerce_hubspot_owner_ids(hubspot_owner_ids)

    def _paginated_search(include_owner_in_api: bool) -> List[Dict[str, Any]]:
        all_results: List[Dict[str, Any]] = []
        after: Optional[str] = None
        while len(all_results) < max_deals:
            filters: List[Dict[str, Any]] = [
                {
                    "propertyName": "hs_is_closed_won",
                    "operator": "EQ",
                    "value": "true",
                },
                {
                    "propertyName": "closedate",
                    "operator": "GTE",
                    "value": str(ms_lo),
                },
                {
                    "propertyName": "closedate",
                    "operator": "LTE",
                    "value": str(ms_hi),
                },
            ]
            if include_owner_in_api and oids:
                _append_hubspot_owner_filters(filters, oids)
            body: Dict[str, Any] = {
                "filterGroups": [{"filters": filters}],
                "properties": [p.strip() for p in deal_properties_csv().split(",")],
                "limit": min(100, max_deals - len(all_results)),
            }
            if after:
                body["after"] = after
            r = requests.post(url, headers=_headers(access_token), json=body, timeout=60)
            r.raise_for_status()
            data = r.json()
            batch = data.get("results", [])
            all_results.extend(batch)
            paging = data.get("paging") or {}
            next_after = (paging.get("next") or {}).get("after")
            if not next_after or len(batch) == 0:
                break
            after = str(next_after)
        return all_results[:max_deals]

    if not oids:
        return _paginated_search(include_owner_in_api=False)
    try:
        return _paginated_search(include_owner_in_api=True)
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 400 and len(oids) > 1:
            deals = _paginated_search(include_owner_in_api=False)
            return [d for d in deals if _deal_matches_owner_stage(d, oids, None)]
        raise


def _deal_matches_owner_stage(
    deal: Dict[str, Any],
    hubspot_owner_ids: Optional[List[str]],
    deal_stage_id: Optional[str],
) -> bool:
    props = deal.get("properties", {})
    oids = _coerce_hubspot_owner_ids(hubspot_owner_ids)
    if oids:
        deal_oid = _normalize_owner_id(props.get("hubspot_owner_id"))
        if deal_oid not in set(oids):
            return False
    sid = (deal_stage_id or "").strip()
    if sid and str(props.get("dealstage") or "").strip() != sid:
        return False
    return True


def search_deals_by_stage_date_quarter(
    access_token: str,
    year: int,
    quarter: int,
    deal_stage_id: str,
    hubspot_owner_ids: Optional[List[str]] = None,
    date_property: str = "createdate",
    max_deals: int = 2000,
) -> List[Dict[str, Any]]:
    """
    Deals in a given pipeline stage with ``createdate`` or ``closedate`` in the calendar quarter.

    Use ``createdate`` for open / non-closed-won stages (close date is often empty).
    """
    dp = (date_property or "createdate").strip().lower()
    if dp not in ("createdate", "closedate"):
        dp = "createdate"
    ms_lo, ms_hi = quarter_close_date_bounds_ms(year, quarter)
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/deals/search"
    all_results: List[Dict[str, Any]] = []
    after: Optional[str] = None
    sid = (deal_stage_id or "").strip()
    if not sid:
        return []
    oids = _coerce_hubspot_owner_ids(hubspot_owner_ids)

    def _paginated(include_owner: bool) -> List[Dict[str, Any]]:
        ar: List[Dict[str, Any]] = []
        af: Optional[str] = None
        while len(ar) < max_deals:
            filters: List[Dict[str, Any]] = [
                {"propertyName": "dealstage", "operator": "EQ", "value": sid},
                {"propertyName": dp, "operator": "GTE", "value": str(ms_lo)},
                {"propertyName": dp, "operator": "LTE", "value": str(ms_hi)},
            ]
            if include_owner and oids:
                _append_hubspot_owner_filters(filters, oids)
            body: Dict[str, Any] = {
                "filterGroups": [{"filters": filters}],
                "properties": [p.strip() for p in deal_properties_csv().split(",")],
                "limit": min(100, max_deals - len(ar)),
            }
            if af:
                body["after"] = af
            r = requests.post(url, headers=_headers(access_token), json=body, timeout=60)
            r.raise_for_status()
            data = r.json()
            batch = data.get("results", [])
            ar.extend(batch)
            paging = data.get("paging") or {}
            next_after = (paging.get("next") or {}).get("after")
            if not next_after or len(batch) == 0:
                break
            af = str(next_after)
        return ar[:max_deals]

    if not oids:
        return _paginated(include_owner=False)
    try:
        out = _paginated(include_owner=True)
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 400 and len(oids) > 1:
            out = _paginated(include_owner=False)
            return [d for d in out if _deal_matches_owner_stage(d, oids, None)]
        raise
    if not out and oids:
        out = _paginated(include_owner=False)
        out = [d for d in out if _deal_matches_owner_stage(d, oids, None)]
    return out


def fetch_closed_won_deals(
    access_token: str,
    max_deals: int = 2000,
    year: Optional[int] = None,
    quarter: Optional[int] = None,
    hubspot_owner_ids: Optional[List[str]] = None,
    deal_stage_id: Optional[str] = None,
    stage_is_closed_won: Optional[bool] = None,
) -> List[Dict[str, Any]]:
    """
    Fetch deals for import.

    - **No deal stage filter** (Any stage): closed-won deals with **close date** in the quarter.
    - **Closed-won stage** selected: same search, then filter to the exact ``dealstage`` id.
    - **Other stages** (negotiation, etc.): ``dealstage`` + **created** date in quarter (close date is often empty).

    ``stage_is_closed_won`` comes from HubSpot pipeline metadata (see ``fetch_deal_pipeline_stages``).
    """
    y, q = (year, quarter) if year is not None and quarter is not None else default_deal_year_quarter()
    sid = (deal_stage_id or "").strip() or None
    oids = _coerce_hubspot_owner_ids(hubspot_owner_ids)

    def _in_quarter_window(deal: Dict[str, Any]) -> bool:
        props = deal.get("properties", {})
        d = _parse_closedate(props.get("closedate"))
        return close_date_in_quarter(d, y, q)

    # Open / mid-funnel stage: search by stage + createdate in quarter
    if sid and stage_is_closed_won is False:
        try:
            return search_deals_by_stage_date_quarter(
                access_token,
                y,
                q,
                sid,
                hubspot_owner_ids=oids if oids else None,
                date_property="createdate",
                max_deals=max_deals,
            )
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code in (400, 403, 404):
                all_deals = fetch_all_deals(access_token, max_deals=max_deals)
                return [
                    d
                    for d in all_deals
                    if _deal_matches_owner_stage(d, oids if oids else None, sid)
                    and _createdate_in_quarter(d.get("properties", {}), y, q)
                ]
            raise

    try:
        deals = search_closed_won_deals(
            access_token,
            max_deals=max_deals,
            year=y,
            quarter=q,
            hubspot_owner_ids=oids if oids else None,
        )
        if not deals and oids:
            deals = search_closed_won_deals(
                access_token,
                max_deals=max_deals,
                year=y,
                quarter=q,
                hubspot_owner_ids=None,
            )
            deals = [d for d in deals if _deal_matches_owner_stage(d, oids, None)]
        if sid:
            deals = [d for d in deals if _deal_matches_owner_stage(d, None, sid)]
        return deals
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code in (400, 403, 404):
            all_deals = fetch_all_deals(access_token, max_deals=max_deals)
            return [
                d
                for d in all_deals
                if _is_closed_won(d.get("properties", {}))
                and _in_quarter_window(d)
                and _deal_matches_owner_stage(d, oids if oids else None, sid)
            ]
        raise


def _parse_closedate(raw: Any) -> Optional[date]:
    """
    Parse HubSpot `closedate` for deals.

    HubSpot may return:
    - Milliseconds since epoch (string or number), sometimes with decimals
    - ISO 8601 strings (e.g. ``2019-12-07T16:50:06.678Z``) — common on read APIs
    - `date` / `datetime` if already parsed
    """
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, date) and not isinstance(raw, datetime):
        return raw
    if isinstance(raw, datetime):
        return raw.date()

    # JSON number (int/float ms)
    if isinstance(raw, (int, float)):
        ts = float(raw)
        if ts >= 1e12:
            ts = ts / 1000.0
        try:
            return datetime.fromtimestamp(ts, tz=timezone.utc).date()
        except (ValueError, OSError, OverflowError):
            return None

    s = str(raw).strip()
    if not s:
        return None

    # ISO 8601 (HubSpot docs show this format for closedate on create; reads often match)
    if "T" in s or (len(s) >= 10 and s[4] == "-" and s[7] == "-"):
        try:
            iso = s.replace("Z", "+00:00") if s.endswith("Z") else s
            if "T" in iso:
                dt = datetime.fromisoformat(iso)
                if dt.tzinfo is not None:
                    dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
                return dt.date()
            return date.fromisoformat(iso[:10])
        except ValueError:
            pass

    # Unix time: HubSpot stores close as ms; epoch seconds are ~1e9, ms ~1e12+
    try:
        ts = float(s)
    except (ValueError, TypeError):
        return None
    if ts >= 1e12:
        ts = ts / 1000.0
    try:
        return datetime.fromtimestamp(ts, tz=timezone.utc).date()
    except (ValueError, OSError, OverflowError):
        return None


def _createdate_in_quarter(props: Dict[str, Any], year: int, quarter: int) -> bool:
    """True if HubSpot ``createdate`` parses to a date in the calendar quarter."""
    d = _parse_closedate(props.get("createdate"))
    return close_date_in_quarter(d, year, quarter)


def map_hubspot_deals_to_parsed(
    deals: List[Dict[str, Any]],
    owners: List[Dict[str, Any]],
    allowed_emails: Optional[Set[str]] = None,
    hubspot_owner_ids: Optional[List[str]] = None,
    deal_stage_id: Optional[str] = None,
    require_closed_won: bool = True,
    payment_status_options: Optional[List[Dict[str, str]]] = None,
) -> List[Dict[str, Any]]:
    """
    Map HubSpot deal objects + owners to a list of dicts compatible with Compensation Tool:
    deal_name, deal_owner (email or full name), amount, payment_status, team (filled by app from user),
    paid_amount, close_date, incentive_eligibility.

    If allowed_emails is set, only deals whose HubSpot owner email (case-insensitive) is in that set
    are included — i.e. owners that exist in User Management in this tool.

    Team is not in HubSpot response; the app will resolve it from deal_owner (email) -> user -> team.

    Quarter / close date range is **not** filtered again here: `fetch_closed_won_deals` already applies
    the same window via HubSpot search (or the list fallback). Re-filtering by parsed `closedate` in UTC
    often dropped valid rows (timezone edges, missing property on read).

    ``require_closed_won``: set False when importing open-pipeline deals (stage + createdate path).
    """
    owner_by_id: Dict[str, Dict[str, Any]] = {}
    for o in owners:
        k = _normalize_owner_id(o.get("id"))
        if k:
            owner_by_id[k] = o
    out: List[Dict[str, Any]] = []
    for d in deals:
        props = d.get("properties", {})
        if require_closed_won and not _is_closed_won(props):
            continue
        closedate = _parse_closedate(props.get("closedate"))
        if not _deal_matches_owner_stage(d, hubspot_owner_ids, deal_stage_id):
            continue
        dealname = (props.get("dealname") or "").strip() or f"Deal {d.get('id', '')}"
        amount_val = props.get("amount")
        try:
            amount = float(amount_val) if amount_val is not None else 0.0
        except (TypeError, ValueError):
            amount = 0.0
        owner_id = props.get("hubspot_owner_id")
        owner_email = None
        owner_name = None
        oid_key = _normalize_owner_id(owner_id)
        if oid_key and oid_key in owner_by_id:
            o = owner_by_id[oid_key]
            owner_email = o.get("email")
            first = o.get("firstName", "") or ""
            last = o.get("lastName", "") or ""
            owner_name = f"{first} {last}".strip() or owner_email
        if allowed_emails is not None:
            em = (owner_email or "").strip().lower()
            if not em or em not in allowed_emails:
                continue
        deal_owner = owner_email or owner_name or ""
        pay_key = _payment_status_property_name()
        pay_raw = props.get(pay_key) if pay_key else None
        if pay_raw is None:
            pay_raw = props.get("payment_status")
        opts = payment_status_options
        payment_status = _normalize_payment_status_for_db(pay_raw, opts)
        ps_label = _resolve_payment_label_from_options(pay_raw, opts)
        if not ps_label and pay_raw is not None and str(pay_raw).strip():
            ps_label = str(pay_raw).strip()
        if not ps_label:
            ps_label = payment_status.replace("_", " ").title()
        paid_amount = effective_paid_amount_from_status(
            {"amount": amount, "payment_status": payment_status, "paid_amount": 0.0}
        )
        out.append({
            "deal_name": dealname,
            "deal_owner": deal_owner,
            "amount": amount,
            "payment_status": payment_status,
            "hubspot_payment_value": str(pay_raw) if pay_raw is not None else "",
            "payment_status_label": ps_label,
            "paid_amount": paid_amount,
            "close_date": closedate,
            "incentive_eligibility": "Eligible",
        })
    return out


def fetch_and_map_hubspot_deals(
    access_token: str,
    allowed_emails: Optional[Set[str]] = None,
    year: Optional[int] = None,
    quarter: Optional[int] = None,
    hubspot_owner_ids: Optional[List[str]] = None,
    deal_stage_id: Optional[str] = None,
    stage_is_closed_won: Optional[bool] = None,
    payment_status_filter: Optional[str] = None,
    payment_hubspot_value_filter: Optional[str] = None,
) -> Tuple[List[Dict[str, Any]], Dict[str, int], List[Dict[str, Any]]]:
    """
    Fetch Closed Won deals and owners from HubSpot and return deal dicts for the Compensation Tool.

    Returns a 3-tuple:
    - **mapped_for_tool**: deals whose owner email is in User Management (for validation / save).
    - **stats**: counts (total closed won from API, included in tool, skipped).
    - **all_closed_mapped**: every deal in range (no User Management filter), for display.

    ``stage_is_closed_won``: from pipeline metadata; when False, deals are open-pipeline (createdate quarter).
    ``payment_hubspot_value_filter``: exact HubSpot **enumeration value** (from property options API).
    ``payment_status_filter``: legacy normalized filter ``PAID`` / ``UNPAID`` / ``PARTIALLY_PAID`` / ``REFUNDED`` / etc.
    """
    y, q = (year, quarter) if year is not None and quarter is not None else default_deal_year_quarter()
    oid_list = _coerce_hubspot_owner_ids(hubspot_owner_ids)
    sid = (deal_stage_id or "").strip() or None
    require_won = not (sid and stage_is_closed_won is False)
    # Distinguish ``None`` (no owner filter) from ``[]`` (explicit: no matching owners, e.g. empty team roster).
    if hubspot_owner_ids is not None and not oid_list:
        stats = {
            "total_in_hubspot": 0,
            "included": 0,
            "skipped": 0,
            "year": y,
            "quarter": q,
        }
        return [], stats, []
    owners = fetch_owners(access_token)
    pay_prop = _payment_status_property_name()
    payment_status_options: Optional[List[Dict[str, str]]] = None
    if pay_prop:
        try:
            payment_status_options = fetch_deal_enumeration_options(access_token, pay_prop)
        except Exception:
            payment_status_options = None

    deals = fetch_closed_won_deals(
        access_token,
        year=y,
        quarter=q,
        hubspot_owner_ids=oid_list if oid_list else None,
        deal_stage_id=sid,
        stage_is_closed_won=stage_is_closed_won,
    )

    def _filter_payment(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        hv = (payment_hubspot_value_filter or "").strip()
        if hv:
            return [r for r in rows if (r.get("hubspot_payment_value") or "").strip() == hv]
        if not payment_status_filter:
            return rows
        want = payment_status_filter.strip().upper()
        return [r for r in rows if (r.get("payment_status") or "").strip().upper() == want]

    all_closed_mapped = _filter_payment(
        map_hubspot_deals_to_parsed(
            deals,
            owners,
            allowed_emails=None,
            hubspot_owner_ids=oid_list if oid_list else None,
            deal_stage_id=sid,
            require_closed_won=require_won,
            payment_status_options=payment_status_options,
        )
    )
    mapped = _filter_payment(
        map_hubspot_deals_to_parsed(
            deals,
            owners,
            allowed_emails=allowed_emails,
            hubspot_owner_ids=oid_list if oid_list else None,
            deal_stage_id=sid,
            require_closed_won=require_won,
            payment_status_options=payment_status_options,
        )
    )
    total = len(deals)
    included = len(mapped)
    stats = {
        "total_in_hubspot": total,
        "included": included,
        "skipped": max(0, total - included) if allowed_emails is not None else 0,
        "year": y,
        "quarter": q,
    }
    return mapped, stats, all_closed_mapped


def _hubspot_team_filters_path() -> Path:
    env = (os.environ.get("HUBSPOT_TEAM_OWNER_FILTERS_JSON") or "").strip()
    if env:
        return Path(env)
    return Path(__file__).resolve().parent / "policy" / "hubspot_team_owner_filters.json"


@lru_cache(maxsize=1)
def load_hubspot_team_owner_filters() -> Dict[str, List[str]]:
    """
    Name tokens per app team (SMB, Enterprise, Account Management) used to filter HubSpot owner labels.

    Edit ``policy/hubspot_team_owner_filters.json`` or set ``HUBSPOT_TEAM_OWNER_FILTERS_JSON`` to a JSON path.
    Matching is case-insensitive substring on the full owner label (name + email).
    """
    p = _hubspot_team_filters_path()
    if not p.is_file():
        return {}
    try:
        with p.open(encoding="utf-8") as f:
            raw = json.load(f)
        return {str(k).strip(): [str(x).strip().lower() for x in (v or [])] for k, v in raw.items()}
    except (OSError, json.JSONDecodeError):
        return {}


def filter_hubspot_owner_labels_for_team(
    labels: List[str],
    team: str,
    extra_tokens: Optional[List[str]] = None,
) -> List[str]:
    """
    Restrict HubSpot owner dropdown labels to those matching roster tokens for the selected team.

    ``team`` is **Any** or a key from the JSON (e.g. **SMB**). Tokens match as substrings in ``label.lower()``).
    ``extra_tokens``: optional substrings (e.g. from User Management emails for that team) merged with JSON tokens.
    """
    if not team or str(team).strip() == "Any":
        return list(labels)
    key = str(team).strip()
    tokens = list(load_hubspot_team_owner_filters().get(key, []))
    for x in extra_tokens or []:
        s = str(x).strip().lower()
        if s and s not in tokens:
            tokens.append(s)
    if not tokens:
        return list(labels)
    out: List[str] = []
    for lbl in labels:
        low = lbl.lower()
        if any(t in low for t in tokens):
            out.append(lbl)
    return out


def _parse_hubspot_datetime_ms_or_iso(raw: Any) -> Optional[datetime]:
    """Parse HubSpot goal_target date fields (ISO string or ms timestamp)."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        try:
            ts = float(raw)
            if ts > 1e12:
                ts = ts / 1000.0
            return datetime.fromtimestamp(ts, tz=timezone.utc)
        except (ValueError, OSError, OverflowError):
            return None
    s = str(raw).strip()
    if not s:
        return None
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except ValueError:
        return None


def _quarter_bounds_utc(year: int, quarter: int) -> Tuple[datetime, datetime]:
    """Inclusive quarter window in UTC for overlap checks with goal slices."""
    sm, em = quarter_month_span(quarter)
    start = datetime(year, sm, 1, 0, 0, 0, tzinfo=timezone.utc)
    last_day = monthrange(year, em)[1]
    end = datetime(year, em, last_day, 23, 59, 59, tzinfo=timezone.utc)
    return start, end


def _intervals_overlap_utc(
    a0: datetime,
    a1: datetime,
    b0: datetime,
    b1: datetime,
) -> bool:
    return a0 <= b1 and a1 >= b0


GOAL_TARGET_SEARCH_PROPERTIES = (
    "hs_goal_name",
    "hs_target_amount",
    "hubspot_owner_id",
    "hs_start_datetime",
    "hs_end_datetime",
    "hs_goal_type",
)


def search_goal_targets_for_year(
    access_token: str,
    year: int,
    limit: int = 100,
) -> List[Dict[str, Any]]:
    """
    Search CRM ``goal_targets`` for slices whose ``hs_start_datetime`` falls in the calendar year.

    Same underlying objects as HubSpot **Sales Goals** / forecast-style quota targets (per-user targets use
    ``hubspot_owner_id``). Requires Private App scope to read goals (e.g. ``crm.objects.goals.read``).
    """
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/goal_targets/search"
    lo = f"{year}-01-01T00:00:00.000Z"
    hi = f"{year + 1}-01-01T00:00:00.000Z"
    body: Dict[str, Any] = {
        "filterGroups": [
            {
                "filters": [
                    {"propertyName": "hs_start_datetime", "operator": "GTE", "value": lo},
                    {"propertyName": "hs_start_datetime", "operator": "LT", "value": hi},
                ]
            }
        ],
        "properties": list(GOAL_TARGET_SEARCH_PROPERTIES),
        "limit": min(100, max(1, limit)),
    }
    out: List[Dict[str, Any]] = []
    after: Optional[str] = None
    try:
        while True:
            if after:
                body["after"] = after
            r = requests.post(url, headers=_headers(access_token), json=body, timeout=90)
            r.raise_for_status()
            data = r.json()
            out.extend(data.get("results") or [])
            paging = data.get("paging") or {}
            nxt = (paging.get("next") or {}).get("after")
            if not nxt:
                break
            after = str(nxt)
    except requests.RequestException as e:
        raise RuntimeError(
            "HubSpot goal_targets search failed (check Private App scopes include Goals read, e.g. crm.objects.goals.read): "
            f"{e}"
        ) from e
    return out


def search_goal_targets_overlapping_quarter(
    access_token: str,
    year: int,
    quarter: int,
    limit: int = 100,
) -> List[Dict[str, Any]]:
    """
    Search CRM ``goal_targets`` whose date range *overlaps* the given calendar quarter.

    Two slices are kept: any goal where ``hs_start_datetime <= quarter_end`` AND ``hs_end_datetime >= quarter_start``.
    Use this instead of :func:`search_goal_targets_for_year` for quarterly Sales Goals so that goals starting
    in late December of the previous year (or spanning quarter boundaries) are not missed.
    """
    q0, q1 = _quarter_bounds_utc(year, quarter)
    q0_str = q0.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    q1_str = q1.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    url = f"{HUBSPOT_API_BASE}/crm/v3/objects/goal_targets/search"
    body: Dict[str, Any] = {
        "filterGroups": [
            {
                "filters": [
                    {"propertyName": "hs_start_datetime", "operator": "LTE", "value": q1_str},
                    {"propertyName": "hs_end_datetime", "operator": "GTE", "value": q0_str},
                ]
            }
        ],
        "properties": list(GOAL_TARGET_SEARCH_PROPERTIES),
        "limit": min(100, max(1, limit)),
    }
    out: List[Dict[str, Any]] = []
    after: Optional[str] = None
    try:
        while True:
            if after:
                body["after"] = after
            r = requests.post(url, headers=_headers(access_token), json=body, timeout=90)
            r.raise_for_status()
            data = r.json()
            out.extend(data.get("results") or [])
            paging = data.get("paging") or {}
            nxt = (paging.get("next") or {}).get("after")
            if not nxt:
                break
            after = str(nxt)
    except requests.RequestException as e:
        raise RuntimeError(
            "HubSpot goal_targets search failed (check Private App scopes include Goals read, e.g. crm.objects.goals.read): "
            f"{e}"
        ) from e
    return out


_REVENUE_GOAL_TYPE_TOKENS = ("REVENUE", "DEAL_REVENUE", "CLOSED_WON_REVENUE", "AMOUNT")


def _is_revenue_goal_type(goal_type: Any) -> bool:
    """Return True if the HubSpot ``hs_goal_type`` looks like a revenue/$ goal (vs deal-count, calls, etc.)."""
    if goal_type is None:
        return True  # be permissive when HubSpot omits the field
    s = str(goal_type).strip().upper()
    if not s:
        return True
    return any(tok in s for tok in _REVENUE_GOAL_TYPE_TOKENS)


def aggregate_goal_targets_by_hubspot_owner_for_quarter(
    goal_results: List[Dict[str, Any]],
    year: int,
    quarter: int,
) -> Dict[str, float]:
    """
    Sum ``hs_target_amount`` per ``hubspot_owner_id`` for goal slices whose time range **overlaps**
    the calendar quarter. Only rows with a non-empty ``hubspot_owner_id`` are included (user-scoped goals).
    """
    q0, q1 = _quarter_bounds_utc(year, quarter)
    totals: Dict[str, float] = {}
    for row in goal_results:
        props = row.get("properties") or {}
        oid = str(props.get("hubspot_owner_id") or "").strip()
        if not oid:
            continue
        if not _is_revenue_goal_type(props.get("hs_goal_type")):
            continue
        try:
            amt = float(props.get("hs_target_amount") or 0)
        except (TypeError, ValueError):
            amt = 0.0
        if amt == 0:
            continue
        st = _parse_hubspot_datetime_ms_or_iso(props.get("hs_start_datetime"))
        en = _parse_hubspot_datetime_ms_or_iso(props.get("hs_end_datetime"))
        if st is None:
            continue
        if en is None:
            en = st
        if not _intervals_overlap_utc(st, en, q0, q1):
            continue
        totals[oid] = totals.get(oid, 0.0) + amt
    return totals


def map_hubspot_owner_ids_to_emails(access_token: str) -> Dict[str, str]:
    """HubSpot owner id (str) -> lowercased email."""
    owners = fetch_owners(access_token)
    out: Dict[str, str] = {}
    for o in owners:
        oid = str(o.get("id", "") or "").strip()
        em = (o.get("email") or "").strip().lower()
        if oid and em:
            out[oid] = em
    return out


def build_hubspot_goal_sync_plan(
    access_token: str,
    year: int,
    quarter: int,
    app_users: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Pull goal targets from HubSpot for ``year``/``quarter`` and match to app users by owner email.

    Returns keys: ``period_label``, ``by_user_id`` (dict user_id -> target float), ``matched_rows``,
    ``hubspot_owners_with_goals_no_user`` (list of dicts), ``users_without_goals`` (list of dicts).
    """
    period_label = f"{year}-Q{quarter}"
    raw = search_goal_targets_overlapping_quarter(access_token, year, quarter)
    by_oid = aggregate_goal_targets_by_hubspot_owner_for_quarter(raw, year, quarter)
    id_to_email = map_hubspot_owner_ids_to_emails(access_token)
    email_to_user: Dict[str, Dict[str, Any]] = {}
    for u in app_users:
        em = (u.get("email") or "").strip().lower()
        if em:
            email_to_user[em] = u

    by_user_id: Dict[int, float] = {}
    seen_owner_no_user: List[Dict[str, Any]] = []

    for oid, amt in by_oid.items():
        em = id_to_email.get(oid)
        if not em:
            seen_owner_no_user.append({"hubspot_owner_id": oid, "target_usd": amt, "reason": "no_email_in_owners_api"})
            continue
        u = email_to_user.get(em)
        if not u:
            seen_owner_no_user.append({"hubspot_owner_id": oid, "email": em, "target_usd": amt, "reason": "no_matching_app_user"})
            continue
        uid = int(u["user_id"])
        by_user_id[uid] = by_user_id.get(uid, 0.0) + float(amt)

    user_by_id = {int(u["user_id"]): u for u in app_users}
    matched_rows: List[Dict[str, Any]] = []
    for uid, total in sorted(by_user_id.items(), key=lambda x: x[0]):
        u = user_by_id.get(uid)
        if u:
            matched_rows.append(
                {
                    "user_id": uid,
                    "full_name": u.get("full_name"),
                    "email": (u.get("email") or "").strip().lower(),
                    "target_usd": float(total),
                }
            )

    users_with_target = set(by_user_id.keys())
    users_without_goals = [
        {"user_id": u["user_id"], "full_name": u.get("full_name"), "email": u.get("email")}
        for u in app_users
        if int(u["user_id"]) not in users_with_target
    ]

    return {
        "period_label": period_label,
        "by_user_id": by_user_id,
        "matched_rows": matched_rows,
        "hubspot_owners_with_goals_no_user": seen_owner_no_user,
        "users_without_goals": users_without_goals,
        "raw_goal_target_count": len(raw),
    }
