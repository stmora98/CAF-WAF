"""Best-effort Azure Retail Prices API lookups, used to estimate the current monthly
cost of a right-sizing candidate and the potential savings from moving to a smaller SKU.
The Retail Prices API is public/unauthenticated - no credential is needed.

This is intentionally best-effort: retail (list) pricing, no reservations/Hybrid Benefit/
negotiated discounts applied, so treat the numbers as estimates, not an invoice.
"""
from __future__ import annotations

from typing import Optional, Tuple

import requests

_RETAIL_PRICES_URL = "https://prices.azure.com/api/retail/prices"
_HOURS_PER_MONTH = 730
_price_cache: dict = {}

# Ordered cheapest -> priciest within each App Service Plan SKU family. Used to suggest a
# smaller SKU for idle/underutilized plans without guessing across unrelated families.
_APP_SERVICE_SKU_FAMILIES = [
    ["B1", "B2", "B3"],
    ["S1", "S2", "S3"],
    ["P1V2", "P2V2", "P3V2"],
    ["P0V3", "P1V3", "P2V3", "P3V3"],
    ["P0MV3", "P1MV3", "P2MV3", "P3MV3"],
    ["I1V2", "I2V2", "I3V2"],
    ["I1MV2", "I2MV2", "I3MV2"],
]


def suggest_smaller_app_service_sku(sku_name: Optional[str]) -> Optional[str]:
    """Returns the next cheaper SKU in the same family (e.g. P2v3 -> P1v3), or None if
    sku_name is already the cheapest in its family or isn't a recognized family."""
    if not sku_name:
        return None
    normalized = sku_name.strip().upper()
    for family in _APP_SERVICE_SKU_FAMILIES:
        if normalized in family:
            idx = family.index(normalized)
            return family[idx - 1] if idx > 0 else None
    return None


def _fetch_hourly_price(service_name: str, sku_name: str, region: str) -> Tuple[Optional[float], Optional[str]]:
    if not sku_name or not region:
        return None, None
    filt = (
        f"serviceName eq '{service_name}' and armRegionName eq '{region}' "
        f"and armSkuName eq '{sku_name}' and priceType eq 'Consumption'"
    )
    try:
        resp = requests.get(_RETAIL_PRICES_URL, params={"$filter": filt}, timeout=20)
        resp.raise_for_status()
        items = resp.json().get("Items", [])
    except Exception:
        return None, None
    prices = [it["retailPrice"] for it in items if it.get("retailPrice") is not None]
    if not prices:
        return None, None
    return min(prices), items[0].get("currencyCode", "USD")


def get_app_service_plan_monthly_cost(sku_name: Optional[str], region: Optional[str], workers: Optional[int] = 1) -> Tuple[Optional[float], Optional[str]]:
    """Returns (estimated monthly retail cost, currency) for an App Service Plan SKU in a
    region, scaled by instance/worker count, or (None, None) if pricing couldn't be found."""
    if not sku_name or not region:
        return None, None
    key = (sku_name.strip().upper(), region.strip().lower())
    if key not in _price_cache:
        _price_cache[key] = _fetch_hourly_price("Azure App Service", key[0], key[1])
    hourly, currency = _price_cache[key]
    if hourly is None:
        return None, None
    return hourly * _HOURS_PER_MONTH * max(int(workers or 1), 1), currency
