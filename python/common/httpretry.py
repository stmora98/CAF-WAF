"""Retry helpers for direct REST calls (Cost Management, Microsoft Graph, Defender for
Endpoint) that aren't covered by a lightweight synchronous Azure SDK.

Equivalent to the Invoke-RestMethodWithRetry / Invoke-GraphPagedRequest helpers
repeated across Invoke-AzureAdvisor/FinOps/Security-CloudShell.ps1.
"""
from __future__ import annotations

import time
from typing import Any, Dict, List, Optional

import requests
from azure.core.credentials import TokenCredential

_TRANSIENT_STATUS = {408, 429, 500, 502, 503, 504}


def get_bearer_token(credential: TokenCredential, scope: str = "https://management.azure.com/.default") -> str:
    return credential.get_token(scope).token


def request_with_retry(
    method: str,
    url: str,
    headers: Dict[str, str],
    json_body: Optional[dict] = None,
    timeout: int = 60,
    max_attempts: int = 5,
) -> requests.Response:
    attempt = 0
    while True:
        resp = requests.request(method, url, headers=headers, json=json_body, timeout=timeout)
        if resp.status_code not in _TRANSIENT_STATUS:
            resp.raise_for_status()
            return resp

        attempt += 1
        if attempt >= max_attempts:
            resp.raise_for_status()
            return resp
        retry_after = int(resp.headers.get("Retry-After", 0) or 0)
        delay = retry_after if retry_after > 0 else min(30, 2 ** attempt)
        print(f"  API returned HTTP {resp.status_code}. Retrying in {delay}s (attempt {attempt}/{max_attempts})...")
        time.sleep(delay)


def paged_get(url: str, headers: Dict[str, str], max_attempts: int = 5) -> List[Any]:
    """Follows @odata.nextLink pagination for Microsoft Graph / Defender for Endpoint APIs."""
    rows: List[Any] = []
    next_link: Optional[str] = url
    while next_link:
        resp = request_with_retry("GET", next_link, headers=headers, timeout=120, max_attempts=max_attempts)
        payload = resp.json()
        rows.extend(payload.get("value", []))
        next_link = payload.get("@odata.nextLink")
    return rows
