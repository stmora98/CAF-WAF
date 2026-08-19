"""Azure Resource Graph query runner with pagination and throttling retry.

Equivalent to the Invoke-SearchAzGraphWithRetry / Run-Query helpers repeated across
every Invoke-Azure*-CloudShell.ps1 script.
"""
from __future__ import annotations

import random
import time
from typing import Any, Dict, List, Optional

from azure.core.credentials import TokenCredential
from azure.core.exceptions import HttpResponseError
from azure.mgmt.resourcegraph import ResourceGraphClient
from azure.mgmt.resourcegraph.models import QueryRequest, QueryRequestOptions

_TRANSIENT_STATUS = {408, 429, 500, 502, 503, 504}

# One ResourceGraphClient per credential, reused across every call/thread instead of
# reconnecting per query - the SDK client is safe for concurrent reads.
_clients: Dict[int, ResourceGraphClient] = {}


def _get_client(credential: TokenCredential) -> ResourceGraphClient:
    key = id(credential)
    client = _clients.get(key)
    if client is None:
        client = ResourceGraphClient(credential)
        _clients[key] = client
    return client


def run_query(
    credential: TokenCredential,
    subscription_ids: List[str],
    query: str,
    page_size: int = 1000,
    max_attempts: int = 6,
) -> List[Dict[str, Any]]:
    """Runs a KQL query across subscriptions, paging via skip_token until exhausted,
    retrying transient 429/5xx throttling with exponential backoff plus jitter (avoids
    many parallel callers retrying in lockstep) - matches the PowerShell scripts'
    Invoke-SearchAzGraphWithRetry behavior.
    """
    client = _get_client(credential)
    rows: List[Dict[str, Any]] = []
    skip_token: Optional[str] = None

    while True:
        options = QueryRequestOptions(top=page_size, skip_token=skip_token, result_format="objectArray")
        request = QueryRequest(query=query, subscriptions=subscription_ids, options=options)

        attempt = 0
        while True:
            try:
                response = client.resources(request)
                break
            except HttpResponseError as exc:
                attempt += 1
                status = exc.status_code or 0
                if attempt >= max_attempts or status not in _TRANSIENT_STATUS:
                    raise
                delay = min(60, 2 ** attempt) + random.uniform(0, 1)
                print(f"  Resource Graph throttled or unavailable (HTTP {status}). "
                      f"Retrying in {delay:.1f}s (attempt {attempt}/{max_attempts})...")
                time.sleep(delay)

        rows.extend(response.data or [])
        skip_token = response.skip_token
        if not skip_token:
            break

    return rows
