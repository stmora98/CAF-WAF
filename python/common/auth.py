"""Shared Azure authentication and subscription discovery for every phase script.

Equivalent to the Connect-AzAccount + Get-AzSubscription bootstrap shared by all the
Invoke-Azure*-CloudShell.ps1 scripts.
"""
from __future__ import annotations

import os
from typing import List

from azure.core.credentials import TokenCredential
from azure.identity import InteractiveBrowserCredential
from azure.mgmt.subscription import SubscriptionClient


def get_credential() -> TokenCredential:
    """Interactive browser sign-in - the Python equivalent of Connect-AzAccount."""
    return InteractiveBrowserCredential()


def list_enabled_subscriptions(credential: TokenCredential) -> list:
    """Returns Subscription objects (subscription_id, display_name, state) that are Enabled."""
    client = SubscriptionClient(credential)
    return [s for s in client.subscriptions.list() if s.state == "Enabled"]


def resolve_subscription_ids(credential: TokenCredential) -> List[str]:
    """Reads AZWORKSHOP_SUBSCRIPTION_IDS (set by launch_workshop.py for a full run) so every
    phase queries the same tenant-wide scope; falls back to enumerating every enabled
    subscription when a phase script is run standalone.
    """
    env_ids = [s for s in os.environ.get("AZWORKSHOP_SUBSCRIPTION_IDS", "").split(",") if s]
    if env_ids:
        return env_ids
    subs = list_enabled_subscriptions(credential)
    if not subs:
        raise RuntimeError("No enabled subscriptions are accessible in the current tenant.")
    return [s.subscription_id for s in subs]
