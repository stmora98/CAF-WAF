"""Ensures this project's third-party dependencies are installed before any Azure SDK
import runs. Every entry-point script (phase*.py, launch_workshop.py) calls
ensure_dependencies() as the very first thing, before importing anything from azure.*
or common.*, mirroring the PowerShell scripts' self-installing Install-Module pattern.
"""
from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path

_REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements.txt"

# One representative top-level module per requirements.txt entry - enough to detect a
# missing/incomplete environment without importing every submodule individually.
_PROBE_MODULES = [
    "azure.identity",
    "azure.mgmt.subscription",
    "azure.mgmt.resourcegraph",
    "azure.mgmt.monitor",
    "azure.mgmt.managementgroups",
    "azure.keyvault.secrets",
    "azure.keyvault.certificates",
    "openpyxl",
    "requests",
]


def ensure_dependencies() -> None:
    """Installs requirements.txt into the current interpreter if any probe module is missing."""
    missing = []
    for module_name in _PROBE_MODULES:
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(module_name)

    if not missing:
        return

    print(f"Installing missing Python dependencies ({', '.join(missing)})...")
    subprocess.check_call([
        sys.executable, "-m", "pip", "install", "--quiet", "-r", str(_REQUIREMENTS_FILE),
    ])
    importlib.invalidate_caches()
