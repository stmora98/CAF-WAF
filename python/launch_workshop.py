"""Azure WAF/CAF Workshop Launcher - Python port of Launch-AzureWorkshop.ps1.

Runs all discovery phases and consolidates output:
  1. phase1_discovery      (Resource inventory)
  2. phase2_advisor        (Advisor recommendations)
  2b. phase2b_finops       (FinOps extended export)
  3. phase3_metrics        (Right-sizing & reliability)
  4. phase4_governance     (Governance data export)
  5. phase5_security       (Defender posture + incidents)
  6. phase6_checklists     (Azure/review-checklists ARG compliance)
  7. generate-dashboard.py (Consolidated dashboard + action items - reused as-is)

All outputs are consolidated into a single timestamped folder alongside this script.

Usage:
    python launch_workshop.py [--output-dir DIR] [--skip-metrics] [--skip-finops]
                               [--keep-previous] [--skip-graph-security]
                               [--subscription-ids ID1,ID2,...]
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import bootstrap
bootstrap.ensure_dependencies()

import phase1_discovery
import phase2_advisor
import phase2b_finops
import phase3_metrics
import phase4_governance
import phase5_security
import phase6_checklists
from common.auth import get_credential, list_enabled_subscriptions

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent


def _is_valid_xlsx(path: Path) -> bool:
    """Checks the local file's zip signature to catch OneDrive-encrypted placeholders
    (matches the PowerShell launcher's OLE-header validation)."""
    try:
        with open(path, "rb") as fh:
            header = fh.read(4)
        return header[:2] == b"PK"
    except OSError:
        return False


def _newest_xlsx(folder: Path):
    files = sorted(folder.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _gui_subscription_picker(all_subs: list):
    """Tkinter multi-select dialog. Returns the selected subscriptions, an empty list if the
    user picked nothing / closed the window (caller defaults to ALL), or None if no GUI/display
    is available (e.g. Cloud Shell) - the caller falls back to the text picker on None."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        return None
    try:
        root = tk.Tk()
    except tk.TclError:
        return None

    root.title("Azure WAF/CAF Workshop - Select Subscriptions")
    root.minsize(480, 420)
    root.geometry("560x520")

    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except tk.TclError:
        pass
    style.configure("Sub.TLabel", foreground="#5c5c5c", font=("Segoe UI", 8))
    style.configure("Count.TLabel", foreground="#0f6cbd", font=("Segoe UI", 9, "bold"))
    style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    result: list = []
    check_vars = {s.subscription_id: tk.BooleanVar(value=False) for s in all_subs}

    header = ttk.Frame(root, padding=(16, 16, 16, 8))
    header.pack(fill="x")
    ttk.Label(header, text="Select subscriptions to scan", font=("Segoe UI", 13, "bold")).pack(anchor="w")
    ttk.Label(header, style="Sub.TLabel", wraplength=520, justify="left",
              text="Choose which subscriptions this run should assess. Leave everything "
                   "unchecked and click OK to scan ALL.").pack(anchor="w", pady=(2, 0))

    toolbar = ttk.Frame(root, padding=(16, 0, 16, 8))
    toolbar.pack(fill="x")
    filter_var = tk.StringVar()
    ttk.Entry(toolbar, textvariable=filter_var, width=24).pack(side="left")
    ttk.Label(toolbar, text="filter", style="Sub.TLabel").pack(side="left", padx=(6, 0))
    count_label = ttk.Label(toolbar, text=f"0 of {len(all_subs)} selected", style="Count.TLabel")
    count_label.pack(side="right")

    list_frame = ttk.Frame(root, padding=(16, 0, 16, 0))
    list_frame.pack(fill="both", expand=True)
    canvas = tk.Canvas(list_frame, highlightthickness=0, bd=0)
    vscroll = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=vscroll.set)
    canvas.pack(side="left", fill="both", expand=True)
    vscroll.pack(side="right", fill="y")
    canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def update_count():
        n = sum(v.get() for v in check_vars.values())
        count_label.config(text=f"{n} of {len(all_subs)} selected")

    def build_rows():
        for child in inner.winfo_children():
            child.destroy()
        query = filter_var.get().strip().lower()
        for sub in all_subs:
            if query and query not in sub.display_name.lower() and query not in sub.subscription_id.lower():
                continue
            row = ttk.Frame(inner, padding=(4, 4))
            row.pack(fill="x", anchor="w")
            ttk.Checkbutton(row, variable=check_vars[sub.subscription_id], command=update_count).pack(side="left")
            text_frame = ttk.Frame(row)
            text_frame.pack(side="left", fill="x", expand=True, padx=(4, 0))
            ttk.Label(text_frame, text=sub.display_name, font=("Segoe UI", 10)).pack(anchor="w")
            ttk.Label(text_frame, text=sub.subscription_id, style="Sub.TLabel").pack(anchor="w")

    filter_var.trace_add("write", lambda *_: build_rows())
    build_rows()

    def select_all():
        for v in check_vars.values():
            v.set(True)
        update_count()

    def select_none():
        for v in check_vars.values():
            v.set(False)
        update_count()

    def on_ok():
        result.extend(s for s in all_subs if check_vars[s.subscription_id].get())
        root.destroy()

    footer = ttk.Frame(root, padding=16)
    footer.pack(fill="x")
    ttk.Button(footer, text="Select All", command=select_all).pack(side="left")
    ttk.Button(footer, text="Select None", command=select_none).pack(side="left", padx=(8, 0))
    ttk.Button(footer, text="Cancel (scan ALL)", command=root.destroy).pack(side="right")
    ttk.Button(footer, text="OK", command=on_ok, style="Accent.TButton").pack(side="right", padx=(0, 8))

    try:
        root.eval("tk::PlaceWindow . center")
    except tk.TclError:
        pass
    root.mainloop()
    return result


def select_subscriptions(all_subs: list, requested_ids):
    """Lets the operator scope the run to a subset of subscriptions instead of the whole
    tenant. --subscription-ids skips the prompt entirely (scripted/unattended runs);
    otherwise a graphical tkinter picker is tried first, falling back to a text numbered
    picker if no display is available (e.g. Cloud Shell), both defaulting to ALL when
    nothing is picked."""
    if requested_ids:
        requested = [s.strip() for s in requested_ids.replace(";", ",").split(",") if s.strip()]
        by_id = {s.subscription_id: s for s in all_subs}
        matched = [by_id[i] for i in requested if i in by_id]
        missing = [i for i in requested if i not in by_id]
        if missing:
            print(f"WARNING: these --subscription-ids were not found among your enabled subscriptions "
                  f"and will be skipped: {', '.join(missing)}")
        if not matched:
            raise RuntimeError("None of the requested --subscription-ids matched an enabled subscription.")
        return matched

    gui_result = _gui_subscription_picker(all_subs)
    if gui_result is not None:
        if not gui_result:
            print("No subscriptions selected in the picker - scanning ALL enabled subscriptions.")
            return all_subs
        return gui_result

    if not sys.stdin.isatty():
        return all_subs

    print("\nEnabled subscriptions in this tenant:")
    for i, s in enumerate(all_subs, start=1):
        print(f"  [{i}] {s.display_name} ({s.subscription_id})")
    selection = input("\nEnter subscription numbers to scan (e.g. 1,3,5), or press Enter to scan ALL: ").strip()
    if not selection:
        return all_subs

    indexes = [int(tok.strip()) - 1 for tok in selection.replace(";", ",").split(",") if tok.strip().lstrip("-").isdigit()]
    picked = [all_subs[i] for i in indexes if 0 <= i < len(all_subs)]
    if not picked:
        print("WARNING: no valid selection recognized - scanning ALL enabled subscriptions instead.")
        return all_subs
    return picked


def run_phase_with_validation(name: str, folder: Path, phase_callable, max_attempts: int = 3, require_workbook: bool = False):
    for attempt in range(1, max_attempts + 1):
        try:
            phase_callable()
        except Exception as exc:
            print(f"  {name} error: {exc}")

        xlsx = _newest_xlsx(folder)
        if not xlsx:
            if not require_workbook:
                return
            if attempt < max_attempts:
                print(f"  ! {name} did not produce an Excel workbook (attempt {attempt}/{max_attempts}). Retrying...")
                continue
            raise RuntimeError(f"{name} failed to produce a required Excel workbook in {folder}.")

        if _is_valid_xlsx(xlsx):
            return

        if attempt < max_attempts:
            print(f"  ! {name} workbook looks OneDrive-encrypted (attempt {attempt}/{max_attempts}). Retrying...")
            time.sleep(5)
        else:
            print(f"  ! {name} workbook still invalid after {max_attempts} attempts (OneDrive sync). "
                  f"Dashboard will treat this phase as unavailable.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Azure WAF/CAF Workshop Launcher (Python port)")
    parser.add_argument("--output-dir", default=None, help='Base output directory (default: "AzureWorkshop" next to this script).')
    parser.add_argument("--skip-metrics", action="store_true", help="Skip the metrics phase (slower, queries per-resource).")
    parser.add_argument("--skip-finops", action="store_true", help="Skip the FinOps extended export.")
    parser.add_argument("--keep-previous", action="store_true", help="Archive an existing output folder instead of deleting it.")
    parser.add_argument("--skip-graph-security", action="store_true",
                         help="Skip interactive Microsoft Graph sign-in for Defender XDR incidents/alerts.")
    parser.add_argument("--subscription-ids", default=None,
                         help="Comma-separated subscription GUIDs to scan (default: interactive picker, or every "
                              "enabled subscription when not running in a terminal).")
    args = parser.parse_args()

    output_dir = Path(args.output_dir) if args.output_dir else (REPO_ROOT / "AzureWorkshop")

    if "onedrive" in str(output_dir).lower():
        print("WARNING: OneDrive may automatically encrypt generated workbooks while they are being written. "
              "If dashboard inputs come back unreadable, pause OneDrive sync for this folder during the run "
              "or re-run with --output-dir pointing outside OneDrive.")

    if output_dir.exists():
        if args.keep_previous:
            archived = output_dir.with_name(f"{output_dir.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            print(f"Archiving previous run to {archived}")
            output_dir.rename(archived)
        else:
            print(f"Removing previous run at {output_dir} ...")
            shutil.rmtree(output_dir)

    phase_folders = {
        "discovery": output_dir / "01_Discovery",
        "advisor": output_dir / "02_Advisor",
        "metrics": output_dir / "03_Metrics",
        "governance": output_dir / "04_Governance",
        "security": output_dir / "05_Security",
        "checklists": output_dir / "06_Checklists",
        "dashboard": output_dir / "07_Dashboard",
    }
    for folder in phase_folders.values():
        folder.mkdir(parents=True, exist_ok=True)

    print("Signing in...")
    credential = get_credential()
    all_subs = list_enabled_subscriptions(credential)
    if not all_subs:
        raise RuntimeError("No enabled subscriptions are accessible in the current tenant.")
    subs = select_subscriptions(all_subs, args.subscription_ids)
    subscription_ids = [s.subscription_id for s in subs]

    print("Azure WAF/CAF Workshop - Discovery Launcher (Python)")
    print(f"Scope:   {len(subscription_ids)} enabled subscription(s)")
    print(f"Output:  {output_dir}\n")

    start_time = time.time()

    def with_output_env(folder: Path, func, *args, **kwargs):
        """Phase scripts read AZWORKSHOP_OUTPUT to build their output path, so this
        just points them at the right phase folder without duplicating that logic."""
        import os
        os.environ["AZWORKSHOP_OUTPUT"] = str(folder)
        return func(*args, credential=credential, subscription_ids=subscription_ids, **kwargs)

    print(" PHASE 1/7: Resource Discovery")
    run_phase_with_validation("Discovery", phase_folders["discovery"],
                               lambda: with_output_env(phase_folders["discovery"], phase1_discovery.run))

    print("\n PHASE 2/7: Advisor Recommendations")
    run_phase_with_validation("Advisor", phase_folders["advisor"],
                               lambda: with_output_env(phase_folders["advisor"], phase2_advisor.run))

    if not args.skip_finops:
        print("\n PHASE 2b: FinOps Extended Export")
        run_phase_with_validation("FinOps", phase_folders["advisor"],
                                   lambda: with_output_env(phase_folders["advisor"], phase2b_finops.run))
    else:
        print("\n PHASE 2b: FinOps Extended Export - SKIPPED")

    if not args.skip_metrics:
        print("\n PHASE 3/7: Metrics & Right-Sizing (this takes longer)")
        run_phase_with_validation("Metrics", phase_folders["metrics"],
                                   lambda: with_output_env(phase_folders["metrics"], phase3_metrics.run))
    else:
        print("\n PHASE 3/7: Metrics - SKIPPED")

    print("\n PHASE 4/7: Governance Visualization")
    run_phase_with_validation("Governance", phase_folders["governance"],
                               lambda: with_output_env(phase_folders["governance"], phase4_governance.run))

    print("\n PHASE 5/7: Security Assessment")
    run_phase_with_validation(
        "Security", phase_folders["security"],
        lambda: with_output_env(phase_folders["security"], phase5_security.run,
                                 skip_graph_security=args.skip_graph_security),
        require_workbook=True,
    )

    print("\n PHASE 6/7: Review Checklists (community WAF checks)")
    run_phase_with_validation("Checklists", phase_folders["checklists"],
                               lambda: with_output_env(phase_folders["checklists"], phase6_checklists.run))

    print("\n PHASE 7/7: Generating Consolidated Dashboard")
    try:
        result = subprocess.run(
            [sys.executable, str(REPO_ROOT / "generate-dashboard.py"), str(output_dir)],
            check=True,
        )
        print(f"  Dashboard generated in {output_dir}/07_Dashboard/")
    except subprocess.CalledProcessError as exc:
        print(f"  Dashboard generation failed: {exc}")

    elapsed_minutes = (time.time() - start_time) / 60
    print("\nALL PHASES COMPLETE")
    print(f"Duration: {elapsed_minutes:.1f} minutes")
    print(f"Output:   {output_dir}")
    print("  01_Discovery/    - Resource inventory (Excel)")
    print("  02_Advisor/      - Advisor recommendations (Excel)")
    print("  03_Metrics/      - Right-sizing analysis (Excel)")
    print("  04_Governance/   - Governance data export (Excel)")
    print("  05_Security/     - Security posture and operations (Excel)")
    print("  06_Checklists/   - Azure/review-checklists ARG compliance (Excel)")
    print("  07_Dashboard/    - Consolidated dashboard (HTML)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
