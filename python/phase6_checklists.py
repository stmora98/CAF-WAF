"""Azure Review Checklists (WAF) - Python port of Invoke-AzureChecklists-CloudShell.ps1.

Runs each community checklist item's embedded Azure Resource Graph query against every
enabled subscription, using the checklist JSON files vendored locally in ../checklists
(no network/GitHub access needed at run time - only Azure Resource Graph).

Source: https://github.com/Azure/review-checklists (community-maintained, not an
official Microsoft product; queries are read-only Resource Graph checks).
"""
from __future__ import annotations

import bootstrap
bootstrap.ensure_dependencies()

import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Optional

from common.argquery import run_query
from common.auth import get_credential, resolve_subscription_ids
from common.excelio import add_sheet, make_output_path, new_workbook, save

# Checklist "waf" values (Reliability/Security/Cost/Operations/Performance) don't match
# this toolkit's pillar names 1:1, so translate them here.
PILLAR_MAP = {
    "reliability": "Reliability",
    "security": "Security",
    "cost": "Cost Optimization",
    "operations": "Operational Excellence",
    "performance": "Performance Efficiency",
}

THREAD_LIMIT = 8
BATCH_SIZE = 8  # unique queries combined per Resource Graph call via `union`


def _to_pillar(waf: Optional[str]) -> str:
    key = (waf or "").strip().lower()
    return PILLAR_MAP.get(key, "Operational Excellence")


def _is_non_compliant(value) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return not value
    return str(value).strip().lower() in ("false", "non-compliant", "noncompliant", "no")


def _execute_batch(credential, subscription_ids, queries: List[str]):
    """Runs a batch of KQL queries as one union'd ARG call (cuts call count ~8x vs.
    one call per query, reducing both wall-clock time and 429 throttling). On failure,
    bisects into smaller batches - down to individual per-query calls if needed - so
    one bad/unsupported query doesn't lose the rest of the batch. Returns
    (rows_per_query, failed_count).
    """
    if len(queries) == 1:
        try:
            return [run_query(credential, subscription_ids, queries[0])], 0
        except Exception:
            return [[]], 1

    combined = "union isfuzzy=true " + ", ".join(
        f"({q.strip()} | extend __batchIdx = {i})" for i, q in enumerate(queries)
    )
    try:
        rows = run_query(credential, subscription_ids, combined)
    except Exception:
        mid = len(queries) // 2
        left_rows, left_failed = _execute_batch(credential, subscription_ids, queries[:mid])
        right_rows, right_failed = _execute_batch(credential, subscription_ids, queries[mid:])
        return left_rows + right_rows, left_failed + right_failed

    per_query_rows: List[list] = [[] for _ in queries]
    for row in rows:
        try:
            idx = int(row.get("__batchIdx"))
        except (TypeError, ValueError):
            continue
        if 0 <= idx < len(queries):
            per_query_rows[idx].append(row)
    return per_query_rows, 0


def run(credential=None, subscription_ids=None, services: Optional[List[str]] = None) -> str:
    credential = credential or get_credential()
    subscription_ids = subscription_ids or resolve_subscription_ids(credential)
    output_path = make_output_path("AzureChecklists")
    wb = new_workbook()

    print("\n=== Azure Review Checklists (community WAF checks) ===")

    checklists_dir = Path(__file__).resolve().parent.parent / "checklists"
    if not checklists_dir.exists():
        print(f"  ! Checklist folder not found: {checklists_dir}")
        add_sheet(wb, "Findings", [], empty_message=f"Checklist folder not found: {checklists_dir}")
        save(wb, output_path)
        return output_path

    checklist_files = sorted(checklists_dir.glob("*_checklist.en.json"))
    if services:
        checklist_files = [f for f in checklist_files if f.name.replace("_checklist.en.json", "") in services]
    print(f"  Found {len(checklist_files)} checklists to scan (local, from ../checklists).\n")

    # Flatten every checklist item, then group by the exact ARG query text - many
    # checklists embed identical/shared queries, so deduplicating avoids re-running
    # the same Resource Graph query more than once.
    all_items = []  # (service_name, item_dict)
    for file in checklist_files:
        try:
            checklist = json.loads(file.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"  ! Skipped {file.name}: {exc}")
            continue

        graph_items = [item for item in checklist.get("items", []) if str(item.get("graph", "")).strip()]
        if not graph_items:
            continue

        service_name = checklist.get("metadata", {}).get("name") or file.name.replace("_checklist.en.json", "")
        print(f"  [{service_name}] {len(graph_items)} automated checks")
        for item in graph_items:
            all_items.append((service_name, item))

    by_query = {}
    for service_name, item in all_items:
        by_query.setdefault(item["graph"].strip(), []).append((service_name, item))

    print(f"  {len(all_items)} checks map to {len(by_query)} unique Resource Graph queries.")

    findings = []
    queries_run = 0
    queries_failed = 0

    unique_queries = list(by_query.keys())
    batches = [unique_queries[i:i + BATCH_SIZE] for i in range(0, len(unique_queries), BATCH_SIZE)]

    def execute(batch_queries: List[str]):
        rows_per_query, failed = _execute_batch(credential, subscription_ids, batch_queries)
        return list(zip(batch_queries, rows_per_query)), failed

    with ThreadPoolExecutor(max_workers=THREAD_LIMIT) as pool:
        futures = [pool.submit(execute, batch) for batch in batches]
        for future in as_completed(futures):
            results, failed = future.result()
            queries_failed += failed
            for query, rows in results:
                queries_run += 1
                non_compliant_rows = [r for r in rows if "compliant" in r and _is_non_compliant(r.get("compliant"))]
                if not non_compliant_rows:
                    continue
                for service_name, item in by_query[query]:
                    pillar = _to_pillar(item.get("waf"))
                    for row in non_compliant_rows:
                        findings.append({
                            "Service": service_name,
                            "WafPillar": pillar,
                            "Category": item.get("category"),
                            "Subcategory": item.get("subcategory"),
                            "Severity": item.get("severity"),
                            "Text": item.get("text"),
                            "Link": item.get("link"),
                            "Guid": item.get("guid"),
                            "ResourceId": row.get("id", ""),
                        })

    print(f"\n  Ran {queries_run} checklist queries ({queries_failed} failed/unsupported).")
    print(f"  {len(findings)} non-compliant findings.")

    add_sheet(wb, "Findings", findings, empty_message="No non-compliant resources found")

    by_pillar_severity = {}
    for f in findings:
        key = (f["WafPillar"], f["Severity"])
        by_pillar_severity[key] = by_pillar_severity.get(key, 0) + 1
    summary_pillar = [{"WafPillar": k[0], "Severity": k[1], "Count": v} for k, v in by_pillar_severity.items()]
    add_sheet(wb, "SummaryByPillar", summary_pillar, empty_message="No data")

    by_service = {}
    for f in findings:
        by_service[f["Service"]] = by_service.get(f["Service"], 0) + 1
    summary_service = sorted(
        [{"Service": k, "Count": v} for k, v in by_service.items()],
        key=lambda r: r["Count"], reverse=True,
    )
    add_sheet(wb, "SummaryByService", summary_service, empty_message="No data")

    save(wb, output_path)
    print("\nChecklist scan complete!")
    print(f"File: {output_path}")
    print("Source: https://github.com/Azure/review-checklists (vendored locally in ../checklists)\n")
    return output_path


if __name__ == "__main__":
    run()
