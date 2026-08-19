"""Azure Advisor Full Export - Python port of Invoke-AzureAdvisor-CloudShell.ps1.

Exports all Advisor recommendations to Excel grouped by pillar, plus reservation and
savings-plan (rate optimization) recommendations from the Consumption REST API.
"""
from __future__ import annotations

import bootstrap
bootstrap.ensure_dependencies()

from common.argquery import run_query
from common.auth import get_credential, list_enabled_subscriptions, resolve_subscription_ids
from common.excelio import add_sheet, make_output_path, new_workbook, save
from common.httpretry import get_bearer_token, request_with_retry


def run(credential=None, subscription_ids=None) -> str:
    credential = credential or get_credential()
    subscription_ids = subscription_ids or resolve_subscription_ids(credential)
    output_path = make_output_path("AzureAdvisor")
    wb = new_workbook()

    print("\n=== Azure Advisor Export ===")

    def q(sheet: str, query: str):
        rows = run_query(credential, subscription_ids, query)
        count = add_sheet(wb, sheet, rows, empty_message="No recommendations")
        print(f"  [{sheet}] {count} rows")
        return rows

    q("AllRecommendations", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category = tostring(properties.category),
         impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project category, impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by category asc, impact desc""")

    q("Reliability", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "HighAvailability"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc""")

    q("Security", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "Security"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc""")

    q("Cost", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "Cost"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution),
         savingsAmount = tostring(properties.extendedProperties.savingsAmount),
         savingsCurrency = tostring(properties.extendedProperties.savingsCurrency),
         annualSavings = tostring(properties.extendedProperties.annualSavingsAmount)
| project impact, impactedType, impactedResource, problem, solution,
          savingsAmount, savingsCurrency, annualSavings,
          resourceGroup, subscriptionId
| order by impact desc""")

    q("OperationalExcellence", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "OperationalExcellence"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc""")

    q("Performance", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| where properties.category == "Performance"
| extend impact = tostring(properties.impact),
         impactedType = tostring(properties.impactedField),
         impactedResource = tostring(properties.impactedValue),
         problem = tostring(properties.shortDescription.problem),
         solution = tostring(properties.shortDescription.solution)
| project impact, impactedType, impactedResource, problem, solution,
          resourceGroup, subscriptionId
| order by impact desc""")

    q("SummaryByCategory", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category = tostring(properties.category),
         impact = tostring(properties.impact)
| summarize Count=count() by category, impact
| order by category asc, impact desc""")

    q("SummaryByResource", """
advisorresources
| where type == "microsoft.advisor/recommendations"
| extend category = tostring(properties.category),
         impactedType = tostring(properties.impactedField)
| summarize Count=count() by impactedType, category
| order by Count desc""")

    # Reservation / Savings Plan recommendations (rate optimization) - subscription-scoped
    # REST call, not Resource Graph, since these come from Microsoft.Consumption, not ARG.
    print("\nChecking reservation & savings plan recommendations...")
    reservation_rows = []
    try:
        access_token = get_bearer_token(credential)
        headers = {"Authorization": f"Bearer {access_token}"}
        subs = [s for s in list_enabled_subscriptions(credential) if s.subscription_id in subscription_ids]
        for sub in subs:
            url = (f"https://management.azure.com/subscriptions/{sub.subscription_id}"
                   f"/providers/Microsoft.Consumption/reservationRecommendations?api-version=2021-10-01")
            try:
                resp = request_with_retry("GET", url, headers=headers, timeout=60)
                for item in resp.json().get("value", []):
                    p = item.get("properties", {})
                    reservation_rows.append({
                        "Subscription": sub.display_name,
                        "SubscriptionId": sub.subscription_id,
                        "ResourceType": p.get("resourceType"),
                        "SkuName": p.get("skuName"),
                        "Location": p.get("location"),
                        "Term": p.get("term"),
                        "LookBackPeriod": p.get("lookBackPeriod"),
                        "RecommendedQty": p.get("recommendedQuantity"),
                        "CostWithNoRI": p.get("costWithNoReservedInstances"),
                        "CostWithRI": p.get("totalCostWithReservedInstances"),
                        "NetSavings": p.get("netSavings"),
                        "Currency": p.get("currency"),
                        "Scope": p.get("scope"),
                    })
            except Exception as exc:
                status = getattr(getattr(exc, "response", None), "status_code", 0)
                if status in (401, 403):
                    print(f"  ! No permission to read reservation recommendations for subscription "
                          f"{sub.display_name} (needs Cost Management Reader).")
                else:
                    print(f"  ! Reservation recommendations unavailable for subscription {sub.display_name}: {exc}")
    except Exception as exc:
        print(f"  ! Could not check reservation recommendations: {exc}")

    count = add_sheet(
        wb, "ReservationRecommendations", reservation_rows,
        empty_message="No reservation/savings plan recommendations found (or permission "
                       "unavailable - requires Cost Management Reader)",
    )
    print(f"  [ReservationRecommendations] {count} rows")

    save(wb, output_path)
    print("\nAdvisor export complete!")
    print(f"File: {output_path}\n")
    return output_path


if __name__ == "__main__":
    run()
