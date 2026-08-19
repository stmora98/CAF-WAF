"""Azure Security Assessment Export - Python port of Invoke-AzureSecurity-CloudShell.ps1.

Collects Defender for Cloud posture through Azure Resource Graph, incidents and alerts
through Microsoft Graph Security, and optional Defender for Endpoint data. Each source
is isolated so a missing license or permission does not stop the rest of the export.
"""
from __future__ import annotations

import bootstrap
bootstrap.ensure_dependencies()

import json
from datetime import datetime, timedelta, timezone

from common.argquery import run_query
from common.auth import get_credential, resolve_subscription_ids
from common.excelio import add_sheet, make_output_path, new_workbook, save
from common.httpretry import get_bearer_token, paged_get, request_with_retry

POSTURE_QUERIES = {
    "SecureScores": """
SecurityResources
| where type == 'microsoft.security/securescores'
| extend percentageScore=todouble(properties.score.percentage), currentScore=todouble(properties.score.current), maxScore=todouble(properties.score.max), weight=todouble(properties.weight)
| project tenantId, subscriptionId, percentageScore, currentScore, maxScore, weight""",
    "ScoreControls": """
SecurityResources
| where type == 'microsoft.security/securescores/securescorecontrols'
| extend controlName=tostring(properties.displayName), controlId=tostring(properties.definition.name), notApplicableResourceCount=toint(properties.notApplicableResourceCount), unhealthyResourceCount=toint(properties.unhealthyResourceCount), healthyResourceCount=toint(properties.healthyResourceCount), percentageScore=todouble(properties.score.percentage), currentScore=todouble(properties.score.current), maxScore=todouble(properties.definition.properties.maxScore), weight=todouble(properties.weight), controlType=tostring(properties.definition.properties.source.sourceType)
| project tenantId, subscriptionId, controlName, controlId, unhealthyResourceCount, healthyResourceCount, notApplicableResourceCount, percentageScore, currentScore, maxScore, weight, controlType
| order by unhealthyResourceCount desc""",
    "Recommendations": """
SecurityResources
| where type == 'microsoft.security/assessments'
| extend recommendationId=name, recommendationName=tostring(properties.displayName), recommendationState=tostring(properties.status.code), recommendationSeverity=tostring(properties.metadata.severity), description=tostring(properties.metadata.description), remediationDescription=tostring(properties.metadata.remediationDescription), assessmentType=tostring(properties.metadata.assessmentType), policyDefinitionId=tostring(properties.metadata.policyDefinitionId), implementationEffort=tostring(properties.metadata.implementationEffort), userImpact=tostring(properties.metadata.userImpact), category=tostring(properties.metadata.categories), threats=tostring(properties.metadata.threats), source=tostring(properties.resourceDetails.Source), affectedResourceId=tostring(properties.resourceDetails.Id), portalLink=tostring(properties.links.azurePortal)
| project tenantId, subscriptionId, recommendationId, recommendationName, recommendationState, recommendationSeverity, affectedResourceId, description, remediationDescription, assessmentType, policyDefinitionId, implementationEffort, userImpact, category, threats, source, portalLink
| order by recommendationSeverity asc, recommendationState desc""",
    "RegulatoryStandards": """
SecurityResources
| where type == 'microsoft.security/regulatorycompliancestandards'
| extend complianceStandard=name, state=tostring(properties.state), passedControls=toint(properties.passedControls), failedControls=toint(properties.failedControls), skippedControls=toint(properties.skippedControls), unsupportedControls=toint(properties.unsupportedControls)
| project tenantId, subscriptionId, complianceStandard, state, passedControls, failedControls, skippedControls, unsupportedControls""",
    "MCSBCompliance": """
SecurityResources
| where type == 'microsoft.security/regulatorycompliancestandards/regulatorycompliancecontrols/regulatorycomplianceassessments'
| extend complianceStandard=extract(@'(?i)/regulatorycompliancestandards/([^/]+)', 1, id), complianceControl=extract(@'(?i)/regulatorycompliancecontrols/([^/]+)', 1, id), assessmentName=tostring(properties.description), state=tostring(properties.state), skippedResources=toint(properties.skippedResources), passedResources=toint(properties.passedResources), failedResources=toint(properties.failedResources)
| where complianceStandard contains 'Azure-Security-Benchmark' or complianceStandard contains 'Microsoft-cloud-security-benchmark' or complianceStandard contains 'MCSB'
| project tenantId, subscriptionId, complianceStandard, complianceControl, assessmentName, state, skippedResources, passedResources, failedResources, id
| order by failedResources desc""",
    "DefenderPlans": """
SecurityResources
| where type == 'microsoft.security/pricings'
| extend planName=name, pricingTier=tostring(properties.pricingTier), subPlan=tostring(properties.subPlan), extensions=tostring(properties.extensions)
| project tenantId, subscriptionId, planName, pricingTier, subPlan, extensions
| order by subscriptionId asc, planName asc""",
    "CloudAlerts": """
SecurityResources
| where type =~ 'microsoft.security/locations/alerts'
| extend alertName=tostring(properties.AlertDisplayName), alertType=tostring(properties.AlertType), systemAlertId=tostring(properties.SystemAlertId), status=tostring(properties.Status), severity=tostring(properties.Severity), description=tostring(properties.Description), remediationSteps=tostring(properties.RemediationSteps), detectedTime=todatetime(properties.DetectedTimeUtc), compromisedEntity=tostring(properties.CompromisedEntity), portalLink=tostring(properties.AlertUri)
| project tenantId, subscriptionId, alertName, alertType, systemAlertId, status, severity, description, remediationSteps, detectedTime, compromisedEntity, portalLink
| order by detectedTime desc""",
}


def _cell_value(value):
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (list, tuple)):
        return "; ".join(str(v) for v in value)
    return json.dumps(value, default=str)


class SecurityExport:
    def __init__(self, credential, subscription_ids, lookback_days: int):
        self.credential = credential
        self.subscription_ids = subscription_ids
        self.lookback_days = lookback_days
        self.wb = new_workbook()
        self.source_status = []

    def add_status(self, source, status, records, details, required_access):
        self.source_status.append({
            "Source": source, "Status": status, "Records": records, "Details": details,
            "RequiredAccess": required_access,
            "CollectedAtUtc": datetime.now(timezone.utc).isoformat(),
        })

    def write_sheet(self, name, rows, empty_message="No data returned by this source."):
        count = add_sheet(self.wb, name, rows, empty_message=empty_message)
        print(f"  [{name}] {count} data rows")
        return count

    def skip(self, sources_sheets, message="Collection skipped by parameter."):
        for source, sheet in sources_sheets:
            self.add_status(source, "Skipped", 0, message, "")
            self.write_sheet(sheet, [], empty_message=message)

    # ─── Defender for Cloud posture (Resource Graph) ──────────────────────
    def collect_posture(self):
        total_records = 0
        failed = []
        for name, query in POSTURE_QUERIES.items():
            try:
                rows = run_query(self.credential, self.subscription_ids, query)
                total_records += len(rows)
                self.write_sheet(name, rows)
            except Exception as exc:
                failed.append(name)
                self.write_sheet(name, [], empty_message=f"Collection failed: {exc}")
                print(f"WARNING: {name} collection failed: {exc}")

        if not failed:
            status, details = "Available", "Secure score, controls, assessments, MCSB, plans, and cloud alerts collected."
        elif len(failed) == len(POSTURE_QUERIES):
            status, details = "Unavailable", f"Failed datasets: {', '.join(failed)}."
        else:
            status, details = "Partial", f"Failed datasets: {', '.join(failed)}."
        self.add_status("Defender for Cloud / Azure Resource Graph", status, total_records, details,
                         "Azure Reader and Microsoft.Security read access at the assessed scopes.")

    # ─── Microsoft Graph Security incidents/alerts ────────────────────────
    def collect_graph_security(self):
        try:
            token = get_bearer_token(self.credential, "https://graph.microsoft.com/.default")
        except Exception as exc:
            for source, sheet in (("Microsoft Graph Security incidents", "Incidents"),
                                   ("Microsoft Graph Security alerts", "Alerts")):
                self.add_status(source, "Error", 0, str(exc),
                                 "Delegated SecurityIncident.Read.All or SecurityAlert.Read.All plus a "
                                 "supported Entra security role.")
                self.write_sheet(sheet, [], empty_message="Microsoft Graph Security authentication was unavailable.")
            print(f"WARNING: Microsoft Graph Security authentication failed: {exc}")
            return

        headers = {"Authorization": f"Bearer {token}"}
        since = (datetime.now(timezone.utc) - timedelta(days=self.lookback_days)).strftime("%Y-%m-%dT%H:%M:%SZ")

        incident_url = (f"https://graph.microsoft.com/v1.0/security/incidents?$top=100"
                         f"&$filter=lastUpdateDateTime ge {since}")
        try:
            raw = paged_get(incident_url, headers)
            incidents = [{
                "IncidentId": i.get("id"), "DisplayName": i.get("displayName"), "Status": i.get("status"),
                "Severity": i.get("severity"), "Classification": i.get("classification"),
                "Determination": i.get("determination"), "AssignedTo": i.get("assignedTo"),
                "CreatedDateTime": i.get("createdDateTime"), "LastUpdateDateTime": i.get("lastUpdateDateTime"),
                "PriorityScore": i.get("priorityScore"), "Description": i.get("description"),
                "Summary": i.get("summary"), "Tags": _cell_value(i.get("customTags")),
                "IncidentWebUrl": i.get("incidentWebUrl"),
            } for i in raw]
            self.write_sheet("Incidents", incidents)
            self.add_status("Microsoft Graph Security incidents", "Available" if incidents else "NoData",
                             len(incidents), f"Incidents updated in the last {self.lookback_days} days.",
                             "SecurityIncident.Read.All and Security Reader, Global Reader, Security Operator, "
                             "or Security Administrator.")
        except Exception as exc:
            self.add_status("Microsoft Graph Security incidents", "Error", 0, str(exc),
                             "SecurityIncident.Read.All and a supported Entra security role.")
            self.write_sheet("Incidents", [], empty_message=f"Incident collection failed: {exc}")
            print(f"WARNING: Incident collection failed: {exc}")

        alert_url = (f"https://graph.microsoft.com/v1.0/security/alerts_v2?$top=100"
                     f"&$filter=lastUpdateDateTime ge {since}")
        try:
            raw = paged_get(alert_url, headers)
            alerts = [{
                "AlertId": a.get("id"), "IncidentId": a.get("incidentId"), "Title": a.get("title"),
                "Status": a.get("status"), "Severity": a.get("severity"), "Classification": a.get("classification"),
                "Determination": a.get("determination"), "ServiceSource": a.get("serviceSource"),
                "DetectionSource": a.get("detectionSource"), "Category": a.get("category"),
                "AssignedTo": a.get("assignedTo"), "CreatedDateTime": a.get("createdDateTime"),
                "LastUpdateDateTime": a.get("lastUpdateDateTime"),
                "MitreTechniques": _cell_value(a.get("mitreTechniques")), "Description": a.get("description"),
                "RecommendedActions": a.get("recommendedActions"), "AlertWebUrl": a.get("alertWebUrl"),
                "IncidentWebUrl": a.get("incidentWebUrl"),
            } for a in raw]
            self.write_sheet("Alerts", alerts)
            self.add_status("Microsoft Graph Security alerts", "Available" if alerts else "NoData",
                             len(alerts), f"Alerts updated in the last {self.lookback_days} days.",
                             "SecurityAlert.Read.All and Security Reader, Global Reader, Security Operator, "
                             "or Security Administrator.")
        except Exception as exc:
            self.add_status("Microsoft Graph Security alerts", "Error", 0, str(exc),
                             "SecurityAlert.Read.All and a supported Entra security role.")
            self.write_sheet("Alerts", [], empty_message=f"Alert collection failed: {exc}")
            print(f"WARNING: Alert collection failed: {exc}")

    # ─── Defender for Endpoint ─────────────────────────────────────────────
    def collect_defender_endpoint(self, max_records: int):
        try:
            token = get_bearer_token(self.credential, "https://api.securitycenter.microsoft.com/.default")
        except Exception as exc:
            for source, sheet in (("Defender for Endpoint machines", "Machines"),
                                   ("Defender for Endpoint recommendations", "EndpointRecommendations"),
                                   ("Defender for Endpoint vulnerabilities", "Vulnerabilities")):
                self.add_status(source, "Unavailable", 0, str(exc),
                                 "A token for api.securitycenter.microsoft.com with the corresponding "
                                 "delegated or application permission.")
                self.write_sheet(sheet, [], empty_message="Defender for Endpoint authentication was unavailable.")
            print(f"WARNING: Defender for Endpoint token acquisition failed: {exc}")
            return

        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        collections = [
            ("Defender for Endpoint machines", "Machines",
             "https://api.securitycenter.microsoft.com/api/machines", "Machine.Read.All",
             lambda r: {"MachineId": r.get("id"), "ComputerDnsName": r.get("computerDnsName"),
                        "OsPlatform": r.get("osPlatform"), "OsVersion": r.get("osVersion"),
                        "HealthStatus": r.get("healthStatus"), "RiskScore": r.get("riskScore"),
                        "ExposureLevel": r.get("exposureLevel"), "OnboardingStatus": r.get("onboardingStatus"),
                        "LastSeen": r.get("lastSeen"), "AadDeviceId": r.get("aadDeviceId"),
                        "RbacGroupName": r.get("rbacGroupName")}),
            ("Defender for Endpoint recommendations", "EndpointRecommendations",
             "https://api.securitycenter.microsoft.com/api/recommendations", "SecurityRecommendation.Read.All",
             lambda r: {"RecommendationId": r.get("id"), "ProductName": r.get("productName"),
                        "RecommendationName": r.get("recommendationName"), "Weaknesses": _cell_value(r.get("weaknesses")),
                        "Vendor": r.get("vendor"), "RecommendedVersion": r.get("recommendedVersion"),
                        "SeverityScore": r.get("severityScore"), "PublicExploit": r.get("publicExploit"),
                        "ActiveAlert": r.get("activeAlert"), "AssociatedThreats": _cell_value(r.get("associatedThreats")),
                        "ExposedMachinesCount": r.get("exposedMachinesCount"), "RemediationType": r.get("remediationType"),
                        "Status": r.get("status")}),
            ("Defender for Endpoint vulnerabilities", "Vulnerabilities",
             "https://api.securitycenter.microsoft.com/api/vulnerabilities", "Vulnerability.Read.All",
             lambda r: {"VulnerabilityId": r.get("id"), "Name": r.get("name"), "Description": r.get("description"),
                        "Severity": r.get("severity"), "CvssV3": r.get("cvssV3"), "ExposedMachines": r.get("exposedMachines"),
                        "PublishedOn": r.get("publishedOn"), "UpdatedOn": r.get("updatedOn"),
                        "PublicExploit": r.get("publicExploit"), "ExploitVerified": r.get("exploitVerified"),
                        "ExploitInKit": r.get("exploitInKit"), "ExploitTypes": _cell_value(r.get("exploitTypes"))}),
        ]

        for source, sheet, url, permission, mapper in collections:
            try:
                raw_rows = []
                next_link = url
                truncated = False
                while next_link:
                    resp = request_with_retry("GET", next_link, headers=headers, timeout=120)
                    payload = resp.json()
                    for row in payload.get("value", []):
                        if len(raw_rows) >= max_records:
                            truncated = True
                            break
                        raw_rows.append(row)
                    next_link = payload.get("@odata.nextLink") if len(raw_rows) < max_records else None
                    if len(raw_rows) >= max_records and next_link:
                        truncated = True
                        break
                mapped = [mapper(r) for r in raw_rows]
                self.write_sheet(sheet, mapped)
                status = "Partial" if truncated else ("Available" if mapped else "NoData")
                details = f"Collection limited to {max_records} records." if truncated else "Collection completed."
                self.add_status(source, status, len(mapped), details, permission)
            except Exception as exc:
                self.add_status(source, "Error", 0, str(exc), permission)
                self.write_sheet(sheet, [], empty_message=f"{source} collection failed: {exc}")
                print(f"WARNING: {source} collection failed: {exc}")

    # ─── Entra identity risk (app credential expiry + guest users) ────────
    def collect_identity_risk(self):
        try:
            token = get_bearer_token(self.credential, "https://graph.microsoft.com/.default")
        except Exception as exc:
            for source, sheet in (("Microsoft Entra app credential expiry", "AppCredentialExpiry"),
                                   ("Microsoft Entra guest users", "GuestUsers")):
                self.add_status(source, "Error", 0, str(exc),
                                 "Delegated Application.Read.All or User.Read.All plus a supported Entra "
                                 "directory role.")
                self.write_sheet(sheet, [], empty_message="Microsoft Graph authentication was unavailable.")
            print(f"WARNING: Microsoft Graph identity authentication failed: {exc}")
            return

        headers = {"Authorization": f"Bearer {token}"}
        now = datetime.now(timezone.utc)

        app_url = ("https://graph.microsoft.com/v1.0/applications?"
                   "$select=id,appId,displayName,passwordCredentials,keyCredentials&$top=999")
        try:
            apps = paged_get(app_url, headers)
            credentials = []
            for app in apps:
                for cred_type, creds in (("Secret", app.get("passwordCredentials") or []),
                                          ("Certificate", app.get("keyCredentials") or [])):
                    for cred in creds:
                        end = cred.get("endDateTime")
                        if not end:
                            continue
                        end_dt = datetime.fromisoformat(end.replace("Z", "+00:00"))
                        days_left = round((end_dt - now).total_seconds() / 86400)
                        if days_left > 90:
                            continue
                        credentials.append({
                            "AppDisplayName": app.get("displayName"), "AppId": app.get("appId"),
                            "CredentialType": cred_type, "StartDateTime": cred.get("startDateTime"),
                            "EndDateTime": end, "DaysUntilExpiry": days_left,
                            "Status": "Expired" if days_left < 0 else "ExpiringSoon",
                        })
            credentials.sort(key=lambda c: c["DaysUntilExpiry"])
            self.write_sheet("AppCredentialExpiry", credentials)
            self.add_status("Microsoft Entra app credential expiry", "Available" if credentials else "NoData",
                             len(credentials), "App registration secrets/certificates expiring within 90 days "
                             "or already expired.",
                             "Application.Read.All and Cloud Application Administrator, Application "
                             "Administrator, or Global Reader.")
        except Exception as exc:
            self.add_status("Microsoft Entra app credential expiry", "Error", 0, str(exc),
                             "Application.Read.All and a supported Entra directory role.")
            self.write_sheet("AppCredentialExpiry", [], empty_message=f"App credential collection failed: {exc}")
            print(f"WARNING: App credential collection failed: {exc}")

        guest_url = ("https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Guest'"
                     "&$select=id,displayName,mail,createdDateTime,accountEnabled&$top=999")
        try:
            raw_guests = paged_get(guest_url, headers)
            guests = [{
                "DisplayName": g.get("displayName"), "Mail": g.get("mail"),
                "CreatedDateTime": g.get("createdDateTime"), "AccountEnabled": g.get("accountEnabled"),
            } for g in raw_guests]
            self.write_sheet("GuestUsers", guests)
            self.add_status("Microsoft Entra guest users", "Available" if guests else "NoData", len(guests),
                             "Guest (B2B) accounts in the tenant.",
                             "User.Read.All and Global Reader or Directory Reader.")
        except Exception as exc:
            self.add_status("Microsoft Entra guest users", "Error", 0, str(exc),
                             "User.Read.All and a supported Entra directory role.")
            self.write_sheet("GuestUsers", [], empty_message=f"Guest user collection failed: {exc}")
            print(f"WARNING: Guest user collection failed: {exc}")


def run(
    credential=None,
    subscription_ids=None,
    lookback_days: int = 30,
    skip_defender_portal: bool = False,
    skip_graph_security: bool = False,
    max_endpoint_records: int = 5000,
) -> str:
    credential = credential or get_credential()
    subscription_ids = subscription_ids or resolve_subscription_ids(credential)
    output_path = make_output_path("AzureSecurity")

    print("\n=== Azure Security Assessment Export ===")
    print(f"Output: {output_path}")
    print(f"Lookback: {lookback_days} days")

    export = SecurityExport(credential, subscription_ids, lookback_days)
    export.collect_posture()

    if skip_defender_portal:
        export.skip([
            ("Microsoft Graph Security incidents", "Incidents"),
            ("Microsoft Graph Security alerts", "Alerts"),
            ("Microsoft Entra app credential expiry", "AppCredentialExpiry"),
            ("Microsoft Entra guest users", "GuestUsers"),
            ("Defender for Endpoint machines", "Machines"),
            ("Defender for Endpoint recommendations", "EndpointRecommendations"),
            ("Defender for Endpoint vulnerabilities", "Vulnerabilities"),
        ])
    else:
        if skip_graph_security:
            export.skip([
                ("Microsoft Graph Security incidents", "Incidents"),
                ("Microsoft Graph Security alerts", "Alerts"),
                ("Microsoft Entra app credential expiry", "AppCredentialExpiry"),
                ("Microsoft Entra guest users", "GuestUsers"),
            ], "Microsoft Graph Security collection skipped by parameter.")
        else:
            export.collect_graph_security()
            export.collect_identity_risk()
        export.collect_defender_endpoint(max_endpoint_records)

    export.write_sheet("SourceStatus", export.source_status)
    save(export.wb, output_path)

    print("\nSecurity export complete.")
    print(f"File: {output_path}")
    print("Review SourceStatus for permissions, licensing, or source availability gaps.")
    return output_path


if __name__ == "__main__":
    run()
