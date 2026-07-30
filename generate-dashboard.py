"""
Azure WAF/CAF Workshop - Consolidated Dashboard Generator

Deterministic Python agent that reads all Excel outputs from the launcher,
analyzes the data, generates action items, and produces a single consolidated
HTML dashboard report.

Usage:
    python3 generate-dashboard.py <output_directory>
    
    Where <output_directory> is the folder created by Launch-AzureWorkshop.ps1
    (e.g., ~/AzureWorkshop_20260730_110000)

Requirements:
    pip install openpyxl  (usually pre-installed in Cloud Shell)
"""

import sys
import os
import json
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    os.system("pip install openpyxl --quiet")
    from openpyxl import load_workbook


# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

class DashboardConfig:
    SEVERITY_COLORS = {
        "critical": "#D13438",
        "high": "#E74856",
        "medium": "#FF8C00",
        "low": "#0078D4",
        "info": "#107C10"
    }
    PILLAR_COLORS = {
        "Reliability": "#0078D4",
        "Security": "#D13438",
        "Cost Optimization": "#107C10",
        "Operational Excellence": "#5C2D91",
        "Performance Efficiency": "#008272"
    }


# ═══════════════════════════════════════════════════════════════════════════════
# DATA READER
# ═══════════════════════════════════════════════════════════════════════════════

class ExcelReader:
    """Reads Excel workbooks and returns structured data."""
    
    @staticmethod
    def read_workbook(path: str) -> dict:
        """Read all sheets from an Excel file into a dict of lists-of-dicts."""
        if not os.path.exists(path):
            return {}
        
        wb = load_workbook(path, read_only=True, data_only=True)
        result = {}
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                result[sheet_name] = []
                continue
            
            headers = [str(h) if h else f"col_{i}" for i, h in enumerate(rows[0])]
            data = []
            for row in rows[1:]:
                record = {}
                for i, val in enumerate(row):
                    if i < len(headers):
                        record[headers[i]] = val if val is not None else ""
                data.append(record)
            result[sheet_name] = data
        
        wb.close()
        return result
    
    @staticmethod
    def read_all(base_dir: str) -> dict:
        """Read all Excel files from the workshop output directory."""
        data = {
            "discovery": {},
            "advisor": {},
            "metrics": {},
            "governance": {}
        }
        
        # Discovery
        path = os.path.join(base_dir, "01_Discovery", "AzureDiscovery.xlsx")
        if os.path.exists(path):
            data["discovery"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Discovery: {sum(len(v) for v in data['discovery'].values())} records across {len(data['discovery'])} sheets")
        
        # Advisor
        path = os.path.join(base_dir, "02_Advisor", "AzureAdvisor.xlsx")
        if os.path.exists(path):
            data["advisor"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Advisor: {sum(len(v) for v in data['advisor'].values())} records across {len(data['advisor'])} sheets")
        
        # Metrics
        path = os.path.join(base_dir, "03_Metrics", "AzureMetrics.xlsx")
        if os.path.exists(path):
            data["metrics"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Metrics: {sum(len(v) for v in data['metrics'].values())} records across {len(data['metrics'])} sheets")
        
        # Governance
        path = os.path.join(base_dir, "04_Governance", "AzureGovernance.xlsx")
        if os.path.exists(path):
            data["governance"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Governance: {sum(len(v) for v in data['governance'].values())} records across {len(data['governance'])} sheets")
        
        return data


# ═══════════════════════════════════════════════════════════════════════════════
# ANALYSIS ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

class InsightEngine:
    """Analyzes collected data and generates insights + action items."""
    
    def __init__(self, data: dict):
        self.data = data
        self.insights = []
        self.action_items = []
        self.scores = {
            "Reliability": 0,
            "Security": 0,
            "Cost Optimization": 0,
            "Operational Excellence": 0,
            "Performance Efficiency": 0
        }
        self.max_score = 100
    
    def analyze(self):
        """Run all analysis passes."""
        self._analyze_reliability()
        self._analyze_security()
        self._analyze_cost()
        self._analyze_operational_excellence()
        self._analyze_performance()
        return self
    
    def _analyze_reliability(self):
        """Assess reliability posture."""
        score = 100
        
        # Check VMs without availability zones
        vms = self.data.get("discovery", {}).get("VMs", [])
        vms_no_zone = [v for v in vms if not v.get("zone") and not v.get("availabilityZone")]
        if vms_no_zone:
            deduction = min(30, len(vms_no_zone) * 3)
            score -= deduction
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "high",
                "title": f"{len(vms_no_zone)} VMs without Availability Zones",
                "description": "These VMs are not protected against datacenter failures. Consider migrating to zone-redundant deployments.",
                "resources": [v.get("name", v.get("Name", "")) for v in vms_no_zone[:10]]
            })
        
        # Check backup vaults
        backups = self.data.get("discovery", {}).get("BackupVaults", [])
        if not backups or (len(backups) == 1 and backups[0].get("Result") == "No resources found"):
            score -= 20
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "critical",
                "title": "No Backup/Recovery Services Vaults found",
                "description": "No backup infrastructure detected. Critical workloads should have backup policies configured.",
                "resources": []
            })
        
        # Check DDoS protection on VNets
        vnets = self.data.get("discovery", {}).get("VNets", []) or self.data.get("governance", {}).get("VNets", [])
        vnets_no_ddos = [v for v in vnets if str(v.get("ddos", v.get("ddosProtection", v.get("enableDdosProtection", "")))).lower() in ("false", "", "none")]
        if vnets_no_ddos:
            score -= 10
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "medium",
                "title": f"{len(vnets_no_ddos)} VNets without DDoS Protection",
                "description": "Enable Azure DDoS Protection Standard for internet-facing workloads.",
                "resources": [v.get("name", v.get("Name", "")) for v in vnets_no_ddos[:10]]
            })
        
        # Advisor HighAvailability recommendations
        ha_recs = self.data.get("advisor", {}).get("Reliability", [])
        high_impact_ha = [r for r in ha_recs if str(r.get("impact", "")).lower() == "high"]
        if high_impact_ha:
            score -= min(20, len(high_impact_ha) * 2)
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "high",
                "title": f"{len(high_impact_ha)} High-Impact Reliability recommendations from Advisor",
                "description": "Azure Advisor has identified critical reliability improvements.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in high_impact_ha[:10]]
            })
        
        self.scores["Reliability"] = max(0, score)
    
    def _analyze_security(self):
        """Assess security posture."""
        score = 100
        
        # Secure score from Defender
        secure_scores = self.data.get("governance", {}).get("SecureScores", [])
        if secure_scores:
            avg_score = sum(float(s.get("pct", s.get("percentage", 0)) or 0) for s in secure_scores) / len(secure_scores)
            if avg_score < 0.5:
                score -= 30
            elif avg_score < 0.7:
                score -= 15
            self.insights.append({
                "pillar": "Security",
                "text": f"Average Defender Secure Score: {avg_score*100:.0f}%"
            })
        
        # Storage accounts with public access
        storage = self.data.get("discovery", {}).get("Storage", [])
        public_storage = [s for s in storage if str(s.get("publicBlob", s.get("allowBlobPublicAccess", ""))).lower() == "true"]
        if public_storage:
            score -= min(25, len(public_storage) * 5)
            self.action_items.append({
                "pillar": "Security",
                "severity": "critical",
                "title": f"{len(public_storage)} Storage accounts allow public blob access",
                "description": "Disable public blob access unless explicitly required. This is a common data exposure risk.",
                "resources": [s.get("name", s.get("Name", "")) for s in public_storage[:10]]
            })
        
        # Key Vaults without purge protection
        keyvaults = self.data.get("discovery", {}).get("KeyVaults", []) or self.data.get("governance", {}).get("KeyVaults", [])
        kv_no_purge = [k for k in keyvaults if str(k.get("purgeProtection", k.get("enablePurgeProtection", ""))).lower() in ("false", "", "none")]
        if kv_no_purge:
            score -= min(15, len(kv_no_purge) * 3)
            self.action_items.append({
                "pillar": "Security",
                "severity": "high",
                "title": f"{len(kv_no_purge)} Key Vaults without Purge Protection",
                "description": "Enable purge protection to prevent permanent deletion of secrets, keys, and certificates.",
                "resources": [k.get("name", k.get("Name", "")) for k in kv_no_purge[:10]]
            })
        
        # Storage without TLS 1.2
        old_tls = [s for s in storage if s.get("tlsVersion", s.get("minimumTlsVersion", "")) and "1.2" not in str(s.get("tlsVersion", s.get("minimumTlsVersion", "")))]
        if old_tls:
            score -= min(15, len(old_tls) * 3)
            self.action_items.append({
                "pillar": "Security",
                "severity": "high",
                "title": f"{len(old_tls)} Storage accounts not enforcing TLS 1.2",
                "description": "Enforce minimum TLS 1.2 for all storage accounts to prevent protocol downgrade attacks.",
                "resources": [s.get("name", s.get("Name", "")) for s in old_tls[:10]]
            })
        
        # Security advisor recommendations
        sec_recs = self.data.get("advisor", {}).get("Security", [])
        if sec_recs:
            score -= min(20, len(sec_recs) * 2)
            self.action_items.append({
                "pillar": "Security",
                "severity": "medium",
                "title": f"{len(sec_recs)} Security recommendations from Advisor",
                "description": "Review Azure Advisor security findings and remediate.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in sec_recs[:10]]
            })
        
        self.scores["Security"] = max(0, score)
    
    def _analyze_cost(self):
        """Assess cost optimization opportunities."""
        score = 100
        
        # Orphaned resources
        orphaned = self.data.get("governance", {}).get("OrphanedResources", [])
        if not orphaned:
            # Try discovery sheets
            disks = self.data.get("discovery", {}).get("UnattachedDisks", [])
            dealloc = self.data.get("discovery", {}).get("DeallocatedVMs", [])
            orphaned = disks + dealloc
        
        if orphaned:
            score -= min(25, len(orphaned) * 2)
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "medium",
                "title": f"{len(orphaned)} Orphaned/unused resources detected",
                "description": "These resources incur costs but are not attached to any workload. Review and delete if unused.",
                "resources": [o.get("Name", o.get("name", "")) for o in orphaned[:10]]
            })
        
        # VMs that are deallocated (still paying for disks)
        dealloc_vms = self.data.get("discovery", {}).get("DeallocatedVMs", [])
        if dealloc_vms and not (len(dealloc_vms) == 1 and dealloc_vms[0].get("Result")):
            score -= min(15, len(dealloc_vms) * 2)
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "low",
                "title": f"{len(dealloc_vms)} Deallocated VMs (still paying for disks/IPs)",
                "description": "Deallocated VMs still incur costs for attached disks and static IPs. Consider deleting if no longer needed.",
                "resources": [v.get("name", v.get("Name", "")) for v in dealloc_vms[:10]]
            })
        
        # Right-sizing from metrics
        vm_metrics = self.data.get("metrics", {}).get("VM_RightSizing", [])
        underutilized = [v for v in vm_metrics if "Idle" in str(v.get("Assessment", "")) or "Underutilized" in str(v.get("Assessment", ""))]
        if underutilized:
            score -= min(20, len(underutilized) * 3)
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "high",
                "title": f"{len(underutilized)} VMs are idle or underutilized",
                "description": "These VMs have <15% average CPU over 30 days. Consider downsizing or deallocating.",
                "resources": [f"{v.get('Name', '')} ({v.get('VMSize', '')} @ {v.get('AvgCPU_Pct', '?')}% CPU)" for v in underutilized[:10]]
            })
        
        # Cost recommendations from Advisor
        cost_recs = self.data.get("advisor", {}).get("Cost", [])
        if cost_recs:
            total_savings = sum(float(r.get("annualSavings", r.get("savingsAmount", 0)) or 0) for r in cost_recs)
            score -= min(20, len(cost_recs) * 2)
            desc = f"Azure Advisor estimates potential savings."
            if total_savings > 0:
                desc = f"Azure Advisor estimates ~${total_savings:,.0f} in potential annual savings."
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "high",
                "title": f"{len(cost_recs)} Cost optimization recommendations",
                "description": desc,
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in cost_recs[:10]]
            })
        
        self.scores["Cost Optimization"] = max(0, score)
    
    def _analyze_operational_excellence(self):
        """Assess operational excellence."""
        score = 100
        
        # Diagnostic settings coverage
        diag = self.data.get("metrics", {}).get("DiagnosticsCoverage", [])
        if diag:
            no_diag = [d for d in diag if str(d.get("HasDiagnostics", d.get("Gap", ""))).lower() in ("false", "no diagnostics configured")]
            if no_diag:
                pct_missing = len(no_diag) / len(diag) * 100
                score -= min(25, int(pct_missing / 4))
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "high",
                    "title": f"{len(no_diag)}/{len(diag)} critical resources missing diagnostic settings ({pct_missing:.0f}%)",
                    "description": "Configure diagnostic settings to send logs to Log Analytics for observability and troubleshooting.",
                    "resources": [d.get("Name", d.get("name", "")) for d in no_diag[:10]]
                })
        
        # Tag coverage
        tag_subs = self.data.get("governance", {}).get("TagUsage_Subscriptions", [])
        if tag_subs:
            no_tags = [s for s in tag_subs if not s.get("tags") or s.get("tagCount", 0) == 0]
            if no_tags:
                score -= min(15, len(no_tags) * 5)
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "medium",
                    "title": f"{len(no_tags)} Subscriptions have no tags",
                    "description": "Implement a tagging strategy (Owner, CostCenter, Environment, Application) for governance and cost allocation.",
                    "resources": [s.get("name", s.get("Name", "")) for s in no_tags[:10]]
                })
        
        # Policy compliance
        compliance = self.data.get("governance", {}).get("PolicyCompliance", [])
        if compliance:
            total_nc = sum(int(c.get("NonCompliantCount", 0) or 0) for c in compliance)
            if total_nc > 100:
                score -= 20
            elif total_nc > 20:
                score -= 10
            if total_nc > 0:
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "medium",
                    "title": f"{total_nc} non-compliant policy evaluations",
                    "description": "Review policy compliance and remediate non-compliant resources or adjust policy assignments.",
                    "resources": [f"{c.get('policyAssignment', c.get('policyAssignmentName', ''))} ({c.get('NonCompliantCount', 0)} violations)" for c in compliance[:10]]
                })
        
        # OpEx Advisor recommendations
        opex_recs = self.data.get("advisor", {}).get("OperationalExcellence", [])
        if opex_recs:
            score -= min(15, len(opex_recs))
        
        self.scores["Operational Excellence"] = max(0, score)
    
    def _analyze_performance(self):
        """Assess performance efficiency."""
        score = 100
        
        # Saturated VMs
        vm_metrics = self.data.get("metrics", {}).get("VM_RightSizing", [])
        saturated = [v for v in vm_metrics if "Saturated" in str(v.get("Assessment", ""))]
        if saturated:
            score -= min(25, len(saturated) * 5)
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "high",
                "title": f"{len(saturated)} VMs are saturated (>80% CPU)",
                "description": "These VMs are consistently at high CPU usage. Consider scaling up or scaling out.",
                "resources": [f"{v.get('Name', '')} ({v.get('VMSize', '')} @ {v.get('AvgCPU_Pct', '?')}% CPU)" for v in saturated[:10]]
            })
        
        # Saturated SQL
        sql_metrics = self.data.get("metrics", {}).get("SQL_RightSizing", [])
        saturated_sql = [s for s in sql_metrics if "Saturated" in str(s.get("Assessment", ""))]
        if saturated_sql:
            score -= min(20, len(saturated_sql) * 5)
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "high",
                "title": f"{len(saturated_sql)} SQL Databases are saturated (>80% DTU/CPU)",
                "description": "These databases are consistently at high utilization. Consider scaling up the tier.",
                "resources": [f"{s.get('Name', '')} ({s.get('SKU', '')} @ {s.get('AvgUsage_Pct', '?')}%)" for s in saturated_sql[:10]]
            })
        
        # Performance advisor recommendations
        perf_recs = self.data.get("advisor", {}).get("Performance", [])
        if perf_recs:
            score -= min(15, len(perf_recs) * 2)
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "medium",
                "title": f"{len(perf_recs)} Performance recommendations from Advisor",
                "description": "Azure Advisor has identified performance improvements.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in perf_recs[:10]]
            })
        
        self.scores["Performance Efficiency"] = max(0, score)
    
    def get_overall_score(self) -> int:
        """Weighted average of all pillar scores."""
        weights = {
            "Reliability": 0.25,
            "Security": 0.25,
            "Cost Optimization": 0.20,
            "Operational Excellence": 0.15,
            "Performance Efficiency": 0.15
        }
        return int(sum(self.scores[p] * weights[p] for p in self.scores))


# ═══════════════════════════════════════════════════════════════════════════════
# HTML DASHBOARD GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

class DashboardGenerator:
    """Generates the final consolidated HTML dashboard."""
    
    def __init__(self, engine: InsightEngine, data: dict):
        self.engine = engine
        self.data = data
    
    def generate(self) -> str:
        overall = self.engine.get_overall_score()
        scores = self.engine.scores
        action_items = sorted(self.engine.action_items, 
                            key=lambda x: {"critical": 0, "high": 1, "medium": 2, "low": 3, "info": 4}.get(x["severity"], 5))
        
        # Count stats
        total_resources = sum(int(r.get("Count", 0) or 0) for r in self.data.get("discovery", {}).get("ResourceSummary", []))
        total_subs = len(self.data.get("discovery", {}).get("Subscriptions", []) or self.data.get("governance", {}).get("Subscriptions", []))
        total_actions = len(action_items)
        critical_actions = len([a for a in action_items if a["severity"] in ("critical", "high")])
        
        # Build action items HTML
        actions_html = ""
        for i, action in enumerate(action_items, 1):
            color = DashboardConfig.SEVERITY_COLORS.get(action["severity"], "#666")
            resources_html = ""
            if action.get("resources"):
                resources_html = "<ul class='resource-list'>" + "".join(f"<li>{r}</li>" for r in action["resources"] if r) + "</ul>"
            
            actions_html += f"""
            <div class="action-item" style="border-left: 4px solid {color};">
                <div class="action-header">
                    <span class="action-num">#{i}</span>
                    <span class="severity-badge" style="background:{color};">{action['severity'].upper()}</span>
                    <span class="pillar-tag">{action['pillar']}</span>
                    <span class="action-title">{action['title']}</span>
                </div>
                <p class="action-desc">{action['description']}</p>
                {resources_html}
            </div>"""
        
        # Build pillar score cards
        pillar_cards = ""
        for pillar, sc in scores.items():
            color = DashboardConfig.PILLAR_COLORS[pillar]
            status_icon = "✅" if sc >= 80 else "⚠️" if sc >= 50 else "🔴"
            pillar_cards += f"""
            <div class="pillar-card">
                <div class="pillar-score" style="color:{color};">{sc}</div>
                <div class="pillar-bar">
                    <div class="pillar-bar-fill" style="width:{sc}%; background:{color};"></div>
                </div>
                <div class="pillar-name">{status_icon} {pillar}</div>
            </div>"""
        
        # Build resource summary table
        resource_summary = self.data.get("discovery", {}).get("ResourceSummary", [])
        top_resources = sorted(resource_summary, key=lambda x: int(x.get("Count", 0) or 0), reverse=True)[:20]
        resource_rows = ""
        for r in top_resources:
            rtype = str(r.get("type", "")).split("/")[-1] if "/" in str(r.get("type", "")) else r.get("type", "")
            resource_rows += f"<tr><td>{r.get('type', '')}</td><td>{r.get('location', '')}</td><td class='num'>{r.get('Count', 0)}</td></tr>"
        
        # Advisor summary
        advisor_summary = self.data.get("advisor", {}).get("SummaryByCategory", [])
        advisor_rows = ""
        for a in advisor_summary:
            advisor_rows += f"<tr><td>{a.get('category', '')}</td><td>{a.get('impact', '')}</td><td class='num'>{a.get('Count', 0)}</td></tr>"
        
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Azure WAF/CAF Workshop - Consolidated Dashboard</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
    font-family: "Segoe UI", -apple-system, sans-serif;
    background: #f5f5f5; color: #333;
    line-height: 1.5;
}}
.dashboard {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}

/* Header */
.header {{
    background: linear-gradient(135deg, #0078D4, #005A9E);
    color: white; padding: 32px; border-radius: 12px;
    margin-bottom: 24px;
}}
.header h1 {{ font-size: 28px; margin-bottom: 8px; }}
.header .subtitle {{ opacity: 0.85; font-size: 14px; }}

/* Overall Score */
.score-section {{
    display: grid; grid-template-columns: 200px 1fr;
    gap: 24px; margin-bottom: 24px;
    background: white; border-radius: 12px; padding: 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}}
.overall-score {{
    display: flex; flex-direction: column; align-items: center; justify-content: center;
}}
.score-circle {{
    width: 140px; height: 140px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 48px; font-weight: 700;
    background: conic-gradient(
        {'#107C10' if overall >= 80 else '#FF8C00' if overall >= 50 else '#D13438'} {overall * 3.6}deg,
        #e0e0e0 {overall * 3.6}deg
    );
    position: relative;
}}
.score-circle::after {{
    content: '{overall}'; position: absolute;
    width: 110px; height: 110px; border-radius: 50%;
    background: white; display: flex; align-items: center; justify-content: center;
    font-size: 42px; font-weight: 700;
    color: {'#107C10' if overall >= 80 else '#FF8C00' if overall >= 50 else '#D13438'};
}}
.score-label {{ margin-top: 8px; font-size: 14px; color: #666; font-weight: 600; }}

.pillars-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;
}}
.pillar-card {{
    background: #f9f9f9; border-radius: 8px; padding: 16px;
    border: 1px solid #eee;
}}
.pillar-score {{ font-size: 32px; font-weight: 700; }}
.pillar-bar {{ height: 6px; background: #e0e0e0; border-radius: 3px; margin: 8px 0; }}
.pillar-bar-fill {{ height: 100%; border-radius: 3px; transition: width 0.5s; }}
.pillar-name {{ font-size: 12px; font-weight: 600; color: #555; }}

/* Stats */
.stats-row {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px; margin-bottom: 24px;
}}
.stat-box {{
    background: white; border-radius: 10px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;
}}
.stat-box .value {{ font-size: 36px; font-weight: 700; color: #0078D4; }}
.stat-box .label {{ font-size: 12px; color: #666; margin-top: 4px; }}

/* Action Items */
.section {{ margin-bottom: 24px; }}
.section-title {{
    font-size: 18px; font-weight: 700; margin-bottom: 16px;
    padding-bottom: 8px; border-bottom: 2px solid #0078D4;
}}
.action-item {{
    background: white; border-radius: 8px; padding: 16px; margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}
.action-header {{ display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }}
.action-num {{ font-weight: 700; color: #888; font-size: 12px; }}
.severity-badge {{
    color: white; padding: 2px 8px; border-radius: 4px;
    font-size: 10px; font-weight: 700; text-transform: uppercase;
}}
.pillar-tag {{
    background: #f0f0f0; padding: 2px 8px; border-radius: 4px;
    font-size: 10px; color: #555;
}}
.action-title {{ font-weight: 600; font-size: 14px; }}
.action-desc {{ color: #555; margin-top: 8px; font-size: 13px; }}
.resource-list {{
    margin-top: 8px; padding-left: 20px;
    font-size: 11px; color: #666; max-height: 120px; overflow-y: auto;
}}
.resource-list li {{ margin: 2px 0; }}

/* Tables */
.data-grid {{
    display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px;
}}
@media (max-width: 900px) {{ .data-grid {{ grid-template-columns: 1fr; }} }}
.table-card {{
    background: white; border-radius: 10px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}}
.table-card h3 {{ font-size: 14px; margin-bottom: 12px; color: #333; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
th {{ background: #f5f5f5; padding: 8px; text-align: left; font-weight: 600; border-bottom: 2px solid #ddd; }}
td {{ padding: 6px 8px; border-bottom: 1px solid #eee; }}
tr:hover td {{ background: #f8fbff; }}
.num {{ text-align: right; font-weight: 600; }}

/* Footer */
.footer {{
    text-align: center; padding: 24px; color: #888; font-size: 11px;
    border-top: 1px solid #ddd; margin-top: 32px;
}}
</style>
</head>
<body>
<div class="dashboard">

<!-- Header -->
<div class="header">
    <h1>☁️ Azure WAF/CAF Workshop - Discovery Report</h1>
    <div class="subtitle">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Consolidated view across all discovery phases</div>
</div>

<!-- Overall Score + Pillar Breakdown -->
<div class="score-section">
    <div class="overall-score">
        <div class="score-circle"></div>
        <div class="score-label">Overall WAF Score</div>
    </div>
    <div class="pillars-grid">
        {pillar_cards}
    </div>
</div>

<!-- Stats Row -->
<div class="stats-row">
    <div class="stat-box"><div class="value">{total_subs}</div><div class="label">Subscriptions</div></div>
    <div class="stat-box"><div class="value">{total_resources:,}</div><div class="label">Total Resources</div></div>
    <div class="stat-box"><div class="value">{total_actions}</div><div class="label">Action Items</div></div>
    <div class="stat-box"><div class="value" style="color:#D13438;">{critical_actions}</div><div class="label">Critical/High Priority</div></div>
</div>

<!-- Action Items -->
<div class="section">
    <div class="section-title">🎯 Action Items ({total_actions})</div>
    {actions_html}
</div>

<!-- Data Summary Tables -->
<div class="data-grid">
    <div class="table-card">
        <h3>📦 Top Resource Types</h3>
        <table>
            <thead><tr><th>Resource Type</th><th>Location</th><th>Count</th></tr></thead>
            <tbody>{resource_rows}</tbody>
        </table>
    </div>
    <div class="table-card">
        <h3>💡 Advisor Summary</h3>
        <table>
            <thead><tr><th>Category</th><th>Impact</th><th>Count</th></tr></thead>
            <tbody>{advisor_rows}</tbody>
        </table>
    </div>
</div>

<!-- Footer -->
<div class="footer">
    Azure WAF/CAF Workshop Discovery Report | Generated by Azure Governance Discovery Toolkit<br>
    This report consolidates findings from Resource Discovery, Azure Advisor, Metrics Analysis, and Governance Visualization.
</div>

</div>
</body>
</html>"""
        
        return html


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 generate-dashboard.py <output_directory>")
        print("  Where <output_directory> is the folder from Launch-AzureWorkshop.ps1")
        sys.exit(1)
    
    base_dir = sys.argv[1]
    if not os.path.isdir(base_dir):
        print(f"❌ Directory not found: {base_dir}")
        sys.exit(1)
    
    print(f"\n{'═'*60}")
    print(f"  Azure WAF/CAF Workshop - Dashboard Generator")
    print(f"{'═'*60}")
    print(f"  Input: {base_dir}\n")
    
    # Read all data
    print("📂 Reading data files...")
    data = ExcelReader.read_all(base_dir)
    
    # Analyze
    print("\n🔍 Analyzing across WAF pillars...")
    engine = InsightEngine(data).analyze()
    
    print(f"\n📊 Scores:")
    for pillar, score in engine.scores.items():
        bar = "█" * (score // 5) + "░" * (20 - score // 5)
        status = "✅" if score >= 80 else "⚠️ " if score >= 50 else "🔴"
        print(f"   {status} {pillar:<25} {bar} {score}/100")
    print(f"\n   Overall WAF Score: {engine.get_overall_score()}/100")
    print(f"   Action Items: {len(engine.action_items)}")
    
    # Generate dashboard
    print("\n🎨 Generating dashboard HTML...")
    generator = DashboardGenerator(engine, data)
    html = generator.generate()
    
    output_path = os.path.join(base_dir, "05_Dashboard", "WAF_Dashboard.html")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    
    print(f"\n{'═'*60}")
    print(f"  ✅ Dashboard generated!")
    print(f"  📁 {output_path}")
    print(f"  💡 Open in any browser to view the interactive report.")
    print(f"{'═'*60}\n")


if __name__ == "__main__":
    main()
