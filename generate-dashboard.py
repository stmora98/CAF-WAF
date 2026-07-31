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
import re
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Optional
from zipfile import BadZipFile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

try:
    from openpyxl import load_workbook
    from openpyxl.utils.exceptions import InvalidFileException
except ImportError:
    os.system("pip install openpyxl --quiet")
    from openpyxl import load_workbook
    from openpyxl.utils.exceptions import InvalidFileException


# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

class DashboardConfig:
    SEVERITY_COLORS = {
     "critical": "var(--cp-danger)",
     "high": "var(--cp-danger)",
     "medium": "var(--cp-warning)",
     "low": "var(--cp-link)",
     "info": "var(--cp-success)"
    }
    PILLAR_COLORS = {
     "Reliability": "var(--cp-link)",
     "Security": "var(--cp-danger)",
     "Cost Optimization": "var(--cp-success)",
     "Operational Excellence": "var(--cp-accent)",
     "Performance Efficiency": "var(--cp-warning)"
    }


# ═══════════════════════════════════════════════════════════════════════════════
# BILINGUAL (EN/ES) SUPPORT
# Both language variants are always rendered; CSS toggles visibility via the
# .i18n-en/.i18n-es classes based on a "lang-es" class on <body> (see the
# embedded <script> in DashboardGenerator.generate()).
# ═══════════════════════════════════════════════════════════════════════════════

def bi(en: str, es: str) -> str:
    """Wrap a piece of text in both language variants for the client-side toggle."""
    return f'<span class="i18n-en">{en}</span><span class="i18n-es">{es}</span>'


PILLAR_ES = {
    "Reliability": "Confiabilidad",
    "Security": "Seguridad",
    "Cost Optimization": "Optimización de Costos",
    "Operational Excellence": "Excelencia Operativa",
    "Performance Efficiency": "Eficiencia de Rendimiento",
}

SEVERITY_ES = {
    "critical": "CRÍTICO",
    "high": "ALTO",
    "medium": "MEDIO",
    "low": "BAJO",
    "info": "INFO",
}

SUBCATEGORY_ES = {
    "Zone Resiliency": "Resiliencia de Zona",
    "Regional Resiliency": "Resiliencia Regional",
    "Data Protection and Recovery": "Protección y Recuperación de Datos",
    "Governance and Compliance": "Gobernanza y Cumplimiento",
    "Scalability": "Escalabilidad",
    "Monitoring and Alerting": "Monitoreo y Alertas",
    "Service Upgrade and Retirement": "Actualización y Retiro de Servicios",
    "Other": "Otro",
    "Compute Optimization": "Optimización de Cómputo",
    "Storage Optimization": "Optimización de Almacenamiento",
    "Network Optimization": "Optimización de Red",
    "Data Performance": "Rendimiento de Datos",
    "Efficiency Optimization": "Optimización de Eficiencia",
    "Failure Mitigation": "Mitigación de Fallas",
    "Safe and Secure Deployment": "Implementación Segura",
}

COST_SOURCE_ES = {
    "Orphaned/unused resources": "Recursos huérfanos/no utilizados",
    "Deallocated VMs": "VMs desasignadas",
    "Idle/underutilized VMs": "VMs inactivas/subutilizadas",
    "Advisor Cost recommendations": "Recomendaciones de costos de Advisor",
}


# ═══════════════════════════════════════════════════════════════════════════════
# DATA READER
# ═══════════════════════════════════════════════════════════════════════════════

class ExcelReader:
    """Reads Excel workbooks and returns structured data."""

    @staticmethod
    def find_workbook(base_dir: str, folder: str, prefix: str) -> Optional[str]:
        """Return the newest workbook matching a phase's output prefix."""
        phase_dir = Path(base_dir) / folder
        candidates = list(phase_dir.glob(f"{prefix}*.xlsx"))
        if not candidates:
            print(f"  ! {prefix}: no workbook found in {phase_dir}")
            return None
        return str(max(candidates, key=lambda path: path.stat().st_mtime))
    
    @staticmethod
    def read_workbook(path: str) -> dict:
        """Read all sheets from an Excel file into a dict of lists-of-dicts."""
        if not os.path.exists(path):
            return {}
        
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
        except (BadZipFile, InvalidFileException, OSError) as exc:
            print(f"  ! Skipping invalid workbook {path}: {exc}")
            return {}
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
            "governance": {},
            "security": {}
        }
        
        # Discovery
        path = ExcelReader.find_workbook(base_dir, "01_Discovery", "AzureDiscovery")
        if path:
            data["discovery"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Discovery: {sum(len(v) for v in data['discovery'].values())} records across {len(data['discovery'])} sheets")
        
        # Advisor
        path = ExcelReader.find_workbook(base_dir, "02_Advisor", "AzureAdvisor")
        if path:
            data["advisor"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Advisor: {sum(len(v) for v in data['advisor'].values())} records across {len(data['advisor'])} sheets")
        
        # Metrics
        path = ExcelReader.find_workbook(base_dir, "03_Metrics", "AzureMetrics")
        if path:
            data["metrics"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Metrics: {sum(len(v) for v in data['metrics'].values())} records across {len(data['metrics'])} sheets")
        
        # Governance
        path = ExcelReader.find_workbook(base_dir, "04_Governance", "AzureGovernance")
        if path:
            data["governance"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Governance: {sum(len(v) for v in data['governance'].values())} records across {len(data['governance'])} sheets")

        # Security
        path = ExcelReader.find_workbook(base_dir, "05_Security", "AzureSecurity")
        if path:
            data["security"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Security: {sum(len(v) for v in data['security'].values())} records across {len(data['security'])} sheets")
        
        return data


# ═══════════════════════════════════════════════════════════════════════════════
# ADVISOR SCORE MODEL
# ═══════════════════════════════════════════════════════════════════════════════

class AdvisorScoreModel:
    """
    Implements the official Azure Advisor Score methodology using the data this
    toolkit collects.
    Reference: https://learn.microsoft.com/en-us/azure/advisor/advisor-score#calculation-of-advisor-score

    Official formula (Reliability / Performance / Operational Excellence):
        Subcategory Score = (Healthy Resources / Total Applicable Resources) * 100
        Category Score = sum((Healthy/Total) * SubcategoryWeight) / sum(SubcategoryWeight) * 100

    Deviations from the live Advisor score (documented, not hidden):
      - Subcategory weights below are copied verbatim from the Microsoft Learn doc.
      - Recommendations are mapped to subcategories via keyword matching on the
        recommendation problem/solution text, because Azure Resource Graph does not
        expose Advisor's internal subcategory tag.
      - "Total Applicable Resources" uses the total discovered resource inventory
        (01_Discovery/ResourceSummary) as a shared pool, because Advisor's exact
        per-subcategory assessed-resource counts aren't exposed outside the portal.
      - Security score uses the Microsoft Defender Secure Score directly (the same
        model Advisor uses for this category), read from 04_Governance/SecureScores.
      - Cost score uses a resource-count healthy ratio instead of retail-cost
        weighting, because this toolkit doesn't call the Azure Retail Prices API.
    """

    RELIABILITY_WEIGHTS = {
        "Zone Resiliency": 30,
        "Regional Resiliency": 25,
        "Data Protection and Recovery": 20,
        "Governance and Compliance": 10,
        "Scalability": 10,
        "Monitoring and Alerting": 5,
        "Service Upgrade and Retirement": 5,
        "Other": 5,
    }
    PERFORMANCE_WEIGHTS = {
        "Compute Optimization": 25,
        "Storage Optimization": 25,
        "Network Optimization": 25,
        "Data Performance": 20,
        "Scalability": 10,
        "Monitoring and Alerting": 5,
        "Service Upgrade and Retirement": 5,
        "Other": 5,
    }
    OPERATIONAL_WEIGHTS = {
        "Efficiency Optimization": 30,
        "Failure Mitigation": 20,
        "Scalability": 10,
        "Monitoring and Alerting": 5,
        "Safe and Secure Deployment": 5,
        "Service Upgrade and Retirement": 5,
        "Other": 5,
    }

    RELIABILITY_RULES = [
        ("Zone Resiliency", ["availability zone", "zone redundant", "zone-redundant", "zonal"]),
        ("Regional Resiliency", ["paired region", "geo-redundant", "geo redundant", "multi-region", "georeplication", "geo replication", "ddos"]),
        ("Data Protection and Recovery", ["backup", "recovery vault", "restore", "snapshot", "soft delete", "point-in-time"]),
        ("Governance and Compliance", ["policy", "compliance", "governance"]),
        ("Scalability", ["scale", "autoscale", "throughput", "capacity"]),
        ("Monitoring and Alerting", ["diagnostic", "alert", "monitor", "log analytics", "insights"]),
        ("Service Upgrade and Retirement", ["upgrade", "retire", "deprecat", "migrate", "end of life", "eol", "outdated version"]),
    ]
    PERFORMANCE_RULES = [
        ("Compute Optimization", ["virtual machine", "vm", "app service plan", "compute"]),
        ("Storage Optimization", ["storage account", "disk", "data warehouse", "blob"]),
        ("Network Optimization", ["traffic manager", "network", "express route", "vpn", "load balancer", "cdn", "front door"]),
        ("Data Performance", ["sql", "database", "cosmos", "query"]),
        ("Scalability", ["scale", "autoscale", "throughput", "capacity"]),
        ("Monitoring and Alerting", ["diagnostic", "alert", "monitor", "log analytics", "insights"]),
        ("Service Upgrade and Retirement", ["upgrade", "retire", "deprecat", "migrate", "end of life", "eol", "outdated version"]),
    ]
    OPERATIONAL_RULES = [
        ("Efficiency Optimization", ["accelerated networking", "efficiency", "configuration"]),
        ("Failure Mitigation", ["failure", "deployment failure", "resiliency", "mitigat"]),
        ("Safe and Secure Deployment", ["safe deployment", "secure deployment", "rollout", "blue-green", "canary", "staged rollout"]),
        ("Scalability", ["scale", "autoscale", "throughput", "capacity"]),
        ("Monitoring and Alerting", ["diagnostic", "alert", "monitor", "log analytics", "insights"]),
        ("Service Upgrade and Retirement", ["upgrade", "retire", "deprecat", "migrate", "end of life", "eol", "outdated version"]),
    ]

    @staticmethod
    def classify(text: str, rules) -> str:
        t = (text or "").lower()
        for subcategory, keywords in rules:
            for kw in keywords:
                if re.search(r"\b" + re.escape(kw.strip()) + r"\b", t):
                    return subcategory
        return "Other"

    @staticmethod
    def bucket(records, rules) -> dict:
        """Group impacted-resource keys into subcategories via keyword classification."""
        buckets = {}
        for rec in records:
            text = " ".join(str(rec.get(f, "")) for f in ("problem", "solution", "impactedType", "title"))
            subcategory = AdvisorScoreModel.classify(text, rules)
            key = str(rec.get("impactedResource") or rec.get("name") or rec.get("Name") or id(rec))
            buckets.setdefault(subcategory, set()).add(key)
        return buckets

    @staticmethod
    def merge_buckets(*bucket_dicts) -> dict:
        merged = {}
        for buckets in bucket_dicts:
            for subcategory, keys in buckets.items():
                merged.setdefault(subcategory, set()).update(keys)
        return merged

    @staticmethod
    def score_category(weights: dict, buckets: dict, total_pool: int, forced_zero: set = frozenset()):
        """Official formula: sum((Healthy/Total)*Weight) / sum(Weight) * 100."""
        if total_pool <= 0:
            return None, []
        rows = []
        weighted_sum = 0.0
        weight_total = 0
        for subcategory, weight in weights.items():
            if subcategory in forced_zero:
                impacted = total_pool
            else:
                impacted = min(len(buckets.get(subcategory, set())), total_pool)
            healthy = max(0, total_pool - impacted)
            pct = (healthy / total_pool) * 100
            weighted_sum += pct * weight
            weight_total += weight
            rows.append({
                "subcategory": subcategory, "weight": weight,
                "healthy": healthy, "total": total_pool, "pct": round(pct, 1)
            })
        score = weighted_sum / weight_total if weight_total else None
        return (round(score) if score is not None else None), rows


# ═══════════════════════════════════════════════════════════════════════════════
# ANALYSIS ENGINE
# ═══════════════════════════════════════════════════════════════════════════════

class InsightEngine:
    """Analyzes collected data, scores each WAF pillar, and generates action items."""
    
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
        self.breakdowns = {}
        self.max_score = 100

    def _resource_pool(self) -> int:
        """Total discovered resources, used as the shared 'assessed resource' pool."""
        summary = self.data.get("discovery", {}).get("ResourceSummary", [])
        return sum(int(r.get("Count", 0) or 0) for r in summary)

    @staticmethod
    def _data_rows(rows: list) -> list:
        """Exclude sentinel rows emitted when a workbook query returns no data."""
        empty_results = {"no data found", "no resources found"}
        return [
            row for row in rows
            if str(row.get("Result", row.get("result", ""))).strip().lower() not in empty_results
            and str(row.get("DataStatus", "")).strip().lower() != "nodata"
        ]
    
    def analyze(self):
        """Run all analysis passes."""
        self._analyze_reliability()
        self._analyze_security()
        self._analyze_cost()
        self._analyze_operational_excellence()
        self._analyze_performance()
        return self
    
    def _analyze_reliability(self):
        """Reliability score using the official Advisor Score subcategory model."""
        total_pool = self._resource_pool()
        forced_zero = set()
        synthetic = []

        # Check VMs without availability zones
        vms = self.data.get("discovery", {}).get("VMs", [])
        vms_no_zone = [v for v in vms if not v.get("zone") and not v.get("availabilityZone")]
        synthetic += [{"problem": "Virtual machine not deployed to an Availability Zone",
                       "impactedResource": v.get("name", v.get("Name", ""))} for v in vms_no_zone]
        if vms_no_zone:
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "high",
                "title": f"{len(vms_no_zone)} VMs without Availability Zones",
                "title_es": f"{len(vms_no_zone)} VMs sin Zonas de Disponibilidad",
                "description": "These VMs are not protected against datacenter failures. Consider migrating to zone-redundant deployments.",
                "description_es": "Estas VMs no están protegidas contra fallas del centro de datos. Considere migrar a implementaciones con redundancia de zona.",
                "resources": [v.get("name", v.get("Name", "")) for v in vms_no_zone[:10]]
            })
        
        # Check backup vaults
        backups = self.data.get("discovery", {}).get("BackupVaults", [])
        if not backups or (len(backups) == 1 and backups[0].get("Result") == "No resources found"):
            forced_zero.add("Data Protection and Recovery")
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "critical",
                "title": "No Backup/Recovery Services Vaults found",
                "title_es": "No se encontraron Vaults de Backup/Recovery Services",
                "description": "No backup infrastructure detected. Critical workloads should have backup policies configured.",
                "description_es": "No se detectó infraestructura de copia de seguridad. Las cargas de trabajo críticas deben tener políticas de backup configuradas.",
                "resources": []
            })
        
        # Check DDoS protection on VNets
        vnets = self.data.get("discovery", {}).get("VNets", []) or self.data.get("governance", {}).get("VNets", [])
        vnets_no_ddos = [v for v in vnets if str(v.get("ddos", v.get("ddosProtection", v.get("enableDdosProtection", "")))).lower() in ("false", "", "none")]
        synthetic += [{"problem": "VNet without DDoS Protection Standard",
                       "impactedResource": v.get("name", v.get("Name", ""))} for v in vnets_no_ddos]
        if vnets_no_ddos:
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "medium",
                "title": f"{len(vnets_no_ddos)} VNets without DDoS Protection",
                "title_es": f"{len(vnets_no_ddos)} VNets sin Protección DDoS",
                "description": "Enable Azure DDoS Protection Standard for internet-facing workloads.",
                "description_es": "Habilite Azure DDoS Protection Standard para cargas de trabajo expuestas a Internet.",
                "resources": [v.get("name", v.get("Name", "")) for v in vnets_no_ddos[:10]]
            })
        
        # Advisor HighAvailability recommendations
        ha_recs = self.data.get("advisor", {}).get("Reliability", [])
        high_impact_ha = [r for r in ha_recs if str(r.get("impact", "")).lower() == "high"]
        if high_impact_ha:
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "high",
                "title": f"{len(high_impact_ha)} High-Impact Reliability recommendations from Advisor",
                "title_es": f"{len(high_impact_ha)} Recomendaciones de Confiabilidad de alto impacto de Advisor",
                "description": "Azure Advisor has identified critical reliability improvements.",
                "description_es": "Azure Advisor identificó mejoras críticas de confiabilidad.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in high_impact_ha[:10]]
            })

        buckets = AdvisorScoreModel.merge_buckets(
            AdvisorScoreModel.bucket(ha_recs, AdvisorScoreModel.RELIABILITY_RULES),
            AdvisorScoreModel.bucket(synthetic, AdvisorScoreModel.RELIABILITY_RULES),
        )
        score, rows = AdvisorScoreModel.score_category(
            AdvisorScoreModel.RELIABILITY_WEIGHTS, buckets, total_pool, forced_zero)
        self.scores["Reliability"] = score if score is not None else 0
        self.breakdowns["Reliability"] = {"available": score is not None, "rows": rows, "pool": total_pool}
    
    def _analyze_security(self):
        """Security score = Microsoft Defender Secure Score average (same model Advisor uses)."""
        security = self.data.get("security", {})
        secure_scores = self._data_rows(security.get("SecureScores", []))
        if not secure_scores:
            secure_scores = self._data_rows(self.data.get("governance", {}).get("SecureScores", []))
        avg_score = None
        if secure_scores:
            values = [
                float(s.get("percentageScore", s.get("pct", s.get("percentage", 0))) or 0)
                for s in secure_scores
            ]
            avg_score = sum(values) / len(values)
            if avg_score <= 1:  # export may store a 0-1 ratio instead of 0-100
                avg_score *= 100
            self.insights.append({
                "pillar": "Security",
                "text": f"Average Defender Secure Score: {avg_score:.0f}%"
            })
        
        # Storage accounts with public access
        storage = self.data.get("discovery", {}).get("Storage", [])
        public_storage = [s for s in storage if str(s.get("publicBlob", s.get("allowBlobPublicAccess", ""))).lower() == "true"]
        if public_storage:
            self.action_items.append({
                "pillar": "Security",
                "severity": "critical",
                "title": f"{len(public_storage)} Storage accounts allow public blob access",
                "title_es": f"{len(public_storage)} Cuentas de almacenamiento permiten acceso público a blobs",
                "description": "Disable public blob access unless explicitly required. This is a common data exposure risk.",
                "description_es": "Deshabilite el acceso público a blobs a menos que sea explícitamente necesario. Este es un riesgo común de exposición de datos.",
                "resources": [s.get("name", s.get("Name", "")) for s in public_storage[:10]]
            })
        
        # Key Vaults without purge protection
        keyvaults = self.data.get("discovery", {}).get("KeyVaults", []) or self.data.get("governance", {}).get("KeyVaults", [])
        kv_no_purge = [k for k in keyvaults if str(k.get("purgeProtection", k.get("enablePurgeProtection", ""))).lower() in ("false", "", "none")]
        if kv_no_purge:
            self.action_items.append({
                "pillar": "Security",
                "severity": "high",
                "title": f"{len(kv_no_purge)} Key Vaults without Purge Protection",
                "title_es": f"{len(kv_no_purge)} Key Vaults sin Protección contra Purga",
                "description": "Enable purge protection to prevent permanent deletion of secrets, keys, and certificates.",
                "description_es": "Habilite la protección contra purga para evitar la eliminación permanente de secretos, claves y certificados.",
                "resources": [k.get("name", k.get("Name", "")) for k in kv_no_purge[:10]]
            })
        
        # Storage without TLS 1.2
        old_tls = [s for s in storage if s.get("tlsVersion", s.get("minimumTlsVersion", "")) and "1.2" not in str(s.get("tlsVersion", s.get("minimumTlsVersion", "")))]
        if old_tls:
            self.action_items.append({
                "pillar": "Security",
                "severity": "high",
                "title": f"{len(old_tls)} Storage accounts not enforcing TLS 1.2",
                "title_es": f"{len(old_tls)} Cuentas de almacenamiento sin exigir TLS 1.2",
                "description": "Enforce minimum TLS 1.2 for all storage accounts to prevent protocol downgrade attacks.",
                "description_es": "Exija TLS 1.2 como mínimo en todas las cuentas de almacenamiento para evitar ataques de downgrade de protocolo.",
                "resources": [s.get("name", s.get("Name", "")) for s in old_tls[:10]]
            })
        
        # Security advisor recommendations (informational — already reflected in Secure Score)
        sec_recs = self._data_rows(self.data.get("advisor", {}).get("Security", []))
        if sec_recs:
            self.action_items.append({
                "pillar": "Security",
                "severity": "medium",
                "title": f"{len(sec_recs)} Security recommendations from Advisor",
                "title_es": f"{len(sec_recs)} Recomendaciones de seguridad de Advisor",
                "description": "Review Azure Advisor security findings and remediate.",
                "description_es": "Revise los hallazgos de seguridad de Azure Advisor y corríjalos.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in sec_recs[:10]]
            })
        
        # Broad-scope Owner/Contributor role assignments (RBAC hygiene, informational)
        OWNER_ROLE_ID = "8e3af657-a8ff-443c-a75c-2fe8c4bcb635"
        CONTRIBUTOR_ROLE_ID = "b24988ac-6180-42a0-ab88-20f7382dd24c"
        role_assignments = self.data.get("governance", {}).get("RoleAssignments", [])
        broad_privileged = [
            r for r in role_assignments
            if any(rid in str(r.get("roleDefId", "")) for rid in (OWNER_ROLE_ID, CONTRIBUTOR_ROLE_ID))
            and "resourceGroups" not in str(r.get("scope_", ""))
        ]
        if broad_privileged:
            self.action_items.append({
                "pillar": "Security",
                "severity": "medium",
                "title": f"{len(broad_privileged)} Owner/Contributor role assignments at subscription scope or higher",
                "title_es": f"{len(broad_privileged)} Asignaciones de rol Owner/Contributor con alcance de suscripción o superior",
                "description": "Review broad-scope privileged role assignments for least-privilege access. Prefer scoping to resource groups.",
                "description_es": "Revise las asignaciones de roles privilegiados de alcance amplio para aplicar el principio de mínimo privilegio. Prefiera limitar el alcance a grupos de recursos.",
                "resources": [f"{r.get('principalType', '')} ({r.get('scope_', '')})" for r in broad_privileged[:10]]
            })
        
        # Microsoft Defender for Cloud plans not enabled (informational)
        defender_plans = self._data_rows(security.get("DefenderPlans", []))
        if not defender_plans:
            defender_plans = self._data_rows(self.data.get("governance", {}).get("DefenderPlans", []))
        disabled_plans = [
            p for p in defender_plans
            if str(p.get("pricingTier", p.get("tier", ""))).lower() == "free"
        ]
        if disabled_plans:
            self.action_items.append({
                "pillar": "Security",
                "severity": "medium",
                "title": f"{len(disabled_plans)} Microsoft Defender for Cloud plans not enabled (Free tier)",
                "title_es": f"{len(disabled_plans)} Planes de Microsoft Defender for Cloud no habilitados (nivel gratuito)",
                "description": "Enable Microsoft Defender for Cloud plans for full threat protection coverage.",
                "description_es": "Habilite los planes de Microsoft Defender for Cloud para una cobertura completa de protección contra amenazas.",
                "resources": sorted({p.get("planName", p.get("name", "")) for p in disabled_plans})[:10]
            })

        posture_recommendations = self._data_rows(security.get("Recommendations", []))
        open_recommendations = [
            recommendation for recommendation in posture_recommendations
            if str(recommendation.get("recommendationState", "")).lower() not in ("healthy", "notapplicable")
        ]
        severe_recommendations = [
            recommendation for recommendation in open_recommendations
            if str(recommendation.get("recommendationSeverity", "")).lower() in ("critical", "high")
        ]
        if severe_recommendations:
            self.action_items.append({
                "pillar": "Security",
                "severity": "high",
                "title": f"{len(severe_recommendations)} high-severity Defender for Cloud recommendations",
                "title_es": f"{len(severe_recommendations)} recomendaciones de alta severidad de Defender for Cloud",
                "description": "Prioritize unhealthy CSPM recommendations that carry critical or high severity.",
                "description_es": "Priorice las recomendaciones de CSPM no saludables con severidad crítica o alta.",
                "resources": [
                    recommendation.get("recommendationName", recommendation.get("affectedResourceId", ""))
                    for recommendation in severe_recommendations[:10]
                ]
            })

        score_controls = self._data_rows(security.get("ScoreControls", []))
        unhealthy_control_resources = sum(
            int(control.get("unhealthyResourceCount", 0) or 0) for control in score_controls
        )
        mcsb_assessments = self._data_rows(security.get("MCSBCompliance", []))
        mcsb_failed_resources = sum(
            int(assessment.get("failedResources", 0) or 0) for assessment in mcsb_assessments
        )

        incidents = self._data_rows(security.get("Incidents", []))
        active_incidents = [
            incident for incident in incidents
            if str(incident.get("Status", "")).lower() not in ("resolved", "redirected")
        ]
        severe_incidents = [
            incident for incident in active_incidents
            if str(incident.get("Severity", "")).lower() in ("critical", "high")
        ]
        if severe_incidents:
            self.action_items.append({
                "pillar": "Security",
                "severity": "critical",
                "title": f"{len(severe_incidents)} active high-severity security incidents",
                "title_es": f"{len(severe_incidents)} incidentes de seguridad activos de alta severidad",
                "description": "Investigate active Microsoft Defender XDR incidents and confirm ownership and containment.",
                "description_es": "Investigue los incidentes activos de Microsoft Defender XDR y confirme su asignación y contención.",
                "resources": [incident.get("DisplayName", incident.get("IncidentId", "")) for incident in severe_incidents[:10]]
            })

        graph_alerts = self._data_rows(security.get("Alerts", []))
        cloud_alerts = self._data_rows(security.get("CloudAlerts", []))
        active_alerts = [
            alert for alert in graph_alerts
            if str(alert.get("Status", "")).lower() not in ("resolved", "dismissed")
        ]
        active_cloud_alerts = [
            alert for alert in cloud_alerts
            if str(alert.get("status", "")).lower() not in ("resolved", "dismissed")
        ]
        source_status = self._data_rows(security.get("SourceStatus", []))
        
        self.scores["Security"] = round(avg_score) if avg_score is not None else 0
        self.breakdowns["Security"] = {
            "available": avg_score is not None,
            "secure_score_pct": round(avg_score, 1) if avg_score is not None else None,
            "subscriptions_scored": len(secure_scores),
            "findings": [
                ("Public Blob Access on Storage Accounts", "Acceso p\u00fablico a blobs en cuentas de almacenamiento", len(public_storage)),
                ("Key Vaults without Purge Protection", "Key Vaults sin Protecci\u00f3n contra Purga", len(kv_no_purge)),
                ("Storage Accounts below TLS 1.2", "Cuentas de almacenamiento por debajo de TLS 1.2", len(old_tls)),
                ("Broad-scope Owner/Contributor Assignments", "Asignaciones Owner/Contributor de alcance amplio", len(broad_privileged)),
                ("Defender for Cloud Plans on Free Tier", "Planes de Defender for Cloud en nivel gratuito", len(disabled_plans)),
                ("Open Defender for Cloud Recommendations", "Recomendaciones abiertas de Defender for Cloud", len(open_recommendations)),
                ("Unhealthy Secure Score Control Resources", "Recursos no saludables en controles de Secure Score", unhealthy_control_resources),
                ("Failed MCSB Resources", "Recursos con error en MCSB", mcsb_failed_resources),
                ("Active Defender XDR Incidents", "Incidentes activos de Defender XDR", len(active_incidents)),
                ("Active Security Alerts", "Alertas de seguridad activas", len(active_alerts) + len(active_cloud_alerts)),
                ("Advisor Security Recommendations", "Recomendaciones de seguridad de Advisor", len(sec_recs)),
            ],
            "open_recommendations": len(open_recommendations),
            "severe_recommendations": len(severe_recommendations),
            "unhealthy_control_resources": unhealthy_control_resources,
            "mcsb_assessments": len(mcsb_assessments),
            "mcsb_failed_resources": mcsb_failed_resources,
            "defender_plans": len(defender_plans),
            "disabled_plans": len(disabled_plans),
            "active_incidents": len(active_incidents),
            "severe_incidents": len(severe_incidents),
            "active_alerts": len(active_alerts) + len(active_cloud_alerts),
            "source_status": source_status,
        }
    
    def _analyze_cost(self):
        """Cost score = healthy-resource ratio (count-based proxy for Advisor's retail-cost weighting)."""
        total_pool = self._resource_pool()
        impacted_keys = set()
        sources = []
        
        # Orphaned resources
        orphaned = self._data_rows(self.data.get("governance", {}).get("OrphanedResources", []))
        if not orphaned:
            # Try discovery sheets
            disks = self._data_rows(self.data.get("discovery", {}).get("UnattachedDisks", []))
            orphaned = disks
        
        if orphaned:
            keys = {str(o.get("Name", o.get("name", id(o)))) for o in orphaned}
            impacted_keys |= keys
            sources.append({
                "name": "Orphaned/unused resources",
                "count": len(keys),
                "unit": "resources",
                "affected": len(keys),
            })
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "medium",
                "title": f"{len(orphaned)} Orphaned/unused resources detected",
                "title_es": f"{len(orphaned)} Recursos huérfanos/no utilizados detectados",
                "description": "These resources incur costs but are not attached to any workload. Review and delete if unused.",
                "description_es": "Estos recursos generan costos pero no están asociados a ninguna carga de trabajo. Revíselos y elimínelos si no se utilizan.",
                "resources": [o.get("Name", o.get("name", "")) for o in orphaned[:10]]
            })
        
        # VMs that are deallocated (still paying for disks)
        dealloc_vms = self._data_rows(self.data.get("discovery", {}).get("DeallocatedVMs", []))
        if dealloc_vms:
            keys = {str(v.get("name", v.get("Name", id(v)))) for v in dealloc_vms}
            impacted_keys |= keys
            sources.append({
                "name": "Deallocated VMs",
                "count": len(keys),
                "unit": "resources",
                "affected": len(keys),
            })
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "low",
                "title": f"{len(dealloc_vms)} Deallocated VMs (still paying for disks/IPs)",
                "title_es": f"{len(dealloc_vms)} VMs desasignadas (aún generan costo por discos/IPs)",
                "description": "Deallocated VMs still incur costs for attached disks and static IPs. Consider deleting if no longer needed.",
                "description_es": "Las VMs desasignadas siguen generando costos por los discos e IPs estáticas asociadas. Considere eliminarlas si ya no son necesarias.",
                "resources": [v.get("name", v.get("Name", "")) for v in dealloc_vms[:10]]
            })
        
        # Right-sizing from metrics
        vm_metrics = self._data_rows(self.data.get("metrics", {}).get("VM_RightSizing", []))
        underutilized = [v for v in vm_metrics if "Idle" in str(v.get("Assessment", "")) or "Underutilized" in str(v.get("Assessment", ""))]
        if underutilized:
            keys = {str(v.get("Name", id(v))) for v in underutilized}
            impacted_keys |= keys
            sources.append({
                "name": "Idle/underutilized VMs",
                "count": len(keys),
                "unit": "resources",
                "affected": len(keys),
            })
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "high",
                "title": f"{len(underutilized)} VMs are idle or underutilized",
                "title_es": f"{len(underutilized)} VMs están inactivas o subutilizadas",
                "description": "These VMs have <15% average CPU over 30 days. Consider downsizing or deallocating.",
                "description_es": "Estas VMs tienen <15% de uso promedio de CPU en 30 días. Considere reducir su tamaño o desasignarlas.",
                "resources": [f"{v.get('Name', '')} ({v.get('VMSize', '')} @ {v.get('AvgCPU_Pct', '?')}% CPU)" for v in underutilized[:10]]
            })
        
        # Cost recommendations from Advisor
        cost_recs = self._data_rows(self.data.get("advisor", {}).get("Cost", []))
        if cost_recs:
            keys = {str(r.get("impactedResource", r.get("impactedValue", id(r)))) for r in cost_recs}
            impacted_keys |= keys
            sources.append({
                "name": "Advisor Cost recommendations",
                "count": len(cost_recs),
                "unit": "recommendations",
                "affected": len(keys),
            })
            total_savings = sum(float(r.get("annualSavings", r.get("savingsAmount", 0)) or 0) for r in cost_recs)
            desc = "Azure Advisor estimates potential savings."
            desc_es = "Azure Advisor estima posibles ahorros."
            if total_savings > 0:
                desc = f"Azure Advisor estimates ~${total_savings:,.0f} in potential annual savings."
                desc_es = f"Azure Advisor estima ~${total_savings:,.0f} en posibles ahorros anuales."
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "high",
                "title": f"{len(cost_recs)} Cost optimization recommendations",
                "title_es": f"{len(cost_recs)} Recomendaciones de optimización de costos",
                "description": desc,
                "description_es": desc_es,
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in cost_recs[:10]]
            })
        
        if total_pool > 0:
            impacted = min(len(impacted_keys), total_pool)
            healthy = max(0, total_pool - impacted)
            score = round((healthy / total_pool) * 100)
        else:
            healthy, score = 0, None
        
        self.scores["Cost Optimization"] = score if score is not None else 0
        self.breakdowns["Cost Optimization"] = {
            "available": score is not None,
            "sources": sources,
            "impacted": min(len(impacted_keys), total_pool),
            "healthy": healthy,
            "total": total_pool,
        }
    
    def _analyze_operational_excellence(self):
        """Operational Excellence score using the official Advisor Score subcategory model."""
        total_pool = self._resource_pool()
        synthetic = []
        
        # Diagnostic settings coverage
        diag = self.data.get("metrics", {}).get("DiagnosticsCoverage", [])
        if diag:
            no_diag = [d for d in diag if str(d.get("HasDiagnostics", d.get("Gap", ""))).lower() in ("false", "no diagnostics configured")]
            if no_diag:
                pct_missing = len(no_diag) / len(diag) * 100
                synthetic += [{"problem": "Resource missing diagnostic settings",
                               "impactedResource": d.get("Name", d.get("name", ""))} for d in no_diag]
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "high",
                    "title": f"{len(no_diag)}/{len(diag)} critical resources missing diagnostic settings ({pct_missing:.0f}%)",
                    "title_es": f"{len(no_diag)}/{len(diag)} recursos críticos sin configuración de diagnóstico ({pct_missing:.0f}%)",
                    "description": "Configure diagnostic settings to send logs to Log Analytics for observability and troubleshooting.",
                    "description_es": "Configure ajustes de diagnóstico para enviar registros a Log Analytics para observabilidad y solución de problemas.",
                    "resources": [d.get("Name", d.get("name", "")) for d in no_diag[:10]]
                })
        
        # Tag coverage — computed from the per-resource `tags` field across Discovery inventory
        # sheets (there's no dedicated tag-usage export; GovViz's Subscriptions.tags is subscription-level only).
        TAGGABLE_SHEETS = ["VMs", "AppServices", "AKS", "VNets", "NSGs", "LoadBalancers",
                           "Firewalls", "PublicIPs", "Storage", "Databases", "KeyVaults"]
        discovery = self.data.get("discovery", {})
        total_taggable, untagged_names = 0, []
        for sheet in TAGGABLE_SHEETS:
            for r in discovery.get(sheet, []):
                total_taggable += 1
                tags = r.get("tags")
                if not tags or str(tags).strip().lower() in ("", "none", "{}"):
                    untagged_names.append(r.get("name", r.get("Name", "")))
        if total_taggable > 0 and untagged_names:
            pct_untagged = len(untagged_names) / total_taggable * 100
            synthetic += [{"problem": "Resource has no tags configured", "impactedResource": n} for n in untagged_names]
            self.action_items.append({
                "pillar": "Operational Excellence",
                "severity": "medium",
                "title": f"{len(untagged_names)}/{total_taggable} resources have no tags ({pct_untagged:.0f}%)",
                "title_es": f"{len(untagged_names)}/{total_taggable} recursos no tienen etiquetas ({pct_untagged:.0f}%)",
                "description": "Implement a tagging strategy (Owner, CostCenter, Environment, Application) for governance and cost allocation.",
                "description_es": "Implemente una estrategia de etiquetado (Owner, CostCenter, Environment, Application) para gobernanza y asignación de costos.",
                "resources": untagged_names[:10]
            })
        
        # Resource locks (protects against accidental delete/modify of critical resources)
        locks = self.data.get("governance", {}).get("Locks", [])
        has_locks = locks and not (len(locks) == 1 and locks[0].get("Result"))
        if not has_locks:
            self.action_items.append({
                "pillar": "Operational Excellence",
                "severity": "medium",
                "title": "No resource locks configured",
                "title_es": "No hay bloqueos de recursos configurados",
                "description": "Apply CanNotDelete or ReadOnly locks to critical resources (networking, databases, key vaults) to prevent accidental deletion or modification.",
                "description_es": "Aplique bloqueos CanNotDelete o ReadOnly a los recursos críticos (redes, bases de datos, key vaults) para evitar eliminaciones o modificaciones accidentales.",
                "resources": []
            })
        
        # Policy compliance
        compliance = self.data.get("governance", {}).get("PolicyCompliance", [])
        if compliance:
            total_nc = sum(int(c.get("NonCompliantCount", 0) or 0) for c in compliance)
            non_compliant = [c for c in compliance if int(c.get("NonCompliantCount", 0) or 0) > 0]
            synthetic += [{"problem": "Non-compliant policy evaluation",
                           "impactedResource": c.get("policyAssignment", c.get("policyAssignmentName", ""))} for c in non_compliant]
            if total_nc > 0:
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "medium",
                    "title": f"{total_nc} non-compliant policy evaluations",
                    "title_es": f"{total_nc} evaluaciones de políticas no conformes",
                    "description": "Review policy compliance and remediate non-compliant resources or adjust policy assignments.",
                    "description_es": "Revise el cumplimiento de políticas y corrija los recursos no conformes o ajuste las asignaciones de políticas.",
                    "resources": [f"{c.get('policyAssignment', c.get('policyAssignmentName', ''))} ({c.get('NonCompliantCount', 0)} violations)" for c in compliance[:10]]
                })
        
        # OpEx Advisor recommendations
        opex_recs = self.data.get("advisor", {}).get("OperationalExcellence", [])
        if opex_recs:
            self.action_items.append({
                "pillar": "Operational Excellence",
                "severity": "medium",
                "title": f"{len(opex_recs)} Operational Excellence recommendations from Advisor",
                "title_es": f"{len(opex_recs)} Recomendaciones de Excelencia Operativa de Advisor",
                "description": "Azure Advisor has identified operational improvements.",
                "description_es": "Azure Advisor identificó mejoras operativas.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in opex_recs[:10]]
            })

        buckets = AdvisorScoreModel.merge_buckets(
            AdvisorScoreModel.bucket(opex_recs, AdvisorScoreModel.OPERATIONAL_RULES),
            AdvisorScoreModel.bucket(synthetic, AdvisorScoreModel.OPERATIONAL_RULES),
        )
        score, rows = AdvisorScoreModel.score_category(AdvisorScoreModel.OPERATIONAL_WEIGHTS, buckets, total_pool)
        self.scores["Operational Excellence"] = score if score is not None else 0
        self.breakdowns["Operational Excellence"] = {"available": score is not None, "rows": rows, "pool": total_pool}
    
    def _analyze_performance(self):
        """Performance score using the official Advisor Score subcategory model."""
        total_pool = self._resource_pool()
        synthetic = []
        
        # Saturated VMs
        vm_metrics = self.data.get("metrics", {}).get("VM_RightSizing", [])
        saturated = [v for v in vm_metrics if "Saturated" in str(v.get("Assessment", ""))]
        synthetic += [{"problem": "Virtual machine CPU saturated",
                       "impactedResource": v.get("Name", "")} for v in saturated]
        if saturated:
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "high",
                "title": f"{len(saturated)} VMs are saturated (>80% CPU)",
                "title_es": f"{len(saturated)} VMs están saturadas (>80% CPU)",
                "description": "These VMs are consistently at high CPU usage. Consider scaling up or scaling out.",
                "description_es": "Estas VMs presentan un uso de CPU consistentemente alto. Considere escalar verticalmente u horizontalmente.",
                "resources": [f"{v.get('Name', '')} ({v.get('VMSize', '')} @ {v.get('AvgCPU_Pct', '?')}% CPU)" for v in saturated[:10]]
            })
        
        # Saturated SQL
        sql_metrics = self.data.get("metrics", {}).get("SQL_RightSizing", [])
        saturated_sql = [s for s in sql_metrics if "Saturated" in str(s.get("Assessment", ""))]
        synthetic += [{"problem": "SQL database DTU/CPU saturated",
                       "impactedResource": s.get("Name", "")} for s in saturated_sql]
        if saturated_sql:
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "high",
                "title": f"{len(saturated_sql)} SQL Databases are saturated (>80% DTU/CPU)",
                "title_es": f"{len(saturated_sql)} Bases de datos SQL están saturadas (>80% DTU/CPU)",
                "description": "These databases are consistently at high utilization. Consider scaling up the tier.",
                "description_es": "Estas bases de datos presentan una utilización consistentemente alta. Considere aumentar el nivel de servicio.",
                "resources": [f"{s.get('Name', '')} ({s.get('SKU', '')} @ {s.get('AvgUsage_Pct', '?')}%)" for s in saturated_sql[:10]]
            })
        
        # Performance advisor recommendations
        perf_recs = self.data.get("advisor", {}).get("Performance", [])
        if perf_recs:
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "medium",
                "title": f"{len(perf_recs)} Performance recommendations from Advisor",
                "title_es": f"{len(perf_recs)} Recomendaciones de Rendimiento de Advisor",
                "description": "Azure Advisor has identified performance improvements.",
                "description_es": "Azure Advisor identificó mejoras de rendimiento.",
                "resources": [r.get("impactedResource", r.get("impactedValue", "")) for r in perf_recs[:10]]
            })

        buckets = AdvisorScoreModel.merge_buckets(
            AdvisorScoreModel.bucket(perf_recs, AdvisorScoreModel.PERFORMANCE_RULES),
            AdvisorScoreModel.bucket(synthetic, AdvisorScoreModel.PERFORMANCE_RULES),
        )
        score, rows = AdvisorScoreModel.score_category(AdvisorScoreModel.PERFORMANCE_WEIGHTS, buckets, total_pool)
        self.scores["Performance Efficiency"] = score if score is not None else 0
        self.breakdowns["Performance Efficiency"] = {"available": score is not None, "rows": rows, "pool": total_pool}
    
    def get_overall_score(self) -> int:
        """Mean of the available category scores (per the official Advisor Score model)."""
        available = [p for p, b in self.breakdowns.items() if b.get("available")]
        if not available:
            return 0
        return round(sum(self.scores[p] for p in available) / len(available))


# ═══════════════════════════════════════════════════════════════════════════════
# HTML DASHBOARD GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

class DashboardGenerator:
    """Generates the final consolidated HTML dashboard."""
    
    def __init__(self, engine: InsightEngine, data: dict, base_dir: Optional[str] = None):
        self.engine = engine
        self.data = data
        self.base_dir = base_dir

    @staticmethod
    def _count_real(rows: list) -> int:
        """Count sheet rows, ignoring the Export-Sheet 'No data found' placeholder row."""
        return len([
            row for row in rows
            if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
        ])

    def _render_security_view(self) -> str:
        """Render dedicated security posture, compliance, and operations details."""
        security = self.data.get("security", {})
        breakdown = self.engine.breakdowns.get("Security", {})

        def real_rows(sheet_name):
            return [
                row for row in security.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def as_int(value):
            try:
                return int(value or 0)
            except (TypeError, ValueError):
                return 0

        def render_table(rows, columns, empty_en, empty_es, limit=12):
            if not rows:
                return f'<p class="empty-state">{bi(empty_en, empty_es)}</p>'
            header = "".join(f"<th>{bi(label_en, label_es)}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for row in rows[:limit]:
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{escape(str(row.get(key, "")))}</td>'
                    for key, _, _, numeric in columns
                )
                body += f"<tr>{cells}</tr>"
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>'

        severity_rank = {"critical": 0, "high": 1, "medium": 2, "low": 3, "informational": 4}
        recommendations = [
            row for row in real_rows("Recommendations")
            if str(row.get("recommendationState", "")).lower() not in ("healthy", "notapplicable")
        ]
        recommendations.sort(key=lambda row: severity_rank.get(str(row.get("recommendationSeverity", "")).lower(), 5))
        controls = sorted(
            real_rows("ScoreControls"),
            key=lambda row: as_int(row.get("unhealthyResourceCount")),
            reverse=True,
        )
        mcsb = sorted(
            real_rows("MCSBCompliance"),
            key=lambda row: as_int(row.get("failedResources")),
            reverse=True,
        )
        defender_plans = real_rows("DefenderPlans")
        source_status = real_rows("SourceStatus")
        incidents = sorted(real_rows("Incidents"), key=lambda row: str(row.get("LastUpdateDateTime", "")), reverse=True)
        alerts = sorted(real_rows("Alerts"), key=lambda row: str(row.get("LastUpdateDateTime", "")), reverse=True)
        endpoint_recommendations = sorted(
            real_rows("EndpointRecommendations"),
            key=lambda row: as_int(row.get("ExposedMachinesCount")),
            reverse=True,
        )
        vulnerabilities = sorted(
            real_rows("Vulnerabilities"),
            key=lambda row: severity_rank.get(str(row.get("Severity", "")).lower(), 5),
        )
        machines = real_rows("Machines")
        at_risk_machines = len([
            machine for machine in machines
            if str(machine.get("RiskScore", "")).lower() in ("critical", "high")
        ])

        secure_score = breakdown.get("secure_score_pct")
        secure_score_text = f"{secure_score:.1f}%" if secure_score is not None else "N/A"
        stats = [
            (secure_score_text, bi("Defender Secure Score", "Secure Score de Defender"), "security-accent"),
            (breakdown.get("open_recommendations", 0), bi("Open CSPM Recommendations", "Recomendaciones CSPM abiertas"), "security-danger"),
            (breakdown.get("mcsb_failed_resources", 0), bi("Failed MCSB Resources", "Recursos con error en MCSB"), "security-warning"),
            (breakdown.get("active_incidents", 0), bi("Active XDR Incidents", "Incidentes XDR activos"), "security-danger"),
            (breakdown.get("active_alerts", 0), bi("Active Security Alerts", "Alertas de seguridad activas"), "security-warning"),
            (at_risk_machines, bi("High-risk Endpoints", "Endpoints de alto riesgo"), "security-danger"),
        ]
        stat_html = "".join(
            f'<div class="security-stat"><div class="value {color_class}">{value}</div><div class="label">{label}</div></div>'
            for value, label, color_class in stats
        )

        status_rows = ""
        for row in source_status:
            status = str(row.get("Status", "Unknown"))
            status_class = re.sub(r"[^a-z]", "", status.lower())
            status_rows += f"""<tr>
                <td>{escape(str(row.get('Source', '')))}</td>
                <td><span class="source-status source-{status_class}">{escape(status)}</span></td>
                <td class="num">{escape(str(row.get('Records', 0)))}</td>
                <td>{escape(str(row.get('Details', '')))}</td>
            </tr>"""
        if source_status:
            source_table = f"""<div class="security-table-wrap"><table>
                <thead><tr><th>{bi('Source', 'Fuente')}</th><th>{bi('Status', 'Estado')}</th><th>{bi('Records', 'Registros')}</th><th>{bi('Details', 'Detalles')}</th></tr></thead>
                <tbody>{status_rows}</tbody>
            </table></div>"""
        else:
            source_table = f'<p class="empty-state">{bi("Security source status is unavailable.", "El estado de las fuentes de seguridad no está disponible.")}</p>'

        recommendations_table = render_table(recommendations, [
            ("recommendationSeverity", "Severity", "Severidad", False),
            ("recommendationName", "Recommendation", "Recomendación", False),
            ("recommendationState", "State", "Estado", False),
            ("affectedResourceId", "Affected Resource", "Recurso afectado", False),
        ], "No open Defender for Cloud recommendations.", "No hay recomendaciones abiertas de Defender for Cloud.")
        controls_table = render_table(controls, [
            ("controlName", "Secure Score Control", "Control de Secure Score", False),
            ("healthyResourceCount", "Healthy", "Saludables", True),
            ("unhealthyResourceCount", "Unhealthy", "No saludables", True),
            ("currentScore", "Current Score", "Puntaje actual", True),
            ("maxScore", "Max Score", "Puntaje máximo", True),
        ], "No Secure Score controls returned.", "No se devolvieron controles de Secure Score.")
        mcsb_table = render_table(mcsb, [
            ("complianceControl", "MCSB Control", "Control MCSB", False),
            ("assessmentName", "Assessment", "Evaluación", False),
            ("state", "State", "Estado", False),
            ("passedResources", "Passed", "Aprobados", True),
            ("failedResources", "Failed", "Con error", True),
        ], "No MCSB assessment data returned.", "No se devolvieron datos de evaluación MCSB.")
        plans_table = render_table(defender_plans, [
            ("planName", "Defender Plan", "Plan de Defender", False),
            ("pricingTier", "Tier", "Nivel", False),
            ("subPlan", "Sub-plan", "Subplan", False),
            ("subscriptionId", "Subscription", "Suscripción", False),
        ], "No Defender plan data returned.", "No se devolvieron datos de planes de Defender.")
        incidents_table = render_table(incidents, [
            ("Severity", "Severity", "Severidad", False),
            ("DisplayName", "Incident", "Incidente", False),
            ("Status", "Status", "Estado", False),
            ("AssignedTo", "Assigned To", "Asignado a", False),
            ("LastUpdateDateTime", "Last Updated", "Última actualización", False),
        ], "No incidents returned for the selected lookback period.", "No se devolvieron incidentes para el período seleccionado.")
        alerts_table = render_table(alerts, [
            ("Severity", "Severity", "Severidad", False),
            ("Title", "Alert", "Alerta", False),
            ("Status", "Status", "Estado", False),
            ("ServiceSource", "Service", "Servicio", False),
            ("LastUpdateDateTime", "Last Updated", "Última actualización", False),
        ], "No alerts returned for the selected lookback period.", "No se devolvieron alertas para el período seleccionado.")
        endpoint_recommendations_table = render_table(endpoint_recommendations, [
            ("RecommendationName", "Endpoint Recommendation", "Recomendación de Endpoint", False),
            ("ExposedMachinesCount", "Exposed Machines", "Equipos expuestos", True),
            ("SeverityScore", "Severity Score", "Puntaje de severidad", True),
            ("Status", "Status", "Estado", False),
        ], "No Defender for Endpoint recommendations returned.", "No se devolvieron recomendaciones de Defender for Endpoint.")
        vulnerabilities_table = render_table(vulnerabilities, [
            ("VulnerabilityId", "CVE", "CVE", False),
            ("Severity", "Severity", "Severidad", False),
            ("CvssV3", "CVSS v3", "CVSS v3", True),
            ("ExposedMachines", "Exposed Machines", "Equipos expuestos", True),
            ("PublicExploit", "Public Exploit", "Exploit público", False),
        ], "No Defender for Endpoint vulnerabilities returned.", "No se devolvieron vulnerabilidades de Defender for Endpoint.")

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow">MICROSOFT DEFENDER</div>
                <h2>{bi('Security posture and operations', 'Postura y operaciones de seguridad')}</h2>
                <p>{bi('CSPM, Microsoft Cloud Security Benchmark, Defender XDR, and endpoint exposure in one assessment.', 'CSPM, Microsoft Cloud Security Benchmark, Defender XDR y exposición de endpoints en una sola evaluación.')}</p>
            </div>
            <div class="security-score-mark">{secure_score_text}</div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title security-title">{bi('CSPM recommendations', 'Recomendaciones CSPM')}</div>
            {recommendations_table}
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Secure Score controls', 'Controles de Secure Score')}</h3>{controls_table}</div>
            <div class="security-panel"><h3>{bi('Microsoft Cloud Security Benchmark', 'Microsoft Cloud Security Benchmark')}</h3>{mcsb_table}</div>
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Defender for Cloud coverage', 'Cobertura de Defender for Cloud')}</h3>{plans_table}</div>
            <div class="security-panel"><h3>{bi('Data source coverage', 'Cobertura de fuentes de datos')}</h3>{source_table}</div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Defender XDR operations', 'Operaciones de Defender XDR')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Recent incidents', 'Incidentes recientes')}</h3>{incidents_table}</div>
                <div class="security-panel"><h3>{bi('Recent alerts', 'Alertas recientes')}</h3>{alerts_table}</div>
            </div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Defender for Endpoint exposure', 'Exposición de Defender for Endpoint')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Security recommendations', 'Recomendaciones de seguridad')}</h3>{endpoint_recommendations_table}</div>
                <div class="security-panel"><h3>{bi('Vulnerabilities', 'Vulnerabilidades')}</h3>{vulnerabilities_table}</div>
            </div>
        </div>"""

    def _render_governance_overview(self) -> str:
        """Surface GovViz data (RBAC, policy, locks, management groups) that isn't reflected in pillar scores."""
        gov = self.data.get("governance", {})
        stats = [
            (bi("Management Groups", "Grupos de Administraci\u00f3n"), self._count_real(gov.get("MgmtGroups", []))),
            (bi("Policy Assignments", "Asignaciones de Pol\u00edticas"), self._count_real(gov.get("PolicyAssignments", []))),
            (bi("Custom Policies", "Pol\u00edticas Personalizadas"), self._count_real(gov.get("CustomPolicies", []))),
            (bi("Role Assignments", "Asignaciones de Roles"), self._count_real(gov.get("RoleAssignments", []))),
            (bi("Custom Roles", "Roles Personalizados"), self._count_real(gov.get("CustomRoles", []))),
            (bi("Resource Locks", "Bloqueos de Recursos"), self._count_real(gov.get("Locks", []))),
        ]
        stat_boxes = "".join(
            f'<div class="stat-box"><div class="value">{count:,}</div><div class="label">{label}</div></div>'
            for label, count in stats
        )

        report_html = ""
        if self.base_dir:
            report_path = os.path.join(self.base_dir, "04_Governance", "AzureGovernance.html")
            if os.path.exists(report_path):
                link_text = bi(
                    "Open full Governance Visualizer report (Hierarchy Map, Tenant Summary, Scope Insights) in a new tab \u2197",
                    "Abrir el informe completo del Visualizador de Gobernanza (Mapa de Jerarqu\u00eda, Resumen del Tenant, Perspectivas de Alcance) en una pesta\u00f1a nueva \u2197"
                )
                report_html = f"""<div class="governance-embed">
            <div class="governance-embed-toolbar">
                <a href="../04_Governance/AzureGovernance.html" target="_blank">{link_text}</a>
            </div>
            <iframe src="../04_Governance/AzureGovernance.html" loading="lazy" onload="try{{const d=this.contentDocument;const panels=[...d.body.children];const fit=()=>{{const visible=panels.filter(el=>this.contentWindow.getComputedStyle(el).display!=='none');const bottom=Math.max(0,...visible.map(el=>el.offsetTop+el.offsetHeight));this.style.height=(bottom+2)+'px';}};fit();const observer=new ResizeObserver(fit);panels.forEach(el=>observer.observe(el));}}catch(e){{}}"></iframe>
        </div>"""
        if not report_html:
            note = bi(
                "Governance Visualizer HTML report not found (04_Governance/AzureGovernance.html) — run Invoke-AzureGovernanceViz-CloudShell.ps1 to generate it.",
                "No se encontr\u00f3 el informe HTML del Visualizador de Gobernanza (04_Governance/AzureGovernance.html); ejecute Invoke-AzureGovernanceViz-CloudShell.ps1 para generarlo."
            )
            report_html = f'<p class="breakdown-note">{note}</p>'

        return f"""<div class="stats-row">{stat_boxes}</div>
        {report_html}"""

    def _render_score_breakdown(self) -> str:
        """Render the 'why' behind each pillar score per the Advisor Score model."""
        b = self.engine.breakdowns
        sections = []
        pillar_label = lambda p: bi(p, PILLAR_ES.get(p, p))
        insufficient_note = bi(
            "Insufficient data — no discovered resource inventory (01_Discovery/ResourceSummary) to compute this pillar's score.",
            "Datos insuficientes: no hay inventario de recursos descubiertos (01_Discovery/ResourceSummary) para calcular el puntaje de este pilar."
        )

        def subcategory_card(pillar):
            info = b.get(pillar, {})
            if not info.get("available"):
                return f"""<div class="breakdown-card">
                <h4>{pillar_label(pillar)}</h4>
                <p class="breakdown-note">{insufficient_note}</p>
            </div>"""
            rows_html = "".join(
                f"<tr><td>{bi(r['subcategory'], SUBCATEGORY_ES.get(r['subcategory'], r['subcategory']))}</td><td class='num'>{r['weight']}</td>"
                f"<td class='num'>{r['healthy']}/{r['total']}</td><td class='num'>{r['pct']}%</td></tr>"
                for r in info["rows"]
            )
            return f"""<div class="breakdown-card">
                <h4>{pillar_label(pillar)} <span class="breakdown-score">{self.engine.scores[pillar]}/100</span></h4>
                <table>
                    <thead><tr><th>{bi('Subcategory', 'Subcategor\u00eda')}</th><th>{bi('Weight', 'Peso')}</th><th>{bi('Healthy/Total', 'Saludable/Total')}</th><th>{bi('Score', 'Puntaje')}</th></tr></thead>
                    <tbody>{rows_html}</tbody>
                </table>
            </div>"""

        sections.append(subcategory_card("Reliability"))
        sections.append(subcategory_card("Performance Efficiency"))
        sections.append(subcategory_card("Operational Excellence"))

        security = b.get("Security", {})
        if security.get("available"):
            findings_html = "".join(
                f"<tr><td>{bi(label_en, label_es)}</td><td class='num'>{count}</td></tr>"
                for label_en, label_es, count in security.get("findings", [])
            )
            note = bi(
                f"Microsoft Defender Secure Score average across {security.get('subscriptions_scored', 0)} subscription(s): <strong>{security.get('secure_score_pct')}%</strong>.",
                f"Promedio del Secure Score de Microsoft Defender en {security.get('subscriptions_scored', 0)} suscripci\u00f3n(es): <strong>{security.get('secure_score_pct')}%</strong>."
            )
            sections.append(f"""<div class="breakdown-card">
                <h4>{pillar_label('Security')} <span class="breakdown-score">{self.engine.scores['Security']}/100</span></h4>
                <p class="breakdown-note">{note}</p>
                <table>
                    <thead><tr><th>{bi('Finding Category', 'Categor\u00eda de Hallazgo')}</th><th>{bi('Count', 'Cantidad')}</th></tr></thead>
                    <tbody>{findings_html}</tbody>
                </table>
            </div>""")
        else:
            note = bi(
                "Insufficient data — no Defender Secure Score records found in 04_Governance/SecureScores.",
                "Datos insuficientes: no se encontraron registros de Secure Score de Defender en 04_Governance/SecureScores."
            )
            sections.append(f"""<div class="breakdown-card">
                <h4>{pillar_label('Security')}</h4>
                <p class="breakdown-note">{note}</p>
            </div>""")

        cost = b.get("Cost Optimization", {})
        if cost.get("available"):
            source_rows = []
            for source in cost.get("sources", []):
                name = source["name"]
                affected = min(source["affected"], cost["total"])
                source_healthy = max(0, cost["total"] - affected)
                source_score = round((source_healthy / cost["total"]) * 100)
                source_rows.append(
                    f"<tr><td>{bi(name, COST_SOURCE_ES.get(name, name))}</td>"
                    f"<td class='num'>{source['count']} / {affected}</td>"
                    f"<td class='num'>{source_healthy}/{cost['total']}</td>"
                    f"<td class='num'>{source_score}%</td></tr>"
                )
            combined_label = bi("Combined (deduplicated)", "Combinado (sin duplicados)")
            combined_findings = sum(source["count"] for source in cost.get("sources", []))
            source_rows.append(
                f"<tr><td><strong>{combined_label}</strong></td>"
                f"<td class='num'><strong>{combined_findings} / {cost['impacted']}</strong></td>"
                f"<td class='num'><strong>{cost['healthy']}/{cost['total']}</strong></td>"
                f"<td class='num'><strong>{self.engine.scores['Cost Optimization']}%</strong></td></tr>"
            )
            sources_html = "".join(source_rows)
            note = bi(
                "Findings/Affected shows finding count first and unique affected scopes or resources second. The combined row deduplicates identifiers.",
                "Hallazgos/Afectados muestra primero los hallazgos y luego los ámbitos o recursos únicos afectados. La fila combinada elimina duplicados."
            )
            sections.append(f"""<div class="breakdown-card cost-breakdown-card">
                <h4>{pillar_label('Cost Optimization')} <span class="breakdown-score">{self.engine.scores['Cost Optimization']}/100</span></h4>
                <p class="breakdown-note">{note}</p>
                <div class="breakdown-table-wrap">
                    <table>
                        <colgroup><col class="cost-source-col"><col class="cost-findings-col"><col><col></colgroup>
                        <thead><tr><th>{bi('Finding Source', 'Fuente del Hallazgo')}</th><th>{bi('Findings/Affected', 'Hallazgos/Afectados')}</th><th>{bi('Healthy/Total', 'Saludable/Total')}</th><th>{bi('Score', 'Puntaje')}</th></tr></thead>
                        <tbody>{sources_html}</tbody>
                    </table>
                </div>
            </div>""")
        else:
            note = bi(
                "Insufficient data — no discovered resource inventory to compute this pillar's score.",
                "Datos insuficientes: no hay inventario de recursos descubiertos para calcular el puntaje de este pilar."
            )
            sections.append(f"""<div class="breakdown-card">
                <h4>{pillar_label('Cost Optimization')}</h4>
                <p class="breakdown-note">{note}</p>
            </div>""")

        return "".join(sections)
    
    def generate(self) -> str:
        overall = self.engine.get_overall_score()
        scores = self.engine.scores
        score_breakdown_html = self._render_score_breakdown()
        governance_overview_html = self._render_governance_overview()
        security_view_html = self._render_security_view()
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
            color = DashboardConfig.SEVERITY_COLORS.get(action["severity"], "var(--cp-text-muted)")
            resources_html = ""
            if action.get("resources"):
                resources_html = "<ul class='resource-list'>" + "".join(f"<li>{r}</li>" for r in action["resources"] if r) + "</ul>"
            
            severity_badge = bi(action['severity'].upper(), SEVERITY_ES.get(action['severity'], action['severity'].upper()))
            pillar_tag = bi(action['pillar'], PILLAR_ES.get(action['pillar'], action['pillar']))
            title_text = bi(action['title'], action.get('title_es', action['title']))
            desc_text = bi(action['description'], action.get('description_es', action['description']))
            actions_html += f"""
            <div class="action-item" style="border-left: 4px solid {color};">
                <div class="action-header">
                    <span class="action-num">#{i}</span>
                    <span class="severity-badge" style="background:{color};">{severity_badge}</span>
                    <span class="pillar-tag">{pillar_tag}</span>
                    <span class="action-title">{title_text}</span>
                </div>
                <p class="action-desc">{desc_text}</p>
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
                <div class="pillar-name">{status_icon} {bi(pillar, PILLAR_ES.get(pillar, pillar))}</div>
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
<script>
    (() => {{
        const param = new URLSearchParams(window.location.search).get("scoutTheme");
        const theme =
            param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
        document.documentElement.setAttribute("data-theme", theme);
    }})();
</script>
<style>
:root {{
    color-scheme: light;
    --cp-bg: #f7f4ef;
    --cp-bg-elevated: #fcfbf8;
    --cp-surface: #ffffff;
    --cp-surface-soft: #f5f5f5;
    --cp-border: #dedede;
    --cp-border-strong: #919191;
    --cp-text: #242424;
    --cp-text-muted: #5c5c5c;
    --cp-text-soft: #6f6f6f;
    --cp-accent: #b11f4b;
    --cp-accent-hover: #9a1a41;
    --cp-accent-soft: rgba(177, 31, 75, 0.08);
    --cp-accent-fg: #ffffff;
    --cp-success: #16a34a;
    --cp-danger: #dc2626;
    --cp-warning: #f59e0b;
    --cp-link: #0078d4;
    --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.12);
    --cp-overlay: rgba(255, 255, 255, 0.8);
    --cp-panel: rgba(255, 255, 255, 0.86);
    --cp-panel-strong: rgba(255, 255, 255, 0.96);
    --cp-sheen: rgba(255, 255, 255, 0.55);
    --cp-highlight: rgba(177, 31, 75, 0.12);
}}
html[data-theme="dark"] {{
    color-scheme: dark;
    --cp-bg: #3d3b3a;
    --cp-bg-elevated: #343231;
    --cp-surface: #292929;
    --cp-surface-soft: #2e2e2e;
    --cp-border: #474747;
    --cp-border-strong: #5f5f5f;
    --cp-text: #dedede;
    --cp-text-muted: #919191;
    --cp-text-soft: #b0b0b0;
    --cp-accent: #fd8ea1;
    --cp-accent-hover: #fb7b91;
    --cp-accent-soft: rgba(253, 142, 161, 0.14);
    --cp-accent-fg: #1a1a1a;
    --cp-success: #4ade80;
    --cp-danger: #f87171;
    --cp-warning: #fbbf24;
    --cp-link: #4da6ff;
    --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.32);
    --cp-overlay: rgba(41, 41, 41, 0.88);
    --cp-panel: rgba(41, 41, 41, 0.72);
    --cp-panel-strong: rgba(41, 41, 41, 0.96);
    --cp-sheen: rgba(255, 255, 255, 0.04);
    --cp-highlight: rgba(253, 142, 161, 0.12);
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
        font-family: "Segoe UI", Aptos, Calibri, -apple-system, BlinkMacSystemFont, sans-serif;
        background: var(--cp-bg); color: var(--cp-text);
    line-height: 1.5;
}}
.dashboard {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}

/* Header */
.header {{
    background: var(--cp-accent);
    color: var(--cp-accent-fg); padding: 32px; border-radius: 10px;
    margin-bottom: 24px; position: relative;
}}
.header h1 {{ font-size: 28px; margin-bottom: 8px; }}
.header .subtitle {{ opacity: 0.85; font-size: 14px; }}

/* Dashboard views */
.view-tabs {{
    display: flex; width: fit-content; gap: 4px; padding: 4px;
    margin-bottom: 24px; background: var(--cp-surface-soft); border: 1px solid var(--cp-border); border-radius: 8px;
}}
.view-tab {{
    border: 0; background: transparent; color: var(--cp-text-muted); cursor: pointer;
    padding: 8px 18px; border-radius: 6px; font: inherit; font-size: 13px; font-weight: 600;
}}
.view-tab[aria-selected="true"] {{ background: var(--cp-surface); color: var(--cp-accent); }}
.dashboard-view[hidden] {{ display: none; }}

/* Dedicated security view */
.security-summary-band {{
    display: flex; justify-content: space-between; align-items: center; gap: 24px;
    background: var(--cp-bg-elevated); color: var(--cp-text); padding: 28px 32px; border-left: 6px solid var(--cp-accent);
    margin-bottom: 18px;
}}
.security-summary-band h2 {{ font-size: 24px; margin: 2px 0 6px; }}
.security-summary-band p {{ color: var(--cp-text-muted); font-size: 13px; }}
.security-eyebrow {{ color: var(--cp-warning); font-size: 10px; font-weight: 700; letter-spacing: 1px; }}
.security-score-mark {{ font-size: 44px; font-weight: 700; color: var(--cp-accent); white-space: nowrap; }}
.security-stats {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; margin-bottom: 24px; }}
.security-stat {{ background: var(--cp-surface); border-top: 3px solid var(--cp-border-strong); padding: 16px; min-width: 0; }}
.security-stat .value {{ font-size: 27px; font-weight: 700; }}
.security-stat .label {{ color: var(--cp-text-muted); font-size: 11px; margin-top: 4px; }}
.security-accent {{ color: var(--cp-link); }}
.security-danger {{ color: var(--cp-danger); }}
.security-warning {{ color: var(--cp-warning); }}
.security-title {{ border-bottom-color: var(--cp-accent); }}
.security-two-column {{ display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 18px; margin-bottom: 24px; }}
.security-panel {{ background: var(--cp-surface); padding: 18px; min-width: 0; border-top: 3px solid var(--cp-border-strong); }}
.security-panel h3 {{ font-size: 14px; margin-bottom: 12px; }}
.security-table-wrap {{ width: 100%; overflow-x: auto; }}
.security-table-wrap table {{ min-width: 540px; }}
.empty-state {{ color: var(--cp-text-muted); font-size: 12px; padding: 16px 0; }}
.source-status {{ display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 10px; font-weight: 700; }}
.source-available, .source-nodata {{ background: var(--cp-highlight); color: var(--cp-success); }}
.source-partial {{ background: var(--cp-accent-soft); color: var(--cp-warning); }}
.source-forbidden, .source-unavailable, .source-error {{ background: var(--cp-accent-soft); color: var(--cp-danger); }}
.source-skipped {{ background: var(--cp-surface-soft); color: var(--cp-text-muted); }}
@media (max-width: 1050px) {{ .security-stats {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }} }}
@media (max-width: 760px) {{
    .view-tabs {{ width: 100%; }}
    .view-tab {{ flex: 1; }}
    .security-summary-band {{ align-items: flex-start; padding: 22px 18px; }}
    .security-score-mark {{ font-size: 30px; }}
    .security-stats {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    .security-two-column {{ grid-template-columns: minmax(0, 1fr); }}
}}

/* Overall Score */
.score-section {{
    display: grid; grid-template-columns: 200px 1fr;
    gap: 24px; margin-bottom: 24px;
    background: var(--cp-surface); border-radius: 10px; padding: 24px;
    box-shadow: var(--cp-shadow);
}}
.overall-score {{
    display: flex; flex-direction: column; align-items: center; justify-content: center;
}}
.score-circle {{
    width: 140px; height: 140px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 48px; font-weight: 700;
    background: conic-gradient(
        {'var(--cp-success)' if overall >= 80 else 'var(--cp-warning)' if overall >= 50 else 'var(--cp-danger)'} {overall * 3.6}deg,
        var(--cp-border) {overall * 3.6}deg
    );
    position: relative;
}}
.score-circle::after {{
    content: '{overall}'; position: absolute;
    width: 110px; height: 110px; border-radius: 50%;
    background: var(--cp-surface); display: flex; align-items: center; justify-content: center;
    font-size: 42px; font-weight: 700;
    color: {'var(--cp-success)' if overall >= 80 else 'var(--cp-warning)' if overall >= 50 else 'var(--cp-danger)'};
}}
.score-label {{ margin-top: 8px; font-size: 14px; color: var(--cp-text-muted); font-weight: 600; }}

.pillars-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;
}}
.pillar-card {{
    background: var(--cp-surface-soft); border-radius: 8px; padding: 16px;
    border: 1px solid var(--cp-border);
}}
.pillar-score {{ font-size: 32px; font-weight: 700; }}
.pillar-bar {{ height: 6px; background: var(--cp-border); border-radius: 3px; margin: 8px 0; }}
.pillar-bar-fill {{ height: 100%; border-radius: 3px; transition: width 0.5s; }}
.pillar-name {{ font-size: 12px; font-weight: 600; color: var(--cp-text-muted); }}

@media (max-width: 700px) {{
    .dashboard {{ padding: 12px; }}
    .header {{ padding: 64px 16px 24px; }}
    .header h1 {{ padding-right: 0; }}
    .lang-toggle {{ top: 16px; right: 16px; }}
    .score-section {{ grid-template-columns: minmax(0, 1fr); padding: 16px; }}
    .pillars-grid {{ grid-template-columns: repeat(auto-fit, minmax(min(200px, 100%), 1fr)); min-width: 0; }}
    .pillar-card {{ min-width: 0; }}
}}

/* Score Breakdown */
.methodology-note {{
    background: var(--cp-accent-soft); border: 1px solid var(--cp-border); border-radius: 8px;
    padding: 12px 16px; font-size: 12px; color: var(--cp-text); margin-bottom: 16px;
}}
.methodology-note a {{ color: var(--cp-link); }}
.breakdown-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 16px;
}}
@media (min-width: 701px) {{ .breakdown-grid {{ grid-auto-rows: 1fr; }} }}
.breakdown-card {{
    background: var(--cp-surface); border-radius: 10px; padding: 16px;
    box-shadow: var(--cp-shadow); min-width: 0;
}}
.breakdown-card h4 {{ font-size: 13px; margin-bottom: 10px; display: flex; justify-content: space-between; }}
.breakdown-score {{ color: var(--cp-link); }}
.breakdown-note {{ font-size: 12px; color: var(--cp-text-muted); }}
.breakdown-table-wrap {{ width: 100%; overflow-x: auto; }}
.cost-breakdown-card table {{ table-layout: fixed; }}
.cost-breakdown-card th,
.cost-breakdown-card td {{ overflow-wrap: anywhere; white-space: normal; }}
.cost-source-col {{ width: 34%; }}
.cost-findings-col {{ width: 26%; }}

/* Governance Visualizer embed */
.governance-embed {{ margin-top: 16px; }}
.governance-embed-toolbar {{
    background: var(--cp-surface); border-radius: 8px 8px 0 0; padding: 10px 16px;
    border: 1px solid var(--cp-border); border-bottom: none; font-size: 12px;
}}
.governance-embed-toolbar a {{ color: var(--cp-link); text-decoration: none; font-weight: 600; }}
/* Fallback height covers most reports; onload JS below grows it to the real content height when the browser allows cross-frame access. */
.governance-embed iframe {{
    width: 100%; height: 1400px; border: 1px solid var(--cp-border); border-radius: 0 0 8px 8px;
    background: var(--cp-surface);
}}

/* Stats */
.stats-row {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px; margin-bottom: 24px;
}}
.stat-box {{
    background: var(--cp-surface); border-radius: 10px; padding: 20px;
    box-shadow: var(--cp-shadow); text-align: center;
}}
.stat-box .value {{ font-size: 36px; font-weight: 700; color: var(--cp-link); }}
.stat-box .label {{ font-size: 12px; color: var(--cp-text-muted); margin-top: 4px; }}

/* Action Items */
.section {{ margin-bottom: 24px; }}
.section-title {{
    font-size: 18px; font-weight: 700; margin-bottom: 16px;
    padding-bottom: 8px; border-bottom: 2px solid var(--cp-accent);
}}
.action-item {{
    background: var(--cp-surface); border-radius: 8px; padding: 16px; margin-bottom: 12px;
    box-shadow: var(--cp-shadow);
}}
.action-header {{ display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }}
.action-num {{ font-weight: 700; color: var(--cp-text-soft); font-size: 12px; }}
.severity-badge {{
    color: var(--cp-accent-fg); padding: 2px 8px; border-radius: 4px;
    font-size: 10px; font-weight: 700; text-transform: uppercase;
}}
.pillar-tag {{
    background: var(--cp-surface-soft); padding: 2px 8px; border-radius: 4px;
    font-size: 10px; color: var(--cp-text-muted);
}}
.action-title {{ font-weight: 600; font-size: 14px; }}
.action-desc {{ color: var(--cp-text-muted); margin-top: 8px; font-size: 13px; }}
.resource-list {{
    margin-top: 8px; padding-left: 20px;
    font-size: 11px; color: var(--cp-text-muted); max-height: 120px; overflow-y: auto;
}}
.resource-list li {{ margin: 2px 0; }}

/* Tables */
.data-grid {{
    display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px;
}}
@media (max-width: 900px) {{ .data-grid {{ grid-template-columns: 1fr; }} }}
.table-card {{
    background: var(--cp-surface); border-radius: 10px; padding: 20px;
    box-shadow: var(--cp-shadow); min-width: 0;
}}
.table-card h3 {{ font-size: 14px; margin-bottom: 12px; color: var(--cp-text); }}
.table-card table {{ table-layout: fixed; }}
.table-card th,
.table-card td {{ overflow-wrap: anywhere; white-space: normal; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
th {{ background: var(--cp-surface-soft); padding: 8px; text-align: left; font-weight: 600; border-bottom: 2px solid var(--cp-border); }}
td {{ padding: 6px 8px; border-bottom: 1px solid var(--cp-border); }}
tr:hover td {{ background: var(--cp-accent-soft); }}
.num {{ text-align: right; font-weight: 600; }}

/* Footer */
.footer {{
    text-align: center; padding: 24px; color: var(--cp-text-soft); font-size: 11px;
    border-top: 1px solid var(--cp-border); margin-top: 32px;
}}

/* Bilingual toggle (default: English shown, Spanish hidden) */
.i18n-es {{ display: none; }}
body.lang-es .i18n-en {{ display: none; }}
body.lang-es .i18n-es {{ display: inline; }}
.lang-toggle {{
    position: absolute; top: 20px; right: 24px;
    background: var(--cp-accent-hover); color: var(--cp-accent-fg);
    border: 1px solid var(--cp-accent-fg); border-radius: 6px;
    padding: 8px 14px; font-size: 12px; font-weight: 600; cursor: pointer;
}}
.lang-toggle:hover {{ background: var(--cp-accent); }}
</style>
</head>
<body>
<div class="dashboard">

<!-- Header -->
<div class="header">
    <button id="langToggleBtn" class="lang-toggle" onclick="toggleLang()" type="button">ES 🇪🇸</button>
    <h1>☁️ {bi("Azure WAF/CAF Workshop - Discovery Report", "Taller Azure WAF/CAF - Informe de Descubrimiento")}</h1>
    <div class="subtitle">{bi("Generated", "Generado")}: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {bi("Consolidated view across all discovery phases", "Vista consolidada de todas las fases de descubrimiento")}</div>
</div>

<div class="view-tabs" role="tablist" aria-label="Dashboard views">
    <button id="overviewTab" class="view-tab" role="tab" aria-selected="true" aria-controls="overviewView" onclick="setDashboardView('overview')" type="button">{bi('Overview', 'Resumen')}</button>
    <button id="securityTab" class="view-tab" role="tab" aria-selected="false" aria-controls="securityView" onclick="setDashboardView('security')" type="button">{bi('Security', 'Seguridad')}</button>
</div>

<main id="overviewView" class="dashboard-view" role="tabpanel" aria-labelledby="overviewTab">
<!-- Overall Score + Pillar Breakdown -->
<div class="score-section">
    <div class="overall-score">
        <div class="score-circle"></div>
        <div class="score-label">{bi("Overall WAF Score", "Puntaje General WAF")}</div>
    </div>
    <div class="pillars-grid">
        {pillar_cards}
    </div>
</div>

<!-- Stats Row -->
<div class="stats-row">
    <div class="stat-box"><div class="value">{total_subs}</div><div class="label">{bi("Subscriptions", "Suscripciones")}</div></div>
    <div class="stat-box"><div class="value">{total_resources:,}</div><div class="label">{bi("Total Resources", "Recursos Totales")}</div></div>
    <div class="stat-box"><div class="value">{total_actions}</div><div class="label">{bi("Action Items", "Elementos de Acci\u00f3n")}</div></div>
    <div class="stat-box"><div class="value" style="color:var(--cp-danger);">{critical_actions}</div><div class="label">{bi("Critical/High Priority", "Prioridad Cr\u00edtica/Alta")}</div></div>
</div>

<!-- Governance Visualizer (Hierarchy Map) -->
<div class="section">
    <div class="section-title">🏛️ {bi("Governance Visualizer", "Visualizador de Gobernanza")}</div>
    {governance_overview_html}
</div>

<!-- Score Methodology -->
<div class="section">
    <div class="section-title">📐 {bi("Score Methodology (Microsoft Advisor Score model)", "Metodolog\u00eda de Puntuaci\u00f3n (modelo de Azure Advisor Score)")}</div>
    <div class="methodology-note">
        {bi("Pillar scores follow the official", "Los puntajes de cada pilar siguen la f\u00f3rmula oficial de")} <a href="https://learn.microsoft.com/en-us/azure/advisor/advisor-score#calculation-of-advisor-score" target="_blank">Azure Advisor Score</a>
        {bi(
            "formulas and subcategory weights instead of an arbitrary point system. Security uses the Microsoft Defender Secure Score directly. "
            "Reliability, Performance, and Operational Excellence use Microsoft\u2019s published subcategory weights \u2014 recommendations are mapped to "
            "subcategories via keyword matching (Azure Resource Graph doesn\u2019t expose Advisor\u2019s internal subcategory tag), and the discovered "
            "resource inventory is used as the shared 'total applicable resources' pool. Cost uses a resource-count healthy ratio instead of "
            "retail-cost weighting, since this toolkit doesn\u2019t call the Azure Retail Prices API.",
            "y los pesos de subcategor\u00eda en lugar de un sistema de puntos arbitrario. Seguridad usa directamente el Secure Score de Microsoft Defender. "
            "Confiabilidad, Rendimiento y Excelencia Operativa usan los pesos de subcategor\u00eda publicados por Microsoft: las recomendaciones se "
            "asignan a subcategor\u00edas mediante coincidencia de palabras clave (Azure Resource Graph no expone la etiqueta interna de subcategor\u00eda de Advisor), "
            "y el inventario de recursos descubiertos se usa como el conjunto compartido de 'recursos totales aplicables'. Costo usa una proporci\u00f3n de "
            "recursos saludables en lugar de ponderaci\u00f3n por costo de lista, ya que esta herramienta no llama a la API de precios de Azure."
        )}
    </div>
    <div class="breakdown-grid">
        {score_breakdown_html}
    </div>
</div>

<!-- Action Items -->
<div class="section">
    <div class="section-title">🎯 {bi(f"Action Items ({total_actions})", f"Elementos de Acci\u00f3n ({total_actions})")}</div>
    {actions_html}
</div>

<!-- Data Summary Tables -->
<div class="data-grid">
    <div class="table-card">
        <h3>📦 {bi("Top Resource Types", "Principales Tipos de Recursos")}</h3>
        <table>
            <thead><tr><th>{bi("Resource Type", "Tipo de Recurso")}</th><th>{bi("Location", "Ubicaci\u00f3n")}</th><th>{bi("Count", "Cantidad")}</th></tr></thead>
            <tbody>{resource_rows}</tbody>
        </table>
    </div>
    <div class="table-card">
        <h3>💡 {bi("Advisor Summary", "Resumen de Advisor")}</h3>
        <table>
            <thead><tr><th>{bi("Category", "Categor\u00eda")}</th><th>{bi("Impact", "Impacto")}</th><th>{bi("Count", "Cantidad")}</th></tr></thead>
            <tbody>{advisor_rows}</tbody>
        </table>
    </div>
</div>
</main>

<main id="securityView" class="dashboard-view" role="tabpanel" aria-labelledby="securityTab" hidden>
    {security_view_html}
</main>

<!-- Footer -->
<div class="footer">
    {bi("Azure WAF/CAF Workshop Discovery Report | Generated by Azure Governance Discovery Toolkit", "Informe de Descubrimiento del Taller Azure WAF/CAF | Generado por Azure Governance Discovery Toolkit")}<br>
    {bi("This report consolidates findings from Resource Discovery, Azure Advisor, Metrics Analysis, and Governance Visualization.", "Este informe consolida los hallazgos de Resource Discovery, Azure Advisor, An\u00e1lisis de M\u00e9tricas y Visualizaci\u00f3n de Gobernanza.")}
</div>

</div>
<script>
function setLang(lang) {{
    document.body.classList.toggle('lang-es', lang === 'es');
    document.documentElement.lang = lang;
    var btn = document.getElementById('langToggleBtn');
    if (btn) {{ btn.textContent = lang === 'es' ? 'EN 🇬🇧' : 'ES 🇪🇸'; }}
    try {{ localStorage.setItem('wafDashboardLang', lang); }} catch (e) {{}}
}}
function toggleLang() {{
    setLang(document.body.classList.contains('lang-es') ? 'en' : 'es');
}}
function setDashboardView(view) {{
    var security = view === 'security';
    document.getElementById('overviewView').hidden = security;
    document.getElementById('securityView').hidden = !security;
    document.getElementById('overviewTab').setAttribute('aria-selected', String(!security));
    document.getElementById('securityTab').setAttribute('aria-selected', String(security));
    try {{ localStorage.setItem('wafDashboardView', view); }} catch (e) {{}}
}}
(function() {{
    var saved = 'en';
    try {{ saved = localStorage.getItem('wafDashboardLang') || 'en'; }} catch (e) {{}}
    setLang(saved);
    var savedView = 'overview';
    try {{ savedView = localStorage.getItem('wafDashboardView') || 'overview'; }} catch (e) {{}}
    setDashboardView(savedView === 'security' ? 'security' : 'overview');
}})();
</script>
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
    generator = DashboardGenerator(engine, data, base_dir)
    html = generator.generate()
    
    output_path = os.path.join(base_dir, "06_Dashboard", "WAF_Dashboard.html")
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
