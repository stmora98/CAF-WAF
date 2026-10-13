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
    pip install openpyxl xlrd  (usually pre-installed in Cloud Shell)
"""

import sys
import os
import json
import re
import hashlib
import base64
import math
from collections import defaultdict
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

try:
    import xlrd
except ImportError:
    xlrd = None


# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

class DashboardConfig:
    SEVERITY_COLORS = {
     "critical": "var(--cp-critical)",
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
# TRILINGUAL (EN/ES/PT) SUPPORT
# All three language variants are always rendered; CSS toggles visibility via the
# .i18n-en/.i18n-es/.i18n-pt classes based on a "lang-es"/"lang-pt" class on <body>
# (see the embedded <script> in DashboardGenerator.generate()).
# ═══════════════════════════════════════════════════════════════════════════════

def bi(en: str, es: str, pt: Optional[str] = None) -> str:
    """Wrap a piece of text in all three language variants for the client-side toggle."""
    pt = pt if pt is not None else en
    return f'<span class="i18n-en">{en}</span><span class="i18n-es">{es}</span><span class="i18n-pt">{pt}</span>'


def it(en: str, es: str, pt: Optional[str] = None) -> str:
    """Azure-Portal-style info tooltip: a small 'i' circle placed next to a section title
    that reveals explanatory text only on hover/focus (keyboard-accessible via tabindex)."""
    return f'<span class="info-tip" tabindex="0"><span aria-hidden="true">i</span><span class="info-tip-bubble">{bi(en, es, pt)}</span></span>'


def _cell_html(value) -> str:
    """Escape a table cell value, except values already produced by bi() (pre-safe markup)."""
    text = str(value)
    if text.startswith('<span class="i18n-'):
        return text
    return escape(text)


PILLAR_ES = {
    "Reliability": "Confiabilidad",
    "Security": "Seguridad",
    "Cost Optimization": "Optimización de Costos",
    "Operational Excellence": "Excelencia Operativa",
    "Performance Efficiency": "Eficiencia de Rendimiento",
}

PILLAR_PT = {
    "Reliability": "Confiabilidade",
    "Security": "Segurança",
    "Cost Optimization": "Otimização de Custos",
    "Operational Excellence": "Excelência Operacional",
    "Performance Efficiency": "Eficiência de Desempenho",
}

# CSS custom-property names (no "var()" wrapper) matching DashboardConfig.PILLAR_COLORS,
# used by Chart.js charts which must resolve actual color values via getComputedStyle().
PILLAR_COLOR_VARS = {
    "Reliability": "--cp-link",
    "Security": "--cp-danger",
    "Cost Optimization": "--cp-success",
    "Operational Excellence": "--cp-accent",
    "Performance Efficiency": "--cp-warning",
}

# Trilingual labels + CSS var name for each severity bucket, reused by every Chart.js
# severity chart across Overview/Security/Reliability/etc.
SEVERITY_LABELS = {
    "critical": {"en": "Critical", "es": "Cr\u00edtico", "pt": "Cr\u00edtico"},
    "high": {"en": "High", "es": "Alto", "pt": "Alto"},
    "medium": {"en": "Medium", "es": "Medio", "pt": "M\u00e9dio"},
    "low": {"en": "Low", "es": "Bajo", "pt": "Baixo"},
    "info": {"en": "Info", "es": "Info", "pt": "Info"},
    "informational": {"en": "Info", "es": "Info", "pt": "Info"},
    "unknown": {"en": "Unknown", "es": "Desconocido", "pt": "Desconhecido"},
}

def _bare_var(css_color: Optional[str]) -> str:
    """Strip a 'var(--x)' wrapper down to the bare '--x' custom-property name for Chart.js
    (which must resolve colors via getComputedStyle(), not the CSS var() function)."""
    if not css_color:
        return "--cp-text-muted"
    return css_color.replace("var(", "").rstrip(")")


SEVERITY_ES = {
    "critical": "CRÍTICO",
    "high": "ALTO",
    "medium": "MEDIO",
    "low": "BAJO",
    "info": "INFO",
}

SEVERITY_PT = {
    "critical": "CRÍTICO",
    "high": "ALTO",
    "medium": "MÉDIO",
    "low": "BAIXO",
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

SUBCATEGORY_PT = {
    "Zone Resiliency": "Resiliência de Zona",
    "Regional Resiliency": "Resiliência Regional",
    "Data Protection and Recovery": "Proteção e Recuperação de Dados",
    "Governance and Compliance": "Governança e Conformidade",
    "Scalability": "Escalabilidade",
    "Monitoring and Alerting": "Monitoramento e Alertas",
    "Service Upgrade and Retirement": "Atualização e Descontinuação de Serviços",
    "Other": "Outro",
    "Compute Optimization": "Otimização de Computação",
    "Storage Optimization": "Otimização de Armazenamento",
    "Network Optimization": "Otimização de Rede",
    "Data Performance": "Desempenho de Dados",
    "Efficiency Optimization": "Otimização de Eficiência",
    "Failure Mitigation": "Mitigação de Falhas",
    "Safe and Secure Deployment": "Implantação Segura",
}


COST_SOURCE_ES = {
    "Orphaned/unused resources": "Recursos huérfanos/no utilizados",
    "Deallocated VMs": "VMs desasignadas",
    "Idle/underutilized VMs": "VMs inactivas/subutilizadas",
    "Advisor Cost recommendations": "Recomendaciones de costos de Advisor",
}

COST_SOURCE_PT = {
    "Orphaned/unused resources": "Recursos órfãos/não utilizados",
    "Deallocated VMs": "VMs desalocadas",
    "Idle/underutilized VMs": "VMs ociosas/subutilizadas",
    "Advisor Cost recommendations": "Recomendações de custo do Advisor",
}

# Shared Portuguese translations for recurring table column headers / labels
# used across the many render_table() closures throughout the file.
LABEL_PT = {
    "Severity": "Severidade",
    "Recommendation": "Recomendação",
    "State": "Estado",
    "Affected Resource": "Recurso Afetado",
    "Secure Score Control": "Controle do Secure Score",
    "Healthy": "Saudáveis",
    "Unhealthy": "Não saudáveis",
    "Current Score": "Pontuação atual",
    "Max Score": "Pontuação máxima",
    "MCSB Control": "Controle MCSB",
    "Assessment": "Avaliação",
    "Passed": "Aprovados",
    "Failed": "Com falha",
    "Defender Plan": "Plano do Defender",
    "Tier": "Nível",
    "Sub-plan": "Subplano",
    "Subscription": "Assinatura",
    "Incident": "Incidente",
    "Status": "Status",
    "Assigned To": "Atribuído a",
    "Last Updated": "Última atualização",
    "Alert": "Alerta",
    "Service": "Serviço",
    "Endpoint Recommendation": "Recomendação de Endpoint",
    "Exposed Machines": "Máquinas expostas",
    "Severity Score": "Pontuação de severidade",
    "CVE": "CVE",
    "Public Exploit": "Exploit público",
    "Key Vault": "Key Vault",
    "Type": "Tipo",
    "Item": "Item",
    "Days Until Expiry": "Dias até expirar",
    "Expires On": "Expira em",
    "Application": "Aplicativo",
    "Guest User": "Usuário convidado",
    "Email": "E-mail",
    "Created": "Criado",
    "Enabled": "Habilitado",
    "Name": "Nome",
    "Id": "Id",
    "Parent": "Pai",
    "Enforcement": "Aplicação",
    "Identity": "Identidade",
    "Scope": "Escopo",
    "Effect": "Efeito",
    "Category": "Categoria",
    "Policy Assignment": "Atribuição de Política",
    "Non-Compliant": "Não Conformes",
    "Principal Type": "Tipo de Principal",
    "Role": "Função",
    "Role Name": "Nome da Função",
    "Description": "Descrição",
    "Plan": "Plano",
    "Score %": "Pontuação %",
    "Resource Type": "Tipo de Recurso",
    "Location": "Localização",
    "Count": "Quantidade",
    "Resource Group": "Grupo de Recursos",
    "Detail": "Detalhe",
    "VNet": "VNet",
    "Address Space": "Espaço de Endereço",
    "Subnets": "Sub-redes",
    "Peerings": "Emparelhamentos",
    "Level": "Nível",
    "Impact": "Impacto",
    "Problem": "Problema",
    "Solution": "Solução",
    "Annual Savings": "Economia Anual",
    "SKU": "SKU",
    "Term": "Prazo",
    "Qty": "Qtd.",
    "Net Savings/yr": "Economia Líquida/ano",
    "VM": "VM",
    "Size": "Tamanho",
    "Avg CPU %": "CPU Média %",
    "Max CPU %": "CPU Máx. %",
    "Database": "Banco de Dados",
    "Metric": "Métrica",
    "Avg Usage %": "Uso Médio %",
    "App Plan": "Plano de App",
    "Workers": "Workers",
    "Storage Account": "Conta de Armazenamento",
    "Avg Daily Txns": "Trans. Diárias Médias",
    "Size (GB)": "Tamanho (GB)",
    "Month": "Mês",
    "Cost": "Custo",
    "6-mo Cost": "Custo 6 meses",
    "Open Findings": "Descobertas Abertas",
    "Top Severity": "Severidade Máxima",
    "Top Issue": "Problema Principal",
    "Check": "Verificação",
    "Resource": "Recurso",
    "Power State": "Estado de Energia",
    "Backup Vault": "Vault de Backup",
    "Workload": "Carga de Trabalho",
    "Protection Status": "Estado de Proteção",
    "Last Backup Status": "Último Status de Backup",
    "Last Backup Time": "Última Hora de Backup",
    "Policy": "Política",
    "Event": "Evento",
    "Impact Start": "Início do Impacto",
    "Mitigated": "Mitigado",
    "Gap": "Lacuna",
}

EMPTY_STATE_PT = {
    "": "",
    "All VMs are deployed across Availability Zones.": "Todas as VMs estão implantadas em Zonas de Disponibilidade.",
    "All VNets have DDoS Protection enabled.": "Todas as VNets têm a Proteção contra DDoS habilitada.",
    "All critical resources have diagnostic settings configured.": "Todos os recursos críticos têm configurações de diagnóstico definidas.",
    "All resources are tagged.": "Todos os recursos têm tags.",
    "No Advisor cost recommendations found.": "Nenhuma recomendação de custo do Advisor encontrada.",
    "No Advisor operational excellence recommendations found.": "Nenhuma recomendação de excelência operacional do Advisor encontrada.",
    "No Advisor performance recommendations found.": "Nenhuma recomendação de desempenho do Advisor encontrada.",
    "No Advisor reliability recommendations found.": "Nenhuma recomendação de confiabilidade do Advisor encontrada.",
    "No App Service Plan right-sizing data found.": "Nenhum dado de dimensionamento de Planos do App Service encontrado.",
    "No Azure Resource Health / Service Health events found.": "Nenhum evento do Azure Resource Health / Service Health encontrado.",
    "No Backup/Recovery Services vaults found.": "Nenhum cofre de Backup/Recovery Services encontrado.",
    "No Defender for Cloud plans found.": "Nenhum plano do Defender for Cloud encontrado.",
    "No Defender for Endpoint recommendations returned.": "Nenhuma recomendação do Defender for Endpoint retornada.",
    "No Defender for Endpoint vulnerabilities returned.": "Nenhuma vulnerabilidade do Defender for Endpoint retornada.",
    "No Defender plan data returned.": "Nenhum dado de plano do Defender retornado.",
    "No MCSB assessment data returned.": "Nenhum dado de avaliação MCSB retornado.",
    "No SQL right-sizing data found.": "Nenhum dado de dimensionamento de SQL encontrado.",
    "No Secure Score controls returned.": "Nenhum controle de Secure Score retornado.",
    "No Secure Score data found.": "Nenhum dado de Secure Score encontrado.",
    "No VM right-sizing data found.": "Nenhum dado de dimensionamento de VMs encontrado.",
    "No actual cost data found (requires Cost Management Reader).": "Nenhum dado de custo real encontrado (requer a função Cost Management Reader).",
    "No alerts returned for the selected lookback period.": "Nenhum alerta retornado para o período selecionado.",
    "No backup-protected items found (vaults exist but nothing is protected, or Backup Reader access is missing).": "Nenhum item protegido por backup encontrado (existem cofres, mas nada está protegido, ou falta acesso de Backup Reader).",
    "No custom policy definitions found.": "Nenhuma definição de política personalizada encontrada.",
    "No custom role definitions found.": "Nenhuma definição de função personalizada encontrada.",
    "No deallocated VMs found.": "Nenhuma VM desalocada encontrada.",
    "No expiring Key Vault secrets or certificates found (or Key Vault Get permission unavailable).": "Nenhum segredo ou certificado do Key Vault próximo da expiração encontrado (ou a permissão Get do Key Vault não está disponível).",
    "No expiring app registration credentials found.": "Nenhuma credencial de registro de aplicativo próxima da expiração encontrada.",
    "No extended FinOps findings.": "Nenhuma descoberta de FinOps estendida.",
    "No guest (B2B) users found.": "Nenhum usuário convidado (B2B) encontrado.",
    "No incidents returned for the selected lookback period.": "Nenhum incidente retornado para o período selecionado.",
    "No management group data found.": "Nenhum dado de grupo de gerenciamento encontrado.",
    "No non-compliant resources found.": "Nenhum recurso não conforme encontrado.",
    "No open Defender for Cloud recommendations.": "Nenhuma recomendação aberta do Defender for Cloud.",
    "No orphaned resources found.": "Nenhum recurso órfão encontrado.",
    "No policy assignments found.": "Nenhuma atribuição de política encontrada.",
    "No policy exemptions found.": "Nenhuma isenção de política encontrada.",
    "No reservation/savings plan opportunities found.": "Nenhuma oportunidade de Reserva/Plano de Economia encontrada.",
    "No resource locks configured.": "Nenhum bloqueio de recurso configurado.",
    "No resource locks found.": "Nenhum bloqueio de recurso encontrado.",
    "No resource summary found.": "Nenhum resumo de recursos encontrado.",
    "No role assignments found.": "Nenhuma atribuição de função encontrada.",
    "No saturated SQL databases found.": "Nenhum banco de dados SQL saturado encontrado.",
    "No saturated VMs found.": "Nenhuma VM saturada encontrada.",
    "No storage activity data found.": "Nenhum dado de atividade de armazenamento encontrado.",
    "No unattached disks found.": "Nenhum disco não anexado encontrado.",
    "No virtual networks found.": "Nenhuma rede virtual encontrada.",
}

def bi_empty(en: str, es: str) -> str:
    """Render a translated empty-state message, falling back to English for new messages."""
    return bi(en, es, EMPTY_STATE_PT.get(en, en))

_CHART_TABLE_BLOCK_UID = 0

def chart_table_block(chart_html: str, table_html: str) -> str:
    """Wrap a Chart.js chart + its underlying table in a 'Chart | Table' toggle (chart shown
    by default, table available as an unmodified drill-down) - reusable across every
    table-to-chart conversion so the original data is never actually removed."""
    return f"""<div class="chart-table-block">
        <div class="chart-table-toggle" role="tablist">
            <button type="button" class="ctt-btn is-active" data-view="chart">{bi('Chart', 'Gr\u00e1fico', 'Gr\u00e1fico')}</button>
            <button type="button" class="ctt-btn" data-view="table">{bi('Table', 'Tabla', 'Tabela')}</button>
        </div>
        <div class="ctt-panel ctt-chart">{chart_html}</div>
        <div class="ctt-panel ctt-table" hidden>{table_html}</div>
    </div>"""

def _resource_link_attr(row: dict) -> str:
    """Return a data-resource-id (+ title) HTML attribute string when the row carries an ARM resource id, so the row can be double-clicked open in the Azure Portal."""
    resource_id = row.get("id") or row.get("Id") or row.get("ResourceId")
    if not resource_id or not str(resource_id).startswith("/"):
        return ""  # not a full ARM path (e.g. a bare GUID) - can't build a portal link from it
    # Plain text only - native title attributes can't hold bi()'s per-language <span> markup.
    tip = "Double-click to open in Azure Portal / Doble clic para abrir en Azure Portal / Clique duas vezes para abrir no Portal do Azure"
    return f' data-resource-id="{escape(str(resource_id))}" title="{escape(tip)}"'
# ═══════════════════════════════════════════════════════════════════════════════
# DATA READER
# ═══════════════════════════════════════════════════════════════════════════════

_SUBSCRIPTION_ID_RE = re.compile(r"/subscriptions/([0-9a-fA-F-]{36})")


def subscription_name_map(data: dict) -> dict:
    """Lowercased subscription id -> display name, from the Discovery/Governance Subscriptions sheets."""
    names = {}
    for phase in ("discovery", "governance"):
        for row in data.get(phase, {}).get("Subscriptions", []):
            sub_id = str(row.get("subscriptionId") or row.get("SubscriptionId") or "").lower()
            if sub_id and row.get("name"):
                names.setdefault(sub_id, str(row["name"]))
    return names


def resolve_subscription_name(row: dict, names: dict) -> str:
    """Display name of the subscription a workbook row belongs to ('' when it can't be determined).
    Rows may carry a name, a GUID, a separate subscriptionId, or only an ARM resource id."""
    current = str(row.get("Subscription") or "").strip()
    if current and current.lower() not in names:
        return current
    sub_id = current.lower() or str(row.get("subscriptionId") or row.get("SubscriptionId") or "").lower()
    if not sub_id:
        for key in ("ResourceId", "id", "resourceId"):
            match = _SUBSCRIPTION_ID_RE.search(str(row.get(key) or ""))
            if match:
                sub_id = match.group(1).lower()
                break
    return names.get(sub_id, "")


def attach_subscription_names(data: dict) -> dict:
    """Give every row that belongs to a known subscription a friendly `Subscription` value, so
    tables show names (not GUIDs) and the dashboard-wide subscription filter can match them."""
    names = subscription_name_map(data)
    if not names:
        return data
    for sheets in data.values():
        for rows in sheets.values():
            for row in rows:
                name = resolve_subscription_name(row, names)
                if name:
                    row["Subscription"] = name
    return data


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
            legacy = ExcelReader._read_legacy_workbook(path)
            if legacy is not None:
                print(f"  ! Read legacy Excel workbook with .xlsx extension: {path}")
                return legacy
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
    def _read_legacy_workbook(path: str) -> Optional[dict]:
        """Read an Excel 97-2003 workbook whose filename may incorrectly end in .xlsx."""
        if xlrd is None:
            return None
        try:
            workbook = xlrd.open_workbook(path, on_demand=True)
        except (OSError, xlrd.biffh.XLRDError):
            return None

        result = {}
        try:
            for sheet_name in workbook.sheet_names():
                sheet = workbook.sheet_by_name(sheet_name)
                if sheet.nrows < 2:
                    result[sheet_name] = []
                    continue
                headers = [
                    str(sheet.cell_value(0, i)) if sheet.cell_value(0, i) else f"col_{i}"
                    for i in range(sheet.ncols)
                ]
                result[sheet_name] = [
                    {
                        headers[i]: sheet.cell_value(row_index, i)
                        if sheet.cell_value(row_index, i) is not None else ""
                        for i in range(sheet.ncols)
                    }
                    for row_index in range(1, sheet.nrows)
                ]
        finally:
            workbook.release_resources()
        return result
    
    @staticmethod
    def read_all(base_dir: str) -> dict:
        """Read all Excel files from the workshop output directory."""
        data = {
            "discovery": {},
            "advisor": {},
            "finops": {},
            "metrics": {},
            "governance": {},
            "security": {},
            "checklists": {}
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

        # FinOps extended (actual cost, reservation utilization, extra recommendations)
        path = ExcelReader.find_workbook(base_dir, "02_Advisor", "AzureFinOps")
        if path:
            data["finops"] = ExcelReader.read_workbook(path)
            print(f"  ✓ FinOps: {sum(len(v) for v in data['finops'].values())} records across {len(data['finops'])} sheets")

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

        # Checklists (Azure/review-checklists community WAF ARG compliance scan)
        path = ExcelReader.find_workbook(base_dir, "06_Checklists", "AzureChecklists")
        if path:
            data["checklists"] = ExcelReader.read_workbook(path)
            print(f"  ✓ Checklists: {sum(len(v) for v in data['checklists'].values())} records across {len(data['checklists'])} sheets")
        
        return attach_subscription_names(data)


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
            name = str(rec.get("impactedResource") or rec.get("name") or rec.get("Name") or id(rec))
            key = f"{rec.get('Subscription') or ''}{KEY_SEPARATOR}{name}"
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


KEY_SEPARATOR = "\x1f"


def round_half_up(value: float) -> int:
    """Rounds .5 up; the page script mirrors this (Python's round() would round half to even)."""
    return int(math.floor(value + 0.5))


class ScoreModel:
    """Pillar scores computed from per-subscription counts, so the dashboard can recompute them in
    the browser for any subscription selection. `selected` is a set of subscription names (None =
    all). Rows whose subscription is unknown are stored under "" and belong to every selection."""

    @staticmethod
    def _total(by_subscription: dict, selected) -> int:
        return sum(
            value for sub, value in by_subscription.items()
            if selected is None or sub == "" or sub in selected
        )

    @staticmethod
    def group_keys(keys) -> dict:
        """Counts subscription-qualified keys ("<subscription>\x1f<name>") per subscription."""
        counts = {}
        for key in keys:
            key = str(key)
            sub = key.split(KEY_SEPARATOR, 1)[0] if KEY_SEPARATOR in key else ""
            counts[sub] = counts.get(sub, 0) + 1
        return counts

    @staticmethod
    def advisor_spec(weights: dict, buckets: dict, pool_by_subscription: dict, forced_zero=()) -> dict:
        return {
            "kind": "advisor",
            "weights": dict(weights),
            "forcedZero": sorted(forced_zero),
            "pool": dict(pool_by_subscription),
            "impacted": {subcategory: ScoreModel.group_keys(keys) for subcategory, keys in buckets.items()},
        }

    @staticmethod
    def pillar(spec: dict, selected=None) -> dict:
        kind = spec["kind"]
        if kind == "advisor":
            pool = ScoreModel._total(spec["pool"], selected)
            if pool <= 0:
                return {"available": False, "score": None, "rows": [], "pool": pool}
            rows, weighted, weight_total = [], 0.0, 0
            for subcategory, weight in spec["weights"].items():
                if subcategory in spec["forcedZero"]:
                    impacted = pool
                else:
                    impacted = min(ScoreModel._total(spec["impacted"].get(subcategory, {}), selected), pool)
                healthy = max(0, pool - impacted)
                pct = healthy / pool * 100
                weighted += pct * weight
                weight_total += weight
                rows.append({"subcategory": subcategory, "weight": weight, "healthy": healthy,
                             "total": pool, "pct": round_half_up(pct * 10) / 10})
            return {"available": True, "score": round_half_up(weighted / weight_total), "rows": rows, "pool": pool}
        if kind == "ratio":
            pool = ScoreModel._total(spec["pool"], selected)
            sources = [
                {"name": source["name"], "unit": source["unit"],
                 "count": ScoreModel._total(source["count"], selected),
                 "affected": ScoreModel._total(source["affected"], selected)}
                for source in spec["sources"]
            ]
            if pool <= 0:
                return {"available": False, "score": None, "sources": sources, "impacted": 0, "healthy": 0, "total": pool}
            impacted = min(ScoreModel._total(spec["impacted"], selected), pool)
            healthy = max(0, pool - impacted)
            return {"available": True, "score": round_half_up(healthy / pool * 100), "sources": sources,
                    "impacted": impacted, "healthy": healthy, "total": pool}
        scored = [pct for sub, pct in spec["scores"] if selected is None or sub == "" or sub in selected]
        if not scored:
            return {"available": False, "score": None, "pct": None, "scored": 0}
        average = sum(scored) / len(scored)
        return {"available": True, "score": round_half_up(average), "pct": round_half_up(average * 10) / 10, "scored": len(scored)}


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
        self._subscription_names = None
        self.score_specs = {}

    def _sub_name_of(self, row) -> str:
        """Subscription display name of a workbook row ('' when unknown)."""
        if self._subscription_names is None:
            self._subscription_names = subscription_name_map(self.data)
        return resolve_subscription_name(row, self._subscription_names) if isinstance(row, dict) else ""

    def _qkey(self, row, name) -> str:
        """Impacted-resource key qualified by subscription, so per-subscription scores can be derived."""
        return f"{self._sub_name_of(row)}{KEY_SEPARATOR}{name}"

    def _pool_by_subscription(self) -> dict:
        """Assessed-resource pool per subscription (Governance ResourceSummary carries a subscription);
        falls back to the Discovery total, which then belongs to every subscription selection."""
        per_subscription = {}
        for row in self._data_rows(self.data.get("governance", {}).get("ResourceSummary", [])):
            sub = self._sub_name_of(row)
            per_subscription[sub] = per_subscription.get(sub, 0) + int(row.get("Count", 0) or 0)
        if sum(per_subscription.values()) > 0:
            return per_subscription
        return {"": self._resource_pool()}

    def _cost_source(self, name: str, unit: str, keys, count_rows=None) -> dict:
        affected = ScoreModel.group_keys(keys)
        count = dict(affected)
        if count_rows is not None:
            count = {}
            for row in count_rows:
                sub = self._sub_name_of(row)
                count[sub] = count.get(sub, 0) + 1
        return {"name": name, "unit": unit, "count": count, "affected": affected}

    def score_model_payload(self) -> dict:
        """Everything the page script needs to recompute every pillar score for a subscription selection."""
        subscriptions = sorted(set((self._subscription_names or subscription_name_map(self.data)).values()))
        return {"subscriptions": subscriptions, "pillars": self.score_specs}

    def _apply_advisor_spec(self, pillar: str, spec: dict) -> None:
        result = ScoreModel.pillar(spec)
        self.score_specs[pillar] = spec
        self.scores[pillar] = result["score"] if result["available"] else 0
        self.breakdowns[pillar] = {"available": result["available"], "rows": result["rows"], "pool": result["pool"]}

    def _subscription_counts(self, rows) -> dict:
        """Subscription name -> number of the rows an action item is based on. Returns {} (the item
        stays visible under every subscription filter) when any row's subscription is unknown."""
        if self._subscription_names is None:
            self._subscription_names = subscription_name_map(self.data)
        counts = {}
        for row in rows:
            name = resolve_subscription_name(row, self._subscription_names) if isinstance(row, dict) else ""
            if not name:
                return {}
            counts[name] = counts.get(name, 0) + 1
        return dict(sorted(counts.items()))

    def _resource_pool(self) -> int:
        """Total discovered resources, used as the shared 'assessed resource' pool."""
        summary = self.data.get("discovery", {}).get("ResourceSummary", [])
        return sum(int(r.get("Count", 0) or 0) for r in summary)

    @staticmethod
    def _data_rows(rows: list) -> list:
        """Exclude sentinel rows emitted when a workbook query returns no data (e.g. Result:
        'No resources found'/'No recommendations' — the sentinel text varies by export script,
        so any truthy Result value is treated as empty, matching the view-layer real_rows() filter)."""
        return [
            row for row in rows
            if not row.get("Result", row.get("result", ""))
            and str(row.get("DataStatus", "")).strip().lower() != "nodata"
        ]

    def _expand_advisor_items(self, recs: list, pillar: str, savings_field: str = None,
                              savings_currency_field: str = "savingsCurrency") -> None:
        """Turn raw Advisor rows into one action item per distinct problem (grounded in the
        actual problem/solution text and impacted resources), instead of a single opaque count."""
        recs = self._data_rows(recs)
        if not recs:
            return
        groups = {}
        for r in recs:
            problem = str(r.get("problem", "")).strip() or f"{pillar} recommendation"
            groups.setdefault(problem, []).append(r)
        severity_rank = {"critical": 3, "high": 2, "medium": 1, "low": 0}
        for problem, rows in groups.items():
            impacts = [str(r.get("impact", "")).strip().lower() for r in rows]
            severity = max((i for i in impacts if i in severity_rank), key=lambda i: severity_rank[i], default="medium")
            solution = next((str(r.get("solution", "")).strip() for r in rows if r.get("solution")), "")
            resources = [str(r.get("impactedResource", r.get("impactedValue", ""))) for r in rows
                         if r.get("impactedResource") or r.get("impactedValue")]
            count = len(rows)
            title = f"{count} resource(s) affected: {problem}" if count > 1 else problem
            desc = solution or "Azure Advisor identified this issue for the impacted resource(s)."
            desc_es = solution or "Azure Advisor identificó este problema para los recursos afectados."
            desc_pt = solution or "O Azure Advisor identificou este problema para os recursos afetados."
            if savings_field:
                total = sum(float(r.get(savings_field, 0) or 0) for r in rows)
                currency = next((r.get(savings_currency_field) for r in rows if r.get(savings_currency_field)), "USD")
                if total > 0:
                    desc += f" Estimated annual savings: ~{total:,.0f} {currency}."
                    desc_es += f" Ahorro anual estimado: ~{total:,.0f} {currency}."
                    desc_pt += f" Economia anual estimada: ~{total:,.0f} {currency}."
            self.action_items.append({
                "pillar": pillar,
                "severity": severity,
                "source": "Azure Advisor",
                "title": title[:140],
                "title_es": title[:140],
                "title_pt": title[:140],
                "description": desc,
                "description_es": desc_es,
                "description_pt": desc_pt,
                "subscriptions": self._subscription_counts(rows),
                "resources": resources[:10]
            })

    def analyze(self):
        """Run all analysis passes."""
        analysis_passes = (
            (self._analyze_reliability, "Resource inventory and reliability analysis"),
            (self._analyze_security, "Security posture assessment"),
            (self._analyze_cost, "Cost and utilization analysis"),
            (self._analyze_operational_excellence, "Governance and monitoring assessment"),
            (self._analyze_performance, "Azure Monitor metrics analysis"),
            (self._analyze_checklists, "Azure/review-checklists"),
        )
        for analyze_pillar, default_source in analysis_passes:
            first_new_item = len(self.action_items)
            analyze_pillar()
            for item in self.action_items[first_new_item:]:
                item.setdefault("source", default_source)
        return self

    def _analyze_checklists(self):
        """Community WAF checks from Azure/review-checklists, run via Resource Graph.
        Adds action items only — pillar scores above still follow the Advisor Score model.
        Source: https://github.com/Azure/review-checklists
        """
        rows = self._data_rows(self.data.get("checklists", {}).get("Findings", []))
        severity_map = {"high": "high", "medium": "medium", "low": "low"}
        grouped = {}
        for row in rows:
            key = row.get("Guid") or row.get("Text")
            if not key:
                continue
            bucket = grouped.setdefault(key, {
                "pillar": row.get("WafPillar") or "Operational Excellence",
                "severity": severity_map.get(str(row.get("Severity", "")).strip().lower(), "medium"),
                "service": str(row.get("Service", "")),
                "text": str(row.get("Text", "")),
                "link": str(row.get("Link", "")),
                "resources": [],
                "rows": []
            })
            bucket["rows"].append(row)
            resource_id = row.get("ResourceId", "")
            if resource_id:
                bucket["resources"].append(escape(str(resource_id)))

        for bucket in grouped.values():
            if bucket["pillar"] not in self.scores:
                continue
            text = escape(bucket["text"])
            title = text if len(text) <= 90 else text[:87] + "..."
            link_html = f' <a href="{escape(bucket["link"])}" target="_blank">Learn more</a>' if bucket["link"] else ""
            count = len(bucket["resources"])
            desc = (f'{text}{link_html} — {count} non-compliant resource(s). '
                    f'Source: Azure/review-checklists community WAF checklist ({escape(bucket["service"])}).')
            self.action_items.append({
                "pillar": bucket["pillar"],
                "severity": bucket["severity"],
                "source": "Azure/review-checklists",
                "title": f'[{escape(bucket["service"])}] {title}',
                "title_es": f'[{escape(bucket["service"])}] {title}',
                "title_pt": f'[{escape(bucket["service"])}] {title}',
                "description": desc,
                "description_es": desc,
                "description_pt": desc,
                "subscriptions": self._subscription_counts(bucket["rows"]),
                "resources": bucket["resources"][:10]
            })
    
    def _analyze_reliability(self):
        """Reliability score using the official Advisor Score subcategory model."""
        total_pool = self._resource_pool()
        forced_zero = set()
        synthetic = []

        # Check VMs without availability zones
        vms = self.data.get("discovery", {}).get("VMs", [])
        vms_no_zone = [v for v in vms if not v.get("zone") and not v.get("availabilityZone")]
        synthetic += [{"problem": "Virtual machine not deployed to an Availability Zone",
                       "impactedResource": v.get("name", v.get("Name", "")),
                       "Subscription": self._sub_name_of(v)} for v in vms_no_zone]
        if vms_no_zone:
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "high",
                "title": f"{len(vms_no_zone)} VMs without Availability Zones",
                "title_es": f"{len(vms_no_zone)} VMs sin Zonas de Disponibilidad",
                "title_pt": f"{len(vms_no_zone)} VMs sem Zonas de Disponibilidade",
                "description": "These VMs are not protected against datacenter failures. Consider migrating to zone-redundant deployments.",
                "description_es": "Estas VMs no están protegidas contra fallas del centro de datos. Considere migrar a implementaciones con redundancia de zona.",
                "description_pt": "Essas VMs não estão protegidas contra falhas de datacenter. Considere migrar para implantações com redundância de zona.",
                "subscriptions": self._subscription_counts(vms_no_zone),
                "resources": [{"name": v.get("name", v.get("Name", "")), "id": v.get("id")} for v in vms_no_zone[:10]]
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
                "title_pt": "Nenhum Vault de Backup/Recovery Services encontrado",
                "description": "No backup infrastructure detected. Critical workloads should have backup policies configured.",
                "description_es": "No se detectó infraestructura de copia de seguridad. Las cargas de trabajo críticas deben tener políticas de backup configuradas.",
                "description_pt": "Nenhuma infraestrutura de backup detectada. Cargas de trabalho críticas devem ter políticas de backup configuradas.",
                "resources": []
            })
        
        # Check DDoS protection on VNets
        vnets = self.data.get("discovery", {}).get("VNets", []) or self.data.get("governance", {}).get("VNets", [])
        vnets_no_ddos = [v for v in vnets if str(v.get("ddos", v.get("ddosProtection", v.get("enableDdosProtection", "")))).lower() in ("false", "", "none")]
        synthetic += [{"problem": "VNet without DDoS Protection Standard",
                       "impactedResource": v.get("name", v.get("Name", "")),
                       "Subscription": self._sub_name_of(v)} for v in vnets_no_ddos]
        if vnets_no_ddos:
            self._add_ddos_action_items(vnets_no_ddos)
        
        # Advisor HighAvailability recommendations
        ha_recs = self.data.get("advisor", {}).get("Reliability", [])
        self._expand_advisor_items(ha_recs, "Reliability")

        buckets = AdvisorScoreModel.merge_buckets(
            AdvisorScoreModel.bucket(ha_recs, AdvisorScoreModel.RELIABILITY_RULES),
            AdvisorScoreModel.bucket(synthetic, AdvisorScoreModel.RELIABILITY_RULES),
        )
        self._apply_advisor_spec("Reliability", ScoreModel.advisor_spec(
            AdvisorScoreModel.RELIABILITY_WEIGHTS, buckets, self._pool_by_subscription(), forced_zero))
    
    def _vnet_public_exposure(self):
        """Public IP exposure per VNet (lowercased VNet id -> row), or None when the Discovery
        export didn't collect it (older workbook, or the exposure query was skipped)."""
        sheet = self.data.get("discovery", {}).get("VNetPublicExposure")
        if sheet is None:
            return None
        return {str(row["VNetId"]).lower(): row for row in self._data_rows(sheet) if row.get("VNetId")}

    def _add_ddos_action_items(self, vnets_no_ddos: list) -> None:
        """DDoS Protection only matters for VNets that host internet-reachable public IPs, so
        severity follows measured exposure instead of treating every unprotected VNet alike."""
        def vnet_ref(vnet: dict) -> dict:
            return {"name": vnet.get("name", vnet.get("Name", "")), "id": vnet.get("id") or vnet.get("ResourceId")}

        exposure = self._vnet_public_exposure()
        if exposure is None:
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "medium",
                "title": f"{len(vnets_no_ddos)} VNets without DDoS Protection",
                "title_es": f"{len(vnets_no_ddos)} VNets sin Protección DDoS",
                "title_pt": f"{len(vnets_no_ddos)} VNets sem Proteção DDoS",
                "description": "Enable Azure DDoS Protection for internet-facing workloads. Public IP exposure per VNet was not collected, so severity could not be refined: re-run the Discovery phase to prioritize.",
                "description_es": "Habilite Azure DDoS Protection para cargas de trabajo expuestas a Internet. No se recolectó la exposición de IPs públicas por VNet, por lo que no se pudo refinar la severidad: vuelva a ejecutar la fase de Discovery para priorizar.",
                "description_pt": "Habilite o Azure DDoS Protection para cargas de trabalho expostas à Internet. A exposição de IPs públicos por VNet não foi coletada, então a severidade não pôde ser refinada: execute novamente a fase de Discovery para priorizar.",
                "subscriptions": self._subscription_counts(vnets_no_ddos),
                "resources": [vnet_ref(v) for v in vnets_no_ddos[:10]]
            })
            return

        def exposure_of(vnet: dict):
            return exposure.get(str(vnet.get("id") or vnet.get("ResourceId") or "").lower())

        exposed = sorted(
            (v for v in vnets_no_ddos if exposure_of(v)),
            key=lambda v: -int(exposure_of(v).get("PublicIPCount") or 0),
        )
        unexposed = [v for v in vnets_no_ddos if not exposure_of(v)]

        if exposed:
            public_ips = sum(int(exposure_of(v).get("PublicIPCount") or 0) for v in exposed)
            paths = sorted({
                path.strip()
                for v in exposed
                for path in str(exposure_of(v).get("ExposedVia", "")).split(",")
                if path.strip()
            })
            via = ", ".join(paths)
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "high",
                "internet_exposed": True,
                "title": f"{len(exposed)} VNets with internet-facing public IPs and no DDoS Protection",
                "title_es": f"{len(exposed)} VNets con IPs públicas expuestas a Internet y sin Protección DDoS",
                "title_pt": f"{len(exposed)} VNets com IPs públicos expostos à Internet e sem Proteção DDoS",
                "description": f"{public_ips} public IP(s) are reachable through these VNets (via {via}). Enable Azure DDoS Network Protection or DDoS IP Protection for them.",
                "description_es": f"{public_ips} IP(s) pública(s) son accesibles a través de estas VNets (mediante {via}). Habilite Azure DDoS Network Protection o DDoS IP Protection para ellas.",
                "description_pt": f"{public_ips} IP(s) público(s) são acessíveis por meio dessas VNets (via {via}). Habilite o Azure DDoS Network Protection ou o DDoS IP Protection para eles.",
                "subscriptions": self._subscription_counts(exposed),
                "resources": [vnet_ref(v) for v in exposed[:10]]
            })
        if unexposed:
            self.action_items.append({
                "pillar": "Reliability",
                "severity": "low",
                "title": f"{len(unexposed)} VNets without DDoS Protection (no public IPs detected)",
                "title_es": f"{len(unexposed)} VNets sin Protección DDoS (sin IPs públicas detectadas)",
                "title_pt": f"{len(unexposed)} VNets sem Proteção DDoS (sem IPs públicos detectados)",
                "description": "No public IPs were found behind these VNets (checked network interfaces, Application Gateway, Firewall, Bastion, gateways, and public load balancers), so a DDoS plan is not a priority. Revisit if public endpoints are added.",
                "description_es": "No se encontraron IPs públicas detrás de estas VNets (se revisaron interfaces de red, Application Gateway, Firewall, Bastion, gateways y balanceadores públicos), por lo que un plan DDoS no es prioritario. Reevalúe si se agregan endpoints públicos.",
                "description_pt": "Nenhum IP público foi encontrado por trás dessas VNets (foram verificados interfaces de rede, Application Gateway, Firewall, Bastion, gateways e balanceadores públicos), então um plano DDoS não é prioritário. Reavalie se endpoints públicos forem adicionados.",
                "subscriptions": self._subscription_counts(unexposed),
                "resources": [vnet_ref(v) for v in unexposed[:10]]
            })

    def _analyze_security(self):
        """Security score = Microsoft Defender Secure Score average (same model Advisor uses)."""
        security = self.data.get("security", {})
        secure_scores = self._data_rows(security.get("SecureScores", []))
        if not secure_scores:
            secure_scores = self._data_rows(self.data.get("governance", {}).get("SecureScores", []))
        avg_score = None
        secure_pairs = []
        if secure_scores:
            values = [
                float(s.get("percentageScore", s.get("pct", s.get("percentage", 0))) or 0)
                for s in secure_scores
            ]
            avg_score = sum(values) / len(values)
            scale = 100.0 if avg_score <= 1 else 1.0  # export may store a 0-1 ratio instead of 0-100
            avg_score *= scale
            secure_pairs = [[self._sub_name_of(s), value * scale] for s, value in zip(secure_scores, values)]
            self.insights.append({
                "pillar": "Security",
                "text": f"Average Defender Secure Score: {avg_score:.0f}%"
            })
        security_spec = {"kind": "secure", "scores": secure_pairs}
        self.score_specs["Security"] = security_spec
        secure_result = ScoreModel.pillar(security_spec)
        
        # Storage accounts with public access
        storage = self.data.get("discovery", {}).get("Storage", [])
        public_storage = [s for s in storage if str(s.get("publicBlob", s.get("allowBlobPublicAccess", ""))).lower() == "true"]
        if public_storage:
            self.action_items.append({
                "pillar": "Security",
                "severity": "critical",
                "title": f"{len(public_storage)} Storage accounts allow public blob access",
                "title_es": f"{len(public_storage)} Cuentas de almacenamiento permiten acceso público a blobs",
                "title_pt": f"{len(public_storage)} Contas de armazenamento permitem acesso público a blobs",
                "description": "Disable public blob access unless explicitly required. This is a common data exposure risk.",
                "description_es": "Deshabilite el acceso público a blobs a menos que sea explícitamente necesario. Este es un riesgo común de exposición de datos.",
                "description_pt": "Desabilite o acesso público a blobs, a menos que seja explicitamente necessário. Este é um risco comum de exposição de dados.",
                "subscriptions": self._subscription_counts(public_storage),
                "resources": [{"name": s.get("name", s.get("Name", "")), "id": s.get("id")} for s in public_storage[:10]]
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
                "title_pt": f"{len(kv_no_purge)} Key Vaults sem Proteção contra Purga",
                "description": "Enable purge protection to prevent permanent deletion of secrets, keys, and certificates.",
                "description_es": "Habilite la protección contra purga para evitar la eliminación permanente de secretos, claves y certificados.",
                "description_pt": "Habilite a proteção contra exclusão definitiva para evitar a exclusão permanente de segredos, chaves e certificados.",
                "subscriptions": self._subscription_counts(kv_no_purge),
                "resources": [{"name": k.get("name", k.get("Name", "")), "id": k.get("id")} for k in kv_no_purge[:10]]
            })
        
        # Storage without TLS 1.2
        old_tls = [s for s in storage if s.get("tlsVersion", s.get("minimumTlsVersion", "")) and "1.2" not in str(s.get("tlsVersion", s.get("minimumTlsVersion", "")))]
        if old_tls:
            self.action_items.append({
                "pillar": "Security",
                "severity": "high",
                "title": f"{len(old_tls)} Storage accounts not enforcing TLS 1.2",
                "title_es": f"{len(old_tls)} Cuentas de almacenamiento sin exigir TLS 1.2",
                "title_pt": f"{len(old_tls)} Contas de armazenamento sem exigir TLS 1.2",
                "description": "Enforce minimum TLS 1.2 for all storage accounts to prevent protocol downgrade attacks.",
                "description_es": "Exija TLS 1.2 como mínimo en todas las cuentas de almacenamiento para evitar ataques de downgrade de protocolo.",
                "description_pt": "Exija TLS 1.2 no mínimo em todas as contas de armazenamento para evitar ataques de downgrade de protocolo.",
                "subscriptions": self._subscription_counts(old_tls),
                "resources": [{"name": s.get("name", s.get("Name", "")), "id": s.get("id")} for s in old_tls[:10]]
            })
        
        # Security advisor recommendations (informational — already reflected in Secure Score)
        sec_recs = self._data_rows(self.data.get("advisor", {}).get("Security", []))
        self._expand_advisor_items(sec_recs, "Security")
        
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
                "title_pt": f"{len(broad_privileged)} Atribuições de função Owner/Contributor com escopo de assinatura ou superior",
                "description": "Review broad-scope privileged role assignments for least-privilege access. Prefer scoping to resource groups.",
                "description_es": "Revise las asignaciones de roles privilegiados de alcance amplio para aplicar el principio de mínimo privilegio. Prefiera limitar el alcance a grupos de recursos.",
                "description_pt": "Revise as atribuições de funções privilegiadas de amplo escopo para aplicar o princípio de menor privilégio. Prefira limitar o escopo a grupos de recursos.",
                "subscriptions": self._subscription_counts(broad_privileged),
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
                "title_pt": f"{len(disabled_plans)} Planos do Microsoft Defender for Cloud não habilitados (nível gratuito)",
                "description": "Enable Microsoft Defender for Cloud plans for full threat protection coverage.",
                "description_es": "Habilite los planes de Microsoft Defender for Cloud para una cobertura completa de protección contra amenazas.",
                "description_pt": "Habilite os planos do Microsoft Defender for Cloud para cobertura completa de proteção contra ameaças.",
                "subscriptions": self._subscription_counts(disabled_plans),
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
                "title_pt": f"{len(severe_recommendations)} recomendações de alta severidade do Defender for Cloud",
                "description": "Prioritize unhealthy CSPM recommendations that carry critical or high severity.",
                "description_es": "Priorice las recomendaciones de CSPM no saludables con severidad crítica o alta.",
                "description_pt": "Priorize as recomendações de CSPM não saudáveis com severidade crítica ou alta.",
                "subscriptions": self._subscription_counts(severe_recommendations),
                "resources": [
                    {"name": recommendation.get("recommendationName", recommendation.get("affectedResourceId", "")), "id": recommendation.get("affectedResourceId")}
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
                "title_pt": f"{len(severe_incidents)} incidentes de segurança ativos de alta severidade",
                "description": "Investigate active Microsoft Defender XDR incidents and confirm ownership and containment.",
                "description_es": "Investigue los incidentes activos de Microsoft Defender XDR y confirme su asignación y contención.",
                "description_pt": "Investigue os incidentes ativos do Microsoft Defender XDR e confirme a responsabilidade e a contenção.",
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
        
        self.scores["Security"] = secure_result["score"] if secure_result["available"] else 0
        self.breakdowns["Security"] = {
            "available": secure_result["available"],
            "secure_score_pct": secure_result["pct"],
            "subscriptions_scored": secure_result["scored"],
            "findings": [
                ("Public Blob Access on Storage Accounts", "Acceso p\u00fablico a blobs en cuentas de almacenamiento", "Acesso p\u00fablico a blobs em contas de armazenamento", len(public_storage)),
                ("Key Vaults without Purge Protection", "Key Vaults sin Protecci\u00f3n contra Purga", "Key Vaults sem Prote\u00e7\u00e3o contra Purga", len(kv_no_purge)),
                ("Storage Accounts below TLS 1.2", "Cuentas de almacenamiento por debajo de TLS 1.2", "Contas de armazenamento abaixo do TLS 1.2", len(old_tls)),
                ("Broad-scope Owner/Contributor Assignments", "Asignaciones Owner/Contributor de alcance amplio", "Atribui\u00e7\u00f5es Owner/Contributor de amplo escopo", len(broad_privileged)),
                ("Defender for Cloud Plans on Free Tier", "Planes de Defender for Cloud en nivel gratuito", "Planos do Defender for Cloud no n\u00edvel gratuito", len(disabled_plans)),
                ("Open Defender for Cloud Recommendations", "Recomendaciones abiertas de Defender for Cloud", "Recomenda\u00e7\u00f5es abertas do Defender for Cloud", len(open_recommendations)),
                ("Unhealthy Secure Score Control Resources", "Recursos no saludables en controles de Secure Score", "Recursos n\u00e3o saud\u00e1veis em controles do Secure Score", unhealthy_control_resources),
                ("Failed MCSB Resources", "Recursos con error en MCSB", "Recursos com falha no MCSB", mcsb_failed_resources),
                ("Active Defender XDR Incidents", "Incidentes activos de Defender XDR", "Incidentes ativos do Defender XDR", len(active_incidents)),
                ("Active Security Alerts", "Alertas de seguridad activas", "Alertas de seguran\u00e7a ativos", len(active_alerts) + len(active_cloud_alerts)),
                ("Advisor Security Recommendations", "Recomendaciones de seguridad de Advisor", "Recomenda\u00e7\u00f5es de seguran\u00e7a do Advisor", len(sec_recs)),
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
            keys = {self._qkey(o, o.get("Name", o.get("name", id(o)))) for o in orphaned}
            impacted_keys |= keys
            sources.append(self._cost_source("Orphaned/unused resources", "resources", keys))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "medium",
                "title": f"{len(orphaned)} Orphaned/unused resources detected",
                "title_es": f"{len(orphaned)} Recursos huérfanos/no utilizados detectados",
                "title_pt": f"{len(orphaned)} Recursos órfãos/não utilizados detectados",
                "description": "These resources incur costs but are not attached to any workload. Review and delete if unused.",
                "description_es": "Estos recursos generan costos pero no están asociados a ninguna carga de trabajo. Revíselos y elimínelos si no se utilizan.",
                "description_pt": "Esses recursos geram custos, mas não estão associados a nenhuma carga de trabalho. Revise e exclua se não forem utilizados.",
                "subscriptions": self._subscription_counts(orphaned),
                "resources": [{"name": o.get("Name", o.get("name", "")), "id": o.get("id") or o.get("ResourceId")} for o in orphaned[:10]]
            })
        
        # VMs that are deallocated (still paying for disks)
        dealloc_vms = self._data_rows(self.data.get("discovery", {}).get("DeallocatedVMs", []))
        if dealloc_vms:
            keys = {self._qkey(v, v.get("name", v.get("Name", id(v)))) for v in dealloc_vms}
            impacted_keys |= keys
            sources.append(self._cost_source("Deallocated VMs", "resources", keys))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "low",
                "title": f"{len(dealloc_vms)} Deallocated VMs (still paying for disks/IPs)",
                "title_es": f"{len(dealloc_vms)} VMs desasignadas (aún generan costo por discos/IPs)",
                "title_pt": f"{len(dealloc_vms)} VMs desalocadas (ainda geram custo por discos/IPs)",
                "description": "Deallocated VMs still incur costs for attached disks and static IPs. Consider deleting if no longer needed.",
                "description_es": "Las VMs desasignadas siguen generando costos por los discos e IPs estáticas asociadas. Considere eliminarlas si ya no son necesarias.",
                "description_pt": "VMs desalocadas continuam gerando custos pelos discos e IPs estáticos associados. Considere excluí-las se não forem mais necessárias.",
                "subscriptions": self._subscription_counts(dealloc_vms),
                "resources": [{"name": v.get("name", v.get("Name", "")), "id": v.get("id")} for v in dealloc_vms[:10]]
            })
        
        # Right-sizing from metrics
        vm_metrics = self._data_rows(self.data.get("metrics", {}).get("VM_RightSizing", []))
        underutilized = [v for v in vm_metrics if "Idle" in str(v.get("Assessment", "")) or "Underutilized" in str(v.get("Assessment", ""))]
        if underutilized:
            keys = {self._qkey(v, v.get("Name", id(v))) for v in underutilized}
            impacted_keys |= keys
            sources.append(self._cost_source("Idle/underutilized VMs", "resources", keys))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "high",
                "title": f"{len(underutilized)} VMs are idle or underutilized",
                "title_es": f"{len(underutilized)} VMs están inactivas o subutilizadas",
                "title_pt": f"{len(underutilized)} VMs estão ociosas ou subutilizadas",
                "description": "These VMs have <15% average CPU over 30 days. Consider downsizing or deallocating.",
                "description_es": "Estas VMs tienen <15% de uso promedio de CPU en 30 días. Considere reducir su tamaño o desasignarlas.",
                "description_pt": "Essas VMs têm menos de 15% de uso médio de CPU em 30 dias. Considere redimensionar ou desalocar.",
                "subscriptions": self._subscription_counts(underutilized),
                "resources": [{"name": f"{v.get('Name', '')} ({v.get('VMSize', '')} @ {v.get('AvgCPU_Pct', '?')}% CPU)", "id": v.get("ResourceId")} for v in underutilized[:10]]
            })
        
        # SQL databases that are oversized or underutilized (DTU/CPU-based)
        sql_metrics = self._data_rows(self.data.get("metrics", {}).get("SQL_RightSizing", []))
        sql_waste = [d for d in sql_metrics if str(d.get("Assessment", "")).startswith(("Oversized", "Underutilized"))]
        if sql_waste:
            keys = {self._qkey(d, d.get("Name", id(d))) for d in sql_waste}
            impacted_keys |= keys
            sources.append(self._cost_source("Oversized/underutilized SQL databases", "resources", keys))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "medium",
                "title": f"{len(sql_waste)} SQL databases are oversized or underutilized",
                "title_es": f"{len(sql_waste)} bases de datos SQL están sobredimensionadas o subutilizadas",
                "title_pt": f"{len(sql_waste)} bancos de dados SQL estão superdimensionados ou subutilizados",
                "description": "These databases run well below their provisioned DTU/CPU tier over 30 days. Consider a lower SKU.",
                "description_es": "Estas bases de datos operan muy por debajo de su nivel de DTU/CPU aprovisionado durante 30 días. Considere un SKU menor.",
                "description_pt": "Esses bancos de dados operam bem abaixo do nível de DTU/CPU provisionado durante 30 dias. Considere um SKU menor.",
                "subscriptions": self._subscription_counts(sql_waste),
                "resources": [{"name": f"{d.get('Name', '')} ({d.get('SKU', '')} @ {d.get('AvgUsage_Pct', '?')}% {d.get('MetricType', '')})", "id": d.get("ResourceId")} for d in sql_waste[:10]]
            })

        # App Service Plans that are idle or underutilized
        plan_metrics = self._data_rows(self.data.get("metrics", {}).get("AppPlan_RightSizing", []))
        plan_waste = [p for p in plan_metrics if str(p.get("Assessment", "")).startswith(("Idle", "Underutilized"))]
        if plan_waste:
            keys = {self._qkey(p, p.get("Name", id(p))) for p in plan_waste}
            impacted_keys |= keys
            sources.append(self._cost_source("Idle/underutilized App Service Plans", "resources", keys))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "medium",
                "title": f"{len(plan_waste)} App Service Plans are idle or underutilized",
                "title_es": f"{len(plan_waste)} planes de App Service están inactivos o subutilizados",
                "title_pt": f"{len(plan_waste)} planos do App Service estão ociosos ou subutilizados",
                "description": "These plans run well below their provisioned CPU/memory over 30 days. Consider scaling down or consolidating apps.",
                "description_es": "Estos planes operan muy por debajo de su CPU/memoria aprovisionada durante 30 días. Considere reducir su tamaño o consolidar aplicaciones.",
                "description_pt": "Esses planos operam bem abaixo da CPU/memória provisionada durante 30 dias. Considere reduzir o tamanho ou consolidar aplicativos.",
                "subscriptions": self._subscription_counts(plan_waste),
                "resources": [{"name": f"{p.get('Name', '')} ({p.get('SKU', '')} @ {p.get('AvgCPU_Pct', '?')}% CPU)", "id": p.get("ResourceId")} for p in plan_waste[:10]]
            })

        # Storage accounts with zero/minimal activity
        storage_metrics = self._data_rows(self.data.get("metrics", {}).get("Storage_Activity", []))
        storage_waste = [s for s in storage_metrics if str(s.get("Assessment", "")).startswith(("Zero", "Minimal"))]
        if storage_waste:
            keys = {self._qkey(s, s.get("Name", id(s))) for s in storage_waste}
            impacted_keys |= keys
            sources.append(self._cost_source("Storage accounts with zero/minimal activity", "resources", keys))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "low",
                "title": f"{len(storage_waste)} Storage accounts show zero/minimal activity",
                "title_es": f"{len(storage_waste)} cuentas de almacenamiento muestran actividad nula/mínima",
                "title_pt": f"{len(storage_waste)} contas de armazenamento apresentam atividade nula/mínima",
                "description": "These storage accounts had near-zero transactions over 30 days. Review for deletion or tier downgrade (e.g. to Archive).",
                "description_es": "Estas cuentas de almacenamiento tuvieron transacciones casi nulas durante 30 días. Considere eliminarlas o cambiarlas a un nivel más económico (p. ej. Archive).",
                "description_pt": "Essas contas de armazenamento tiveram transações quase nulas durante 30 dias. Considere excluí-las ou migrá-las para um nível mais econômico (ex.: Archive).",
                "subscriptions": self._subscription_counts(storage_waste),
                "resources": [{"name": f"{s.get('Name', '')} ({s.get('SKU', '')} @ {s.get('AvgDailyTxns', '?')} txns/day)", "id": s.get("ResourceId")} for s in storage_waste[:10]]
            })

        # Cost recommendations from Advisor
        cost_recs = self._data_rows(self.data.get("advisor", {}).get("Cost", []))
        if cost_recs:
            keys = {self._qkey(r, r.get("impactedResource", r.get("impactedValue", id(r)))) for r in cost_recs}
            impacted_keys |= keys
            sources.append(self._cost_source("Advisor Cost recommendations", "recommendations", keys, cost_recs))
            self._expand_advisor_items(cost_recs, "Cost Optimization", savings_field="annualSavings")

        # Reservation / Savings Plan recommendations (rate optimization). Informational only -
        # these are opportunities on already-healthy resources, so they don't affect the score,
        # same treatment as the community checklist findings.
        reservation_recs = [
            r for r in self._data_rows(self.data.get("advisor", {}).get("ReservationRecommendations", []))
            if r.get("ResourceType") or r.get("SkuName")
        ]
        if reservation_recs:
            total_net_savings = sum(float(r.get("NetSavings", 0) or 0) for r in reservation_recs)
            currency = next((r.get("Currency") for r in reservation_recs if r.get("Currency")), "USD")
            desc = f"Buying the recommended Reserved Instances/Savings Plans could save ~{total_net_savings:,.0f} {currency}/year vs. pay-as-you-go."
            desc_es = f"Comprar las Reservas/Planes de Ahorro recomendados podría ahorrar ~{total_net_savings:,.0f} {currency}/año frente a pago por uso."
            desc_pt = f"Comprar as Reservas/Planos de Economia recomendados pode economizar ~{total_net_savings:,.0f} {currency}/ano em relação ao pagamento conforme o uso."
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "medium",
                "title": f"{len(reservation_recs)} Reservation/Savings Plan opportunities",
                "title_es": f"{len(reservation_recs)} oportunidades de Reservas/Planes de Ahorro",
                "title_pt": f"{len(reservation_recs)} oportunidades de Reservas/Planos de Economia",
                "description": desc,
                "description_es": desc_es,
                "description_pt": desc_pt,
                "subscriptions": self._subscription_counts(reservation_recs),
                "resources": [f"{r.get('ResourceType', '')} {r.get('SkuName', '')} x{r.get('RecommendedQty', '?')} ({r.get('Term', '')})" for r in reservation_recs[:10]]
            })

        # FinOps extended recommendations (unattached public IPs, stopped VMs, backendless
        # gateways/load balancers, empty SQL elastic pools, non-spot AKS pools, missing
        # Hybrid Benefit) - the extra Azure Resource Graph checks FinOps Hub runs daily.
        finops = self.data.get("finops", {})
        extended_sheets = {
            "UnattachedPublicIPs": ("Unattached Public IPs", "IPs Públicas No Adjuntas", "IPs Públicos Não Anexados"),
            "StoppedVMs": ("Stopped (not deallocated) VMs", "VMs Detenidas (no desasignadas)", "VMs Paradas (não desalocadas)"),
            "BackendlessAppGateways": ("Backendless Application Gateways", "Application Gateways sin Backend", "Application Gateways sem Backend"),
            "BackendlessLoadBalancers": ("Backendless Load Balancers", "Load Balancers sin Backend", "Load Balancers sem Backend"),
            "EmptySqlElasticPools": ("Empty SQL Elastic Pools", "Grupos Elásticos SQL Vacíos", "Pools Elásticos SQL Vazios"),
            "NonSpotAKSPools": ("Non-Spot autoscaling AKS pools", "Pools AKS con autoescalado sin Spot", "Pools AKS com autoescala sem Spot"),
            "VMsWithoutHybridBenefit": ("VMs without Azure Hybrid Benefit", "VMs sin Azure Hybrid Benefit", "VMs sem Azure Hybrid Benefit"),
            "SqlVMsWithoutHybridBenefit": ("SQL VMs without Azure Hybrid Benefit", "VMs SQL sin Azure Hybrid Benefit", "VMs SQL sem Azure Hybrid Benefit"),
        }
        for sheet_name, (title_en, title_es, title_pt) in extended_sheets.items():
            rows = self._data_rows(finops.get(sheet_name, []))
            if not rows:
                continue
            keys = {self._qkey(r, r.get("name", id(r))) for r in rows}
            impacted_keys |= keys
            sources.append(self._cost_source(title_en, "resources", keys, rows))
            self.action_items.append({
                "pillar": "Cost Optimization",
                "severity": "low",
                "title": f"{len(rows)} {title_en}",
                "title_es": f"{len(rows)} {title_es}",
                "title_pt": f"{len(rows)} {title_pt}",
                "description": "Extra FinOps optimization check (mirrors Azure FinOps Hub's daily Resource Graph queries). Review and remediate if unused.",
                "description_es": "Verificación adicional de optimización FinOps (equivalente a las consultas diarias de Resource Graph de Azure FinOps Hub). Revise y corrija si no se utiliza.",
                "description_pt": "Verificação adicional de otimização FinOps (equivalente às consultas diárias do Resource Graph do Azure FinOps Hub). Revise e corrija se não for utilizado.",
                "subscriptions": self._subscription_counts(rows),
                "resources": [str(r.get("name", "")) for r in rows[:10]]
            })

        spec = {"kind": "ratio", "pool": self._pool_by_subscription(),
                "impacted": ScoreModel.group_keys(impacted_keys), "sources": sources}
        result = ScoreModel.pillar(spec)
        self.score_specs["Cost Optimization"] = spec
        self.scores["Cost Optimization"] = result["score"] if result["available"] else 0
        self.breakdowns["Cost Optimization"] = {
            "available": result["available"],
            "sources": result["sources"],
            "impacted": result["impacted"],
            "healthy": result["healthy"],
            "total": result["total"],
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
                               "impactedResource": d.get("Name", d.get("name", "")),
                               "Subscription": self._sub_name_of(d)} for d in no_diag]
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "high",
                    "title": f"{len(no_diag)}/{len(diag)} critical resources missing diagnostic settings ({pct_missing:.0f}%)",
                    "title_es": f"{len(no_diag)}/{len(diag)} recursos críticos sin configuración de diagnóstico ({pct_missing:.0f}%)",
                    "title_pt": f"{len(no_diag)}/{len(diag)} recursos críticos sem configuração de diagnóstico ({pct_missing:.0f}%)",
                    "description": "Configure diagnostic settings to send logs to Log Analytics for observability and troubleshooting.",
                    "description_es": "Configure ajustes de diagnóstico para enviar registros a Log Analytics para observabilidad y solución de problemas.",
                    "description_pt": "Configure as configurações de diagnóstico para enviar logs ao Log Analytics para observabilidade e solução de problemas.",
                    "subscriptions": self._subscription_counts(no_diag),
                    "resources": [{"name": d.get("Name", d.get("name", "")), "id": d.get("ResourceId")} for d in no_diag[:10]]
                })
        
        # Tag coverage — computed from the per-resource `tags` field across Discovery inventory
        # sheets (there's no dedicated tag-usage export; GovViz's Subscriptions.tags is subscription-level only).
        TAGGABLE_SHEETS = ["VMs", "AppServices", "AKS", "VNets", "NSGs", "LoadBalancers",
                           "Firewalls", "PublicIPs", "Storage", "Databases", "KeyVaults"]
        discovery = self.data.get("discovery", {})
        total_taggable, untagged = 0, []
        for sheet in TAGGABLE_SHEETS:
            for r in discovery.get(sheet, []):
                total_taggable += 1
                tags = r.get("tags")
                if not tags or str(tags).strip().lower() in ("", "none", "{}"):
                    untagged.append({"name": r.get("name", r.get("Name", "")), "id": r.get("id"), "Subscription": self._sub_name_of(r)})
        if total_taggable > 0 and untagged:
            pct_untagged = len(untagged) / total_taggable * 100
            synthetic += [{"problem": "Resource has no tags configured", "impactedResource": n["name"], "Subscription": n["Subscription"]} for n in untagged]
            self.action_items.append({
                "pillar": "Operational Excellence",
                "severity": "medium",
                "title": f"{len(untagged)}/{total_taggable} resources have no tags ({pct_untagged:.0f}%)",
                "title_es": f"{len(untagged)}/{total_taggable} recursos no tienen etiquetas ({pct_untagged:.0f}%)",
                "title_pt": f"{len(untagged)}/{total_taggable} recursos não têm tags ({pct_untagged:.0f}%)",
                "description": "Implement a tagging strategy (Owner, CostCenter, Environment, Application) for governance and cost allocation.",
                "description_es": "Implemente una estrategia de etiquetado (Owner, CostCenter, Environment, Application) para gobernanza y asignación de costos.",
                "description_pt": "Implemente uma estratégia de tags (Owner, CostCenter, Environment, Application) para governança e alocação de custos.",
                "subscriptions": self._subscription_counts(untagged),
                "resources": untagged[:10]
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
                "title_pt": "Nenhum bloqueio de recurso configurado",
                "description": "Apply CanNotDelete or ReadOnly locks to critical resources (networking, databases, key vaults) to prevent accidental deletion or modification.",
                "description_es": "Aplique bloqueos CanNotDelete o ReadOnly a los recursos críticos (redes, bases de datos, key vaults) para evitar eliminaciones o modificaciones accidentales.",
                "description_pt": "Aplique bloqueios CanNotDelete ou ReadOnly a recursos críticos (redes, bancos de dados, key vaults) para evitar exclusões ou modificações acidentais.",
                "resources": []
            })
        
        # Policy compliance
        compliance = self.data.get("governance", {}).get("PolicyCompliance", [])
        if compliance:
            total_nc = sum(int(c.get("NonCompliantCount", 0) or 0) for c in compliance)
            non_compliant = [c for c in compliance if int(c.get("NonCompliantCount", 0) or 0) > 0]
            synthetic += [{"problem": "Non-compliant policy evaluation",
                           "impactedResource": c.get("policyAssignment", c.get("policyAssignmentName", "")),
                           "Subscription": self._sub_name_of(c)} for c in non_compliant]
            if total_nc > 0:
                self.action_items.append({
                    "pillar": "Operational Excellence",
                    "severity": "medium",
                    "title": f"{total_nc} non-compliant policy evaluations",
                    "title_es": f"{total_nc} evaluaciones de políticas no conformes",
                    "title_pt": f"{total_nc} avaliações de política não conformes",
                    "description": "Review policy compliance and remediate non-compliant resources or adjust policy assignments.",
                    "description_es": "Revise el cumplimiento de políticas y corrija los recursos no conformes o ajuste las asignaciones de políticas.",
                    "description_pt": "Revise a conformidade de políticas e corrija os recursos não conformes ou ajuste as atribuições de política.",
                    "subscriptions": self._subscription_counts(compliance),
                    "resources": [f"{c.get('policyAssignment', c.get('policyAssignmentName', ''))} ({c.get('NonCompliantCount', 0)} violations)" for c in compliance[:10]]
                })
        
        # OpEx Advisor recommendations
        opex_recs = self.data.get("advisor", {}).get("OperationalExcellence", [])
        self._expand_advisor_items(opex_recs, "Operational Excellence")

        buckets = AdvisorScoreModel.merge_buckets(
            AdvisorScoreModel.bucket(opex_recs, AdvisorScoreModel.OPERATIONAL_RULES),
            AdvisorScoreModel.bucket(synthetic, AdvisorScoreModel.OPERATIONAL_RULES),
        )
        self._apply_advisor_spec("Operational Excellence", ScoreModel.advisor_spec(
            AdvisorScoreModel.OPERATIONAL_WEIGHTS, buckets, self._pool_by_subscription()))
    
    def _analyze_performance(self):
        """Performance score using the official Advisor Score subcategory model."""
        total_pool = self._resource_pool()
        synthetic = []
        
        # Saturated VMs
        vm_metrics = self.data.get("metrics", {}).get("VM_RightSizing", [])
        saturated = [v for v in vm_metrics if "Saturated" in str(v.get("Assessment", ""))]
        synthetic += [{"problem": "Virtual machine CPU saturated",
                       "impactedResource": v.get("Name", ""),
                       "Subscription": self._sub_name_of(v)} for v in saturated]
        if saturated:
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "high",
                "title": f"{len(saturated)} VMs are saturated (>80% CPU)",
                "title_es": f"{len(saturated)} VMs están saturadas (>80% CPU)",
                "title_pt": f"{len(saturated)} VMs estão saturadas (>80% CPU)",
                "description": "These VMs are consistently at high CPU usage. Consider scaling up or scaling out.",
                "description_es": "Estas VMs presentan un uso de CPU consistentemente alto. Considere escalar verticalmente u horizontalmente.",
                "description_pt": "Essas VMs apresentam uso de CPU consistentemente alto. Considere escalar verticalmente ou horizontalmente.",
                "subscriptions": self._subscription_counts(saturated),
                "resources": [{"name": f"{v.get('Name', '')} ({v.get('VMSize', '')} @ {v.get('AvgCPU_Pct', '?')}% CPU)", "id": v.get("ResourceId")} for v in saturated[:10]]
            })
        
        # Saturated SQL
        sql_metrics = self.data.get("metrics", {}).get("SQL_RightSizing", [])
        saturated_sql = [s for s in sql_metrics if "Saturated" in str(s.get("Assessment", ""))]
        synthetic += [{"problem": "SQL database DTU/CPU saturated",
                       "impactedResource": s.get("Name", ""),
                       "Subscription": self._sub_name_of(s)} for s in saturated_sql]
        if saturated_sql:
            self.action_items.append({
                "pillar": "Performance Efficiency",
                "severity": "high",
                "title": f"{len(saturated_sql)} SQL Databases are saturated (>80% DTU/CPU)",
                "title_es": f"{len(saturated_sql)} Bases de datos SQL están saturadas (>80% DTU/CPU)",
                "title_pt": f"{len(saturated_sql)} Bancos de dados SQL estão saturados (>80% DTU/CPU)",
                "description": "These databases are consistently at high utilization. Consider scaling up the tier.",
                "description_es": "Estas bases de datos presentan una utilización consistentemente alta. Considere aumentar el nivel de servicio.",
                "description_pt": "Esses bancos de dados apresentam utilização consistentemente alta. Considere aumentar o nível de serviço.",
                "subscriptions": self._subscription_counts(saturated_sql),
                "resources": [{"name": f"{s.get('Name', '')} ({s.get('SKU', '')} @ {s.get('AvgUsage_Pct', '?')}%)", "id": s.get("ResourceId")} for s in saturated_sql[:10]]
            })
        
        # Performance advisor recommendations
        perf_recs = self.data.get("advisor", {}).get("Performance", [])
        self._expand_advisor_items(perf_recs, "Performance Efficiency")

        buckets = AdvisorScoreModel.merge_buckets(
            AdvisorScoreModel.bucket(perf_recs, AdvisorScoreModel.PERFORMANCE_RULES),
            AdvisorScoreModel.bucket(synthetic, AdvisorScoreModel.PERFORMANCE_RULES),
        )
        self._apply_advisor_spec("Performance Efficiency", ScoreModel.advisor_spec(
            AdvisorScoreModel.PERFORMANCE_WEIGHTS, buckets, self._pool_by_subscription()))
    
    def get_overall_score(self) -> int:
        """Mean of the available category scores (per the official Advisor Score model)."""
        available = [p for p, b in self.breakdowns.items() if b.get("available")]
        if not available:
            return 0
        return round_half_up(sum(self.scores[p] for p in available) / len(available))


class ScoreHistoryTracker:
    """Persists overall/pillar scores and open action-item ids across runs so re-running the
    workshop against the same environment can show score deltas and highlight resolved items.
    Stored as a JSON file SIBLING to base_dir (e.g. "AzureWorkshop.history.json" next to
    "AzureWorkshop/") rather than inside it, since Launch-AzureWorkshop.ps1 wipes/archives
    that folder on every run."""

    MAX_RUNS = 30

    def __init__(self, base_dir: str):
        self.path = base_dir.rstrip("/\\") + ".history.json"
        self.runs = self._load()

    def _load(self) -> list:
        if not os.path.isfile(self.path):
            return []
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            return loaded.get("runs", []) if isinstance(loaded, dict) else []
        except (json.JSONDecodeError, OSError):
            return []

    def compare_and_record(self, overall: int, scores: dict, action_items: list, action_item_id_fn) -> dict:
        """Compare current scores/action items against the most recent previous run, append +
        persist the current snapshot, and return the comparison (deltas are None on the first run)."""
        previous = self.runs[-1] if self.runs else None

        current_open = [
            {"id": action_item_id_fn(a), "pillar": a["pillar"], "title": a["title"], "severity": a["severity"]}
            for a in action_items
        ]
        current_ids = {i["id"] for i in current_open}

        comparison = {"previous": previous, "resolved": [], "new_count": 0, "overall_delta": None, "pillar_deltas": {}}
        if previous:
            previous_ids = {i["id"] for i in previous.get("action_items", [])}
            comparison["resolved"] = [i for i in previous.get("action_items", []) if i["id"] not in current_ids]
            comparison["new_count"] = len(current_ids - previous_ids)
            comparison["overall_delta"] = overall - previous.get("overall_score", overall)
            for pillar, sc in scores.items():
                prev_sc = previous.get("pillar_scores", {}).get(pillar)
                comparison["pillar_deltas"][pillar] = (sc - prev_sc) if prev_sc is not None else None

        self.runs.append({
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "overall_score": overall,
            "pillar_scores": dict(scores),
            "action_items": current_open,
        })
        self.runs = self.runs[-self.MAX_RUNS:]
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump({"runs": self.runs}, f, indent=2)
        except OSError as e:
            print(f"⚠️  Could not save score history ({self.path}): {e}")

        return comparison


# ═══════════════════════════════════════════════════════════════════════════════
# HTML DASHBOARD GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

class DashboardGenerator:
    """Generates the final consolidated HTML dashboard."""
    
    def __init__(self, engine: InsightEngine, data: dict, base_dir: Optional[str] = None, progress: Optional[dict] = None):
        self.engine = engine
        self.data = data
        self.base_dir = base_dir
        self.progress = progress or {"previous": None, "resolved": [], "new_count": 0, "overall_delta": None, "pillar_deltas": {}}
        # Chart.js chart configs (canvasId -> config) contributed by any _render_*_view()
        # method - consumed by initAllCharts() in the embedded <script>, serialized once
        # in generate().
        self._chart_data: dict = {}

    @staticmethod
    def _count_real(rows: list) -> int:
        """Count sheet rows, ignoring the Export-Sheet 'No data found' placeholder row."""
        return len([
            row for row in rows
            if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
        ])

    _ICON_FILE_CACHE: dict = {}
    _icon_instance_counter = 0

    def _azure_icon(self, name: str, css_class: str = "az-icon") -> str:
        """Inline an official Microsoft Azure architecture SVG icon (assets/azure-icons/*.svg).

        Icon ids/gradient refs are namespaced per-embed since the same icon may be
        rendered many times in one page (duplicate SVG ids are otherwise invalid HTML).
        """
        svg = self._ICON_FILE_CACHE.get(name)
        if svg is None:
            icon_path = Path(__file__).parent / "assets" / "azure-icons" / f"{name}.svg"
            svg = icon_path.read_text(encoding="utf-8")
            self._ICON_FILE_CACHE[name] = svg
        DashboardGenerator._icon_instance_counter += 1
        uid = DashboardGenerator._icon_instance_counter
        svg = re.sub(r'id="([^"]+)"', lambda m: f'id="{m.group(1)}-{uid}"', svg)
        svg = re.sub(r'url\(#([^)]+)\)', lambda m: f'url(#{m.group(1)}-{uid})', svg)
        svg = re.sub(r"<svg ", f'<svg class="{css_class}" ', svg, count=1)
        return svg

    # Original mark for this toolkit (favicon only): a dark-on-transparent and a
    # white-on-transparent variant, swapped automatically via prefers-color-scheme so the
    # icon stays legible whether the browser/OS is in light or dark mode.
    _LOGO_PNG_CACHE: dict = {}

    def _favicon_png_base64(self, variant: str) -> str:
        """Read one of the toolkit's own favicon PNGs ('dark' or 'white') and return it
        base64-encoded, used for the browser favicon."""
        cached = DashboardGenerator._LOGO_PNG_CACHE.get(variant)
        if cached is None:
            filename = "dark logo.png" if variant == "dark" else "white logo.png"
            png_path = Path(__file__).parent / "assets" / "branding" / filename
            cached = base64.b64encode(png_path.read_bytes()).decode("ascii")
            DashboardGenerator._LOGO_PNG_CACHE[variant] = cached
        return cached

    _VENDOR_JS_CACHE: dict = {}

    def _vendor_js(self, filename: str) -> str:
        """Read a vendored JS library (assets/vendor/*.js) so it can be inlined into the
        generated HTML - keeps the dashboard a single offline-viewable file with no CDN calls."""
        content = self._VENDOR_JS_CACHE.get(filename)
        if content is None:
            js_path = Path(__file__).parent / "assets" / "vendor" / filename
            content = js_path.read_text(encoding="utf-8") if js_path.exists() else ""
            self._VENDOR_JS_CACHE[filename] = content
        return content

    _RESOURCE_KIND_ICON = {
        "VMs": "resource-compute", "DeallocatedVMs": "resource-compute",
        "AppServices": "resource-compute", "AKS": "resource-compute",
        "VNets": "resource-network", "NSGs": "resource-network",
        "LoadBalancers": "resource-network", "Firewalls": "resource-network",
        "PrivateEndpoints": "resource-network", "PublicIPs": "resource-network",
        "DNS": "resource-network",
        "Storage": "resource-storage", "UnattachedDisks": "resource-storage",
        "Databases": "resource-database",
        "KeyVaults": "resource-security", "KeyVaultExpirations": "resource-security",
        "ManagedIdentities": "resource-security", "PolicyAssignments": "resource-security",
        "Monitoring": "resource-monitor", "ResourceHealthEvents": "resource-monitor",
        "AdvisorRecommendations": "resource-monitor",
        "BackupVaults": "resource-backup", "BackupProtectedItems": "resource-backup",
    }

    def _resource_kind_icon(self, kind: str) -> str:
        return self._azure_icon(self._RESOURCE_KIND_ICON.get(kind, "resource-generic"), "az-icon resource-icon")

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
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            MAX_TABLE_ROWS = 300
            row_total = len(rows)
            rows = rows[:MAX_TABLE_ROWS]
            header = "".join(f"<th>{bi(label_en, label_es, LABEL_PT.get(label_en, label_en))}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for i, row in enumerate(rows):
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{escape(str(row.get(key, "")))}</td>'
                    for key, _, _, numeric in columns
                )
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                row_attr += _resource_link_attr(row)
                body += f"<tr{row_attr}>{cells}</tr>"
            more = ""
            if len(rows) > limit:
                more_text = bi(f"+ {len(rows) - limit} more", f"+ {len(rows) - limit} m\u00e1s", f"+ {len(rows) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            if row_total > MAX_TABLE_ROWS:
                more += f'<p class="table-truncated-note">{bi(f"Showing first {MAX_TABLE_ROWS} of {row_total} rows for performance - see the source export for the full list.", f"Mostrando las primeras {MAX_TABLE_ROWS} de {row_total} filas por rendimiento - consulte el archivo de origen para ver la lista completa.", f"Mostrando as primeiras {MAX_TABLE_ROWS} de {row_total} linhas por desempenho - consulte o arquivo de origem para a lista completa.")}</p>'
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>{more}'

        def disc_real_rows(sheet_name):
            discovery = self.data.get("discovery", {})
            return [
                row for row in discovery.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def render_compliance_bars(rows):
            """Pass-rate bar per regulatory/framework compliance standard (RegulatoryStandards sheet)."""
            totals = {}
            for row in rows:
                name = str(row.get("complianceStandard") or "Unknown")
                passed = as_int(row.get("passedControls"))
                failed = as_int(row.get("failedControls"))
                skipped = as_int(row.get("skippedControls"))
                unsupported = as_int(row.get("unsupportedControls"))
                entry = totals.setdefault(name, {"passed": 0, "total": 0})
                entry["passed"] += passed
                entry["total"] += passed + failed + skipped + unsupported
            if not totals:
                return f'<p class="empty-state">{bi("No regulatory compliance standards found.", "No se encontraron est\u00e1ndares de cumplimiento normativo.", "Nenhum padr\u00e3o de conformidade regulat\u00f3ria encontrado.")}</p>'
            items = sorted(totals.items(), key=lambda kv: (kv[1]["passed"] / kv[1]["total"]) if kv[1]["total"] else 0)
            rows_html = ""
            for name, entry in items:
                pct = (entry["passed"] / entry["total"] * 100) if entry["total"] else 0
                color = "var(--cp-danger)" if pct < 60 else "var(--cp-warning)" if pct < 85 else "var(--cp-success)"
                rows_html += f"""<div class="cost-bar-row">
                    <span class="cost-bar-label">{escape(name)}</span>
                    <div class="cost-bar-track"><div class="cost-bar-fill" style="width:{pct:.0f}%; background:{color};"></div></div>
                    <span class="cost-bar-value">{entry['passed']}/{entry['total']}</span>
                    <span class="cost-bar-pct">{pct:.0f}%</span>
                </div>"""
            return f'<div class="cost-bar-chart">{rows_html}</div>'

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

        # CSPM recommendations: Healthy/Unhealthy/Not Applicable donut (raw, unfiltered counts)
        rec_state_counts = defaultdict(int)
        for row in real_rows("Recommendations"):
            state = str(row.get("recommendationState", "")).lower()
            rec_state_counts["healthy" if state == "healthy" else "notapplicable" if state == "notapplicable" else "unhealthy"] += 1
        rec_state_present = [s for s in ("healthy", "unhealthy", "notapplicable") if rec_state_counts.get(s, 0) > 0]
        rec_state_labels = {"healthy": {"en": "Healthy", "es": "Saludable", "pt": "Saud\u00e1vel"},
                             "unhealthy": {"en": "Unhealthy", "es": "No saludable", "pt": "N\u00e3o saud\u00e1vel"},
                             "notapplicable": {"en": "Not Applicable", "es": "No aplicable", "pt": "N\u00e3o aplic\u00e1vel"}}
        rec_state_colors = {"healthy": "--cp-success", "unhealthy": "--cp-danger", "notapplicable": "--cp-text-muted"}
        if rec_state_present:
            self._chart_data["cspmStatus"] = {
                "type": "donut", "canvasId": "cspmStatusChart",
                "labels": {lang: [rec_state_labels[s][lang] for s in rec_state_present] for lang in ("en", "es", "pt")},
                "values": [rec_state_counts[s] for s in rec_state_present],
                "colorVars": [rec_state_colors[s] for s in rec_state_present],
            }
        # CSPM findings by severity (open/unresolved recommendations only)
        cspm_severity_counts = defaultdict(int)
        for row in recommendations:
            cspm_severity_counts[str(row.get("recommendationSeverity", "")).lower() or "unknown"] += 1
        cspm_severity_present = [s for s in ("critical", "high", "medium", "low") if cspm_severity_counts.get(s, 0) > 0]
        if cspm_severity_present:
            self._chart_data["cspmSeverity"] = {
                "type": "hbar", "canvasId": "cspmSeverityChart",
                "labels": {lang: [SEVERITY_LABELS.get(s, {}).get(lang, s.capitalize()) for s in cspm_severity_present] for lang in ("en", "es", "pt")},
                "values": [cspm_severity_counts[s] for s in cspm_severity_present],
                "colorVars": [_bare_var(DashboardConfig.SEVERITY_COLORS.get(s)) for s in cspm_severity_present],
            }
        cspm_chart_html = f"""<div class="chart-mini-grid">
            <div><div class="chart-mini-title">{bi('By status', 'Por estado', 'Por status')}</div><div class="chart-card-body" style="height:190px;">{'<canvas id="cspmStatusChart"></canvas>' if rec_state_present else f'<p class="empty-state">{bi_empty("No data.", "Sin datos.")}</p>'}</div></div>
            <div><div class="chart-mini-title">{bi('By severity', 'Por severidad', 'Por severidade')}</div><div class="chart-card-body" style="height:190px;">{'<canvas id="cspmSeverityChart"></canvas>' if cspm_severity_present else f'<p class="empty-state">{bi_empty("No open recommendations.", "Sin recomendaciones abiertas.")}</p>'}</div></div>
        </div>"""

        # Vulnerabilities by severity + exploit status (only using fields already present in the data)
        vuln_severity_counts = defaultdict(int)
        for row in vulnerabilities:
            sev = str(row.get("Severity", "")).lower()
            vuln_severity_counts[sev if sev in ("critical", "high", "medium", "low") else "unknown"] += 1
        vuln_severity_present = [s for s in ("critical", "high", "medium", "low", "unknown") if vuln_severity_counts.get(s, 0) > 0]
        if vuln_severity_present:
            self._chart_data["vulnSeverity"] = {
                "type": "hbar", "canvasId": "vulnSeverityChart",
                "labels": {lang: [SEVERITY_LABELS.get(s, {}).get(lang, s.capitalize()) for s in vuln_severity_present] for lang in ("en", "es", "pt")},
                "values": [vuln_severity_counts[s] for s in vuln_severity_present],
                "colorVars": [_bare_var(DashboardConfig.SEVERITY_COLORS.get(s)) for s in vuln_severity_present],
            }
        has_exploit_fields = any(("ExploitVerified" in row or "PublicExploit" in row) for row in vulnerabilities)
        exploit_counts = defaultdict(int)
        if has_exploit_fields:
            for row in vulnerabilities:
                verified = str(row.get("ExploitVerified", "")).lower() in ("true", "1", "yes")
                public = str(row.get("PublicExploit", "")).lower() in ("true", "1", "yes")
                exploit_counts["verified" if verified else "public" if public else "none"] += 1
        exploit_present = [s for s in ("verified", "public", "none") if exploit_counts.get(s, 0) > 0]
        exploit_labels = {"verified": {"en": "Verified exploit", "es": "Exploit verificado", "pt": "Exploit verificado"},
                           "public": {"en": "Public exploit", "es": "Exploit p\u00fablico", "pt": "Exploit p\u00fablico"},
                           "none": {"en": "No known exploit", "es": "Sin exploit conocido", "pt": "Sem exploit conhecido"}}
        exploit_colors = {"verified": "--cp-danger", "public": "--cp-warning", "none": "--cp-success"}
        if exploit_present:
            self._chart_data["vulnExploit"] = {
                "type": "donut", "canvasId": "vulnExploitChart",
                "labels": {lang: [exploit_labels[s][lang] for s in exploit_present] for lang in ("en", "es", "pt")},
                "values": [exploit_counts[s] for s in exploit_present],
                "colorVars": [exploit_colors[s] for s in exploit_present],
            }
        vuln_chart_html = f"""<div class="chart-mini-grid">
            <div><div class="chart-mini-title">{bi('By severity', 'Por severidad', 'Por severidade')}</div><div class="chart-card-body" style="height:190px;">{'<canvas id="vulnSeverityChart"></canvas>' if vuln_severity_present else f'<p class="empty-state">{bi_empty("No vulnerabilities returned.", "No se devolvieron vulnerabilidades.")}</p>'}</div></div>
            <div><div class="chart-mini-title">{bi('Exploit status', 'Estado del exploit', 'Status do exploit')}</div><div class="chart-card-body" style="height:190px;">{'<canvas id="vulnExploitChart"></canvas>' if exploit_present else f'<p class="empty-state">{bi_empty("No exploit data available.", "No hay datos de exploit disponibles.")}</p>'}</div></div>
        </div>"""

        machines = real_rows("Machines")
        at_risk_machines = len([
            machine for machine in machines
            if str(machine.get("RiskScore", "")).lower() in ("critical", "high")
        ])

        regulatory_standards = real_rows("RegulatoryStandards")
        keyvault_expirations = sorted(
            [row for row in disc_real_rows("KeyVaultExpirations") if row.get("DaysUntilExpiry") not in (None, "")],
            key=lambda row: as_int(row.get("DaysUntilExpiry")),
        )
        expired_kv_items = [row for row in keyvault_expirations if as_int(row.get("DaysUntilExpiry")) < 0]
        app_credential_expiry = sorted(real_rows("AppCredentialExpiry"), key=lambda row: as_int(row.get("DaysUntilExpiry")))
        expired_app_credentials = [row for row in app_credential_expiry if str(row.get("Status", "")).lower() == "expired"]
        guest_users = real_rows("GuestUsers")

        secure_score = breakdown.get("secure_score_pct")
        secure_score_text = f'<span data-secure-pct>{secure_score:.1f}%</span>' if secure_score is not None else '<span data-secure-pct>N/A</span>'
        stats = [
            (secure_score_text, bi("Defender Secure Score", "Secure Score de Defender", "Secure Score do Defender"), "security-accent"),
            (breakdown.get("open_recommendations", 0), bi("Open CSPM Recommendations", "Recomendaciones CSPM abiertas", "Recomendações CSPM abertas"), "security-danger"),
            (breakdown.get("mcsb_failed_resources", 0), bi("Failed MCSB Resources", "Recursos con error en MCSB", "Recursos com falha no MCSB"), "security-warning"),
            (breakdown.get("active_incidents", 0), bi("Active XDR Incidents", "Incidentes XDR activos", "Incidentes XDR ativos"), "security-danger"),
            (breakdown.get("active_alerts", 0), bi("Active Security Alerts", "Alertas de seguridad activas", "Alertas de segurança ativos"), "security-warning"),
            (at_risk_machines, bi("High-risk Endpoints", "Endpoints de alto riesgo", "Endpoints de alto risco"), "security-danger"),
            (len(expired_kv_items) + len(expired_app_credentials), bi("Expired Credentials", "Credenciales expiradas", "Credenciais expiradas"), "security-danger" if (expired_kv_items or expired_app_credentials) else ""),
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
                <thead><tr><th>{bi('Source', 'Fuente', 'Fonte')}</th><th>{bi('Status', 'Estado', 'Status')}</th><th>{bi('Records', 'Registros', 'Registros')}</th><th>{bi('Details', 'Detalles', 'Detalhes')}</th></tr></thead>
                <tbody>{status_rows}</tbody>
            </table></div>"""
        else:
            source_table = f'<p class="empty-state">{bi("Security source status is unavailable.", "El estado de las fuentes de seguridad no está disponible.", "O status das fontes de segurança não está disponível.")}</p>'

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
            ("Subscription", "Subscription", "Suscripción", False),
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
        if endpoint_recommendations:
            def score_bucket(score):
                if score >= 9:
                    return "critical"
                if score >= 7:
                    return "high"
                if score >= 4:
                    return "medium"
                return "low"
            bucket_counts = defaultdict(int)
            for r in endpoint_recommendations:
                bucket_counts[score_bucket(as_int(r.get("SeverityScore")))] += 1
            bucket_present = [b for b in ("critical", "high", "medium", "low") if bucket_counts.get(b, 0) > 0]
            self._chart_data["endpointRecsChart"] = {
                "type": "donut", "canvasId": "endpointRecsChart",
                "labels": {lang: [SEVERITY_LABELS.get(b, {}).get(lang, b.capitalize()) for b in bucket_present] for lang in ("en", "es", "pt")},
                "values": [bucket_counts[b] for b in bucket_present],
                "colorVars": [_bare_var(DashboardConfig.SEVERITY_COLORS.get(b)) for b in bucket_present],
            }
            endpoint_recs_chart_html = '<div class="chart-card-body" style="height:190px;"><canvas id="endpointRecsChart"></canvas></div>'
            endpoint_recommendations_block = chart_table_block(endpoint_recs_chart_html, endpoint_recommendations_table)
        else:
            endpoint_recommendations_block = endpoint_recommendations_table
        vulnerabilities_table = render_table(vulnerabilities, [
            ("VulnerabilityId", "CVE", "CVE", False),
            ("Severity", "Severity", "Severidad", False),
            ("CvssV3", "CVSS v3", "CVSS v3", True),
            ("ExposedMachines", "Exposed Machines", "Equipos expuestos", True),
            ("PublicExploit", "Public Exploit", "Exploit público", False),
        ], "No Defender for Endpoint vulnerabilities returned.", "No se devolvieron vulnerabilidades de Defender for Endpoint.")
        compliance_bars = render_compliance_bars(regulatory_standards)

        recommendations_block = chart_table_block(cspm_chart_html, recommendations_table) if recommendations else recommendations_table
        vulnerabilities_block = chart_table_block(vuln_chart_html, vulnerabilities_table) if vulnerabilities else vulnerabilities_table

        # Secure Score controls: horizontal progress bars (Control Name, Current %, Current/Max)
        if controls:
            controls_bars = ""
            for c in controls:
                current = as_int(c.get("currentScore"))
                maximum = as_int(c.get("maxScore")) or 1
                pct = round(current / maximum * 100) if maximum else 0
                bar_color = "var(--cp-success)" if pct >= 80 else "var(--cp-warning)" if pct >= 50 else "var(--cp-danger)"
                controls_bars += f"""<div class="control-progress-row">
                    <div class="control-progress-label">{escape(str(c.get('controlName', '')))}</div>
                    <div class="control-progress-track"><div class="control-progress-fill" style="width:{pct}%; background:{bar_color};"></div></div>
                    <div class="control-progress-value">{current}/{maximum} ({pct}%)</div>
                </div>"""
            controls_chart_html = f'<div class="control-progress-list">{controls_bars}</div>'
            controls_block = chart_table_block(controls_chart_html, controls_table)
        else:
            controls_block = controls_table

        keyvault_expirations_table = render_table(keyvault_expirations, [
            ("VaultName", "Key Vault", "Key Vault", False),
            ("ItemType", "Type", "Tipo", False),
            ("ItemName", "Item", "Elemento", False),
            ("DaysUntilExpiry", "Days Until Expiry", "D\u00edas para Expirar", True),
            ("ExpiresOn", "Expires On", "Expira el", False),
        ], "No expiring Key Vault secrets or certificates found (or Key Vault Get permission unavailable).", "No se encontraron secretos o certificados de Key Vault por expirar (o falta el permiso Get de Key Vault).")
        app_credential_table = render_table(app_credential_expiry, [
            ("AppDisplayName", "Application", "Aplicaci\u00f3n", False),
            ("CredentialType", "Type", "Tipo", False),
            ("DaysUntilExpiry", "Days Until Expiry", "D\u00edas para Expirar", True),
            ("EndDateTime", "Expires On", "Expira el", False),
            ("Status", "Status", "Estado", False),
        ], "No expiring app registration credentials found.", "No se encontraron credenciales de registros de aplicaciones por expirar.")
        guest_users_table = render_table(guest_users, [
            ("DisplayName", "Guest User", "Usuario invitado", False),
            ("Mail", "Email", "Correo", False),
            ("CreatedDateTime", "Created", "Creado", False),
            ("AccountEnabled", "Enabled", "Habilitado", False),
        ], "No guest (B2B) users found.", "No se encontraron usuarios invitados (B2B).")

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow">MICROSOFT DEFENDER</div>
                <h2>{bi('Security posture and operations', 'Postura y operaciones de seguridad', 'Postura e operações de segurança')}</h2>
                <p>{bi('CSPM, Microsoft Cloud Security Benchmark, Defender XDR, and endpoint exposure in one assessment.', 'CSPM, Microsoft Cloud Security Benchmark, Defender XDR y exposición de endpoints en una sola evaluación.', 'CSPM, Microsoft Cloud Security Benchmark, Defender XDR e exposição de endpoints em uma única avaliação.')}</p>
            </div>
            <div class="security-score-mark">{secure_score_text}</div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title security-title">{bi('CSPM recommendations', 'Recomendaciones CSPM', 'Recomendações CSPM')}{it('Defender for Cloud Cloud Security Posture Management findings from Azure Resource Graph - unresolved recommendations against your subscriptions.', 'Hallazgos de Cloud Security Posture Management de Defender for Cloud, obtenidos vía Azure Resource Graph: recomendaciones sin resolver en tus suscripciones.', 'Descobertas de Cloud Security Posture Management do Defender for Cloud, via Azure Resource Graph: recomendações não resolvidas em suas assinaturas.')}</div>
            {recommendations_block}
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Regulatory and framework compliance', 'Cumplimiento normativo y de marcos', 'Conformidade regulatória e de frameworks')}{it('Compliance against built-in regulatory standards (e.g. ISO 27001, PCI-DSS) and the Microsoft Cloud Security Benchmark (MCSB), as evaluated by Defender for Cloud.', 'Cumplimiento con estándares regulatorios integrados (p.ej. ISO 27001, PCI-DSS) y el Microsoft Cloud Security Benchmark (MCSB), evaluado por Defender for Cloud.', 'Conformidade com padrões regulatórios internos (ex.: ISO 27001, PCI-DSS) e o Microsoft Cloud Security Benchmark (MCSB), avaliada pelo Defender for Cloud.')}</div>
            {compliance_bars}
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Secure Score controls', 'Controles de Secure Score', 'Controles do Secure Score')}</h3>{controls_block}</div>
            <div class="security-panel"><h3>{bi('Microsoft Cloud Security Benchmark', 'Microsoft Cloud Security Benchmark', 'Microsoft Cloud Security Benchmark')}</h3>{mcsb_table}</div>
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Defender for Cloud coverage', 'Cobertura de Defender for Cloud', 'Cobertura do Defender for Cloud')}</h3>{plans_table}</div>
            <div class="security-panel"><h3>{bi('Data source coverage', 'Cobertura de fuentes de datos', 'Cobertura de fontes de dados')}</h3>{source_table}</div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Defender XDR operations', 'Operaciones de Defender XDR', 'Operações do Defender XDR')}{it('Recent security incidents and alerts from Microsoft Defender XDR, via Microsoft Graph - requires a signed-in session with Security Reader access.', 'Incidentes y alertas de seguridad recientes de Microsoft Defender XDR, vía Microsoft Graph - requiere una sesión iniciada con acceso de Security Reader.', 'Incidentes e alertas de segurança recentes do Microsoft Defender XDR, via Microsoft Graph - requer uma sessão autenticada com acesso de Security Reader.')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Recent incidents', 'Incidentes recientes', 'Incidentes recentes')}</h3>{incidents_table}</div>
                <div class="security-panel"><h3>{bi('Recent alerts', 'Alertas recientes', 'Alertas recentes')}</h3>{alerts_table}</div>
            </div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Defender for Endpoint exposure', 'Exposición de Defender for Endpoint', 'Exposição do Defender for Endpoint')}{it('Onboarded devices, endpoint security recommendations and vulnerabilities reported by Defender for Endpoint.', 'Dispositivos incorporados, recomendaciones de seguridad de endpoints y vulnerabilidades reportadas por Defender for Endpoint.', 'Dispositivos integrados, recomendações de segurança de endpoints e vulnerabilidades reportadas pelo Defender for Endpoint.')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Security recommendations', 'Recomendaciones de seguridad', 'Recomendações de segurança')}</h3>{endpoint_recommendations_block}</div>
                <div class="security-panel"><h3>{bi('Vulnerabilities', 'Vulnerabilidades', 'Vulnerabilidades')}</h3>{vulnerabilities_block}</div>
            </div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Credential and identity hygiene', 'Higiene de credenciales e identidad', 'Higiene de credenciais e identidade')}{it('Expiring app registration secrets/certificates and guest (B2B) accounts, sourced from Microsoft Graph.', 'Secretos/certificados de registros de aplicaciones por expirar y cuentas invitadas (B2B), obtenidos de Microsoft Graph.', 'Segredos/certificados de registros de aplicativos a expirar e contas convidadas (B2B), obtidos do Microsoft Graph.')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Key Vault secret/certificate expiration', 'Expiraci\u00f3n de secretos/certificados de Key Vault', 'Expiração de segredos/certificados do Key Vault')}</h3>{keyvault_expirations_table}</div>
                <div class="security-panel"><h3>{bi('App registration credential expiration', 'Expiraci\u00f3n de credenciales de registros de aplicaciones', 'Expiração de credenciais de registro de aplicativos')}</h3>{app_credential_table}</div>
            </div>
            <div class="security-panel">
                <h3>{bi('Guest (B2B) users', 'Usuarios invitados (B2B)', 'Usuários convidados (B2B)')}</h3>
                {guest_users_table}
            </div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Action items', 'Elementos de acción', 'Itens de ação')}</div>
            {self._render_actions_for_pillar("Security")}
        </div>"""

    _IMPLEMENTATION_GUIDES = (
        (r"\b(ddos)\b", "Azure DDoS Protection", "https://learn.microsoft.com/azure/ddos-protection/ddos-protection-overview"),
        (r"\b(defender for cloud|defender cspm|defender for (servers|storage|sql|containers|app service|apis|key vault|resource manager))\b", "Microsoft Defender for Cloud", "https://learn.microsoft.com/azure/defender-for-cloud/defender-for-cloud-introduction"),
        (r"\b(backups?|recovery services|cross region restore|recovery point)\b", "Azure Backup", "https://learn.microsoft.com/azure/backup/backup-overview"),
        (r"\b(diagnostic settings?|diagnostic logs?|log analytics|azure monitor based alerts|container monitoring|application health monitoring)\b", "Azure Monitor", "https://learn.microsoft.com/azure/azure-monitor/fundamentals/overview"),
        (r"\b(policy assignment|policy evaluation|policy compliance|resource locks?|resources? (?:have|has) no tags)\b", "Azure Policy", "https://learn.microsoft.com/azure/governance/policy/overview"),
        (r"\b(orphaned|unattached|deallocated|underutilized|unused/empty|reservation/savings plan|reserved instance|savings plan)\b", "Azure Cost Management", "https://learn.microsoft.com/azure/cost-management-billing/costs/overview-cost-management"),
        (r"\b(postgresql|postgres)\b", "Azure Database for PostgreSQL", "https://learn.microsoft.com/azure/well-architected/service-guides/postgresql"),
        (r"\b(cosmos db|cosmosdb)\b", "Azure Cosmos DB", "https://learn.microsoft.com/azure/well-architected/service-guides/cosmos-db"),
        (r"\b(azure sql|sql database|sql server|sql managed instance)\b", "Azure SQL", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-sql-database"),
        (r"\b(aks|kubernetes|kubernetes service)\b", "Azure Kubernetes Service", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-kubernetes-service"),
        (r"\b(app service|web app)\b", "Azure App Service", "https://learn.microsoft.com/azure/well-architected/service-guides/app-service-web-apps"),
        (r"\b(application gateway)\b", "Azure Application Gateway", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-application-gateway"),
        (r"\b(load balancer)\b", "Azure Load Balancer", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-load-balancer"),
        (r"\b(azure firewall|firewall policy)\b", "Azure Firewall", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-firewall"),
        (r"\b(front door)\b", "Azure Front Door", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-front-door"),
        (r"\b(service bus)\b", "Azure Service Bus", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-service-bus"),
        (r"\b(event hub|event hubs)\b", "Azure Event Hubs", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-event-hubs"),
        (r"\b(api management)\b", "Azure API Management", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-api-management"),
        (r"\b(azure function|function app)\b", "Azure Functions", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-functions"),
        (r"\b(azure managed redis|azure cache for redis)\b", "Azure Managed Redis", "https://learn.microsoft.com/azure/redis/overview"),
        (r"\b(key vault)\b", "Azure Key Vault", "https://learn.microsoft.com/azure/key-vault/general/overview"),
        (r"\b(storage accounts?|blob storage)\b", "Azure Blob Storage", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-blob-storage"),
        (r"\b(managed disk|data disk|premium ssd)\b", "Azure Disk Storage", "https://learn.microsoft.com/azure/well-architected/service-guides/azure-disk-storage"),
        (r"\b(virtual network|vnet)\b", "Azure Virtual Network", "https://learn.microsoft.com/azure/well-architected/service-guides/virtual-network"),
        (r"(?:\b(?:virtual machine|virtual machines|vms|vmss)\b|vm\(s\))", "Azure Virtual Machines", "https://learn.microsoft.com/azure/well-architected/service-guides/virtual-machines"),
    )

    _PILLAR_GUIDES = {
        "Reliability": "https://learn.microsoft.com/azure/well-architected/reliability/",
        "Security": "https://learn.microsoft.com/azure/well-architected/security/",
        "Cost Optimization": "https://learn.microsoft.com/azure/well-architected/cost-optimization/",
        "Operational Excellence": "https://learn.microsoft.com/azure/well-architected/operational-excellence/",
        "Performance Efficiency": "https://learn.microsoft.com/azure/well-architected/performance-efficiency/",
    }

    _VISIBLE_RECOMMENDATIONS = 4

    @staticmethod
    def _affected_count(action: dict) -> int:
        title = str(action.get("title", ""))
        match = re.match(r"^\s*(\d+)(?:/\d+)?\b", title)
        if match:
            return int(match.group(1))
        return max(1, len(action.get("resources", [])))

    def _implementation_guide(self, action: dict) -> dict:
        searchable = " ".join(str(action.get(field, "")) for field in ("title", "description")).lower()
        for pattern, service, url in self._IMPLEMENTATION_GUIDES:
            if re.search(pattern, searchable, re.IGNORECASE):
                return {"service": service, "url": url, "confidence": "direct"}
        pillar = str(action.get("pillar", "Operational Excellence"))
        return {
            "service": f"{pillar} architecture review",
            "url": self._PILLAR_GUIDES.get(pillar, "https://learn.microsoft.com/azure/well-architected/"),
            "confidence": "review",
        }

    @staticmethod
    def _normalized_title(action: dict) -> str:
        title = re.sub(r"^\d+(?:/\d+)?\s+(?:resource\(s\) affected:\s*)?", "", str(action.get("title", "")).strip().lower())
        return re.sub(r"\s+", " ", title)

    def _recommend_azure_resources(self, limit: int = 12) -> list:
        """Return the highest-priority observed remediations with optional implementation guidance.

        The first `limit` cards are chosen by severity and affected scope (at most three per pillar).
        On top of that, each pillar reserves one extra slot for its best remaining high/critical
        finding, so a serious finding with a small scope is not crowded out by high-volume ones.
        Advisor labels most findings "High", so the reserved slot is ranked by severity, then
        confirmed internet exposure, then findings measured by this toolkit before generic Advisor ones.
        """
        severity_rank = {"critical": 4, "high": 3, "medium": 2, "low": 1, "info": 0}

        def severity_of(item: dict) -> int:
            return severity_rank.get(str(item.get("severity", "")).lower(), 0)

        ranked = sorted(
            self.engine.action_items,
            key=lambda item: (-severity_of(item), -self._affected_count(item), str(item.get("title", ""))),
        )

        recommendations = []
        pillar_counts = defaultdict(int)
        seen = set()
        for action in ranked:
            normalized = self._normalized_title(action)
            if not normalized or normalized in seen:
                continue
            pillar = str(action.get("pillar", "Operational Excellence"))
            if pillar_counts[pillar] >= 3:
                continue
            seen.add(normalized)
            pillar_counts[pillar] += 1
            recommendations.append({
                "action": action,
                "affected": self._affected_count(action),
                "guide": self._implementation_guide(action),
                "reserved": False,
            })
            if len(recommendations) >= limit:
                break

        reserved_ranked = sorted(
            (item for item in self.engine.action_items if severity_of(item) >= 3),
            key=lambda item: (
                -severity_of(item),
                not item.get("internet_exposed", False),
                str(item.get("source", "")) == "Azure Advisor",
                -self._affected_count(item),
                str(item.get("title", "")),
            ),
        )
        reserved_pillars = set()
        for action in reserved_ranked:
            normalized = self._normalized_title(action)
            pillar = str(action.get("pillar", "Operational Excellence"))
            if not normalized or normalized in seen or pillar in reserved_pillars:
                continue
            seen.add(normalized)
            reserved_pillars.add(pillar)
            recommendations.append({
                "action": action,
                "affected": self._affected_count(action),
                "guide": self._implementation_guide(action),
                "reserved": True,
            })
        return recommendations

    def _render_azure_resource_recommendations(self) -> str:
        recommendations = self._recommend_azure_resources()
        if not recommendations:
            return f'<p class="empty-state">{bi("No additional Azure resources are indicated by the current findings.", "Los hallazgos actuales no indican recursos adicionales de Azure.", "Os achados atuais n\u00e3o indicam recursos adicionais do Azure.")}</p>'

        priority_es = {"high": "Alta", "medium": "Media", "low": "Baja"}
        priority_pt = {"high": "Alta", "medium": "Média", "low": "Baixa"}
        cards = []
        for index, recommendation in enumerate(recommendations):
            action = recommendation["action"]
            guide = recommendation["guide"]
            hidden_attr = ' hidden' if index >= self._VISIBLE_RECOMMENDATIONS else ''
            extra_class = ' extra-row' if index >= self._VISIBLE_RECOMMENDATIONS else ''
            priority = str(action.get("severity", "medium")).lower()
            if priority not in priority_es:
                priority = "high" if priority == "critical" else "medium"
            resources = []
            for resource in action.get("resources", [])[:3]:
                if isinstance(resource, dict):
                    resource = resource.get("name") or resource.get("id") or ""
                if resource:
                    resources.append(str(resource))
            evidence_items = (
                f"<li>{bi('Source', 'Fuente', 'Fonte')}: {escape(str(action.get('source', 'Assessment analysis')))}</li>"
                f"<li>{bi('WAF pillar', 'Pilar WAF', 'Pilar WAF')}: {escape(str(action.get('pillar', '')))}</li>"
            )
            if recommendation.get("reserved"):
                evidence_items += (
                    f"<li>{bi('Reserved slot: high severity, smaller scope', 'Espacio reservado: severidad alta, alcance menor', 'Espaço reservado: severidade alta, escopo menor')}</li>"
                )
            evidence_items += "".join(
                f"<li>{bi('Example resource', 'Recurso de ejemplo', 'Recurso de exemplo')}: {escape(resource)}</li>"
                for resource in resources
            )
            confidence_text = bi(
                "Direct service match" if guide["confidence"] == "direct" else "Architecture review required",
                "Asociación directa al servicio" if guide["confidence"] == "direct" else "Requiere revisión de arquitectura",
                "Associação direta ao serviço" if guide["confidence"] == "direct" else "Requer revisão de arquitetura",
            )
            subscriptions_attr = f' data-subscriptions="{escape(json.dumps(action["subscriptions"], ensure_ascii=False))}"' if action.get("subscriptions") else ""
            cards.append(f"""<article class="resource-rec-card resource-rec-{priority}{extra_class}"{hidden_attr}{subscriptions_attr}>
                <div class="resource-rec-heading">
                    <div>
                        <span class="resource-rec-priority">{bi(priority.title() + ' priority', 'Prioridad ' + priority_es[priority], 'Prioridade ' + priority_pt[priority])}</span>
                        <h3>{bi(escape(str(action.get('title', ''))), escape(str(action.get('title_es', action.get('title', '')))), escape(str(action.get('title_pt', action.get('title', '')))))}</h3>
                    </div>
                    <span class="resource-rec-count">{recommendation['affected']} {bi('affected', 'afectados', 'afetados')}</span>
                </div>
                <p>{bi(escape(str(action.get('description', ''))), escape(str(action.get('description_es', action.get('description', '')))), escape(str(action.get('description_pt', action.get('description', '')))))}</p>
                <div class="resource-rec-implementation">
                    <span>{bi('Suggested implementation', 'Implementación sugerida', 'Implementação sugerida')}</span>
                    <strong>{escape(guide['service'])}</strong>
                    <small>{confidence_text}</small>
                </div>
                <div class="resource-rec-evidence-label">{bi('Evidence', 'Evidencia', 'Evidência')}</div>
                <ul>{evidence_items}</ul>
                <a class="resource-rec-link" href="{escape(guide['url'])}" target="_blank" rel="noopener">{bi('Review official guidance', 'Revisar guía oficial', 'Revisar orientação oficial')} &rarr;</a>
            </article>""")
        grid = f'<div class="resource-rec-grid" data-visible="{self._VISIBLE_RECOMMENDATIONS}">{"".join(cards)}</div>'
        hidden_count = len(recommendations) - self._VISIBLE_RECOMMENDATIONS
        if hidden_count <= 0:
            return grid
        count = f'<span class="rec-more-count">{hidden_count}</span>'
        more_text = bi(f"Show {count} more", f"Mostrar {count} m\u00e1s", f"Mostrar mais {count}")
        less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
        return (
            f'{grid}<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
            f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
        )

    def _render_governance_view(self) -> str:
        """Render management-group hierarchy, policy/RBAC, and scope details natively (no iframe to a separate HTML file)."""
        gov = self.data.get("governance", {})

        def real_rows(sheet_name):
            return [
                row for row in gov.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def render_table(rows, columns, empty_en, empty_es, limit=15):
            if not rows:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            MAX_TABLE_ROWS = 300
            row_total = len(rows)
            rows = rows[:MAX_TABLE_ROWS]
            header = "".join(f"<th>{bi(label_en, label_es, LABEL_PT.get(label_en, label_en))}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for i, row in enumerate(rows):
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{escape(str(row.get(key, "")))}</td>'
                    for key, _, _, numeric in columns
                )
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                row_attr += _resource_link_attr(row)
                body += f"<tr{row_attr}>{cells}</tr>"
            more = ""
            if len(rows) > limit:
                more_text = bi(f"+ {len(rows) - limit} more", f"+ {len(rows) - limit} m\u00e1s", f"+ {len(rows) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            if row_total > MAX_TABLE_ROWS:
                more += f'<p class="table-truncated-note">{bi(f"Showing first {MAX_TABLE_ROWS} of {row_total} rows for performance - see the source export for the full list.", f"Mostrando las primeras {MAX_TABLE_ROWS} de {row_total} filas por rendimiento - consulte el archivo de origen para ver la lista completa.", f"Mostrando as primeiras {MAX_TABLE_ROWS} de {row_total} linhas por desempenho - consulte o arquivo de origem para a lista completa.")}</p>'
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>{more}'

        mg_rows = real_rows("MgmtGroups")
        subs_rows = real_rows("Subscriptions")
        hierarchy_inferred = False
        if not mg_rows:
            subscriptions_by_parent = defaultdict(list)
            for subscription in subs_rows:
                parent_name = str(subscription.get("mgParent", "") or "").strip()
                if parent_name:
                    subscriptions_by_parent[parent_name].append(subscription)
            if subscriptions_by_parent:
                hierarchy_inferred = True
                mg_rows = [
                    {
                        "Level": 0,
                        "DisplayName": parent_name,
                        "Id": f"inferred-{hashlib.sha1(parent_name.encode('utf-8')).hexdigest()[:12]}",
                        "Path": parent_name,
                        "ChildMGs": 0,
                        "Subscriptions": len(subscriptions),
                        "SubNames": "; ".join(
                            str(subscription.get("name", subscription.get("subscriptionId", "")))
                            for subscription in subscriptions
                        ),
                        "ParentId": "",
                        "Inferred": True,
                    }
                    for parent_name, subscriptions in sorted(subscriptions_by_parent.items())
                ]
        policy_assignments = real_rows("PolicyAssignments")
        custom_policies = real_rows("CustomPolicies")
        policy_compliance = real_rows("PolicyCompliance")
        role_assignments = real_rows("RoleAssignments")
        custom_roles = real_rows("CustomRoles")
        defender_plans = real_rows("DefenderPlans")
        secure_scores = real_rows("SecureScores")
        resource_summary = real_rows("ResourceSummary")
        orphaned = real_rows("OrphanedResources")
        vnets = real_rows("VNets")
        locks = real_rows("Locks")
        policy_exemptions = real_rows("PolicyExemptions")

        non_compliant_count = sum(int(r.get("NonCompliantCount", 0) or 0) for r in policy_compliance) or len(policy_compliance)
        stats = [
            (len(mg_rows), bi("Management Groups", "Grupos de Administraci\u00f3n", "Grupos de Gerenciamento")),
            (len(subs_rows), bi("Subscriptions", "Suscripciones", "Assinaturas")),
            (len(policy_assignments), bi("Policy Assignments", "Asignaciones de Pol\u00edticas", "Atribui\u00e7\u00f5es de Pol\u00edtica")),
            (non_compliant_count, bi("Non-Compliant Resources", "Recursos No Conformes", "Recursos N\u00e3o Conformes")),
            (len(role_assignments), bi("Role Assignments", "Asignaciones de Roles", "Atribui\u00e7\u00f5es de Fun\u00e7\u00e3o")),
            (len(locks), bi("Resource Locks", "Bloqueos de Recursos", "Bloqueios de Recursos")),
            (len(orphaned), bi("Orphaned Resources", "Recursos Hu\u00e9rfanos", "Recursos \u00d3rf\u00e3os")),
            (len(policy_exemptions), bi("Policy Exemptions", "Exenciones de Pol\u00edtica", "Isen\u00e7\u00f5es de Pol\u00edtica")),
        ]
        stat_html = "".join(
            f'<div class="security-stat"><div class="value security-accent">{count:,}</div><div class="label">{label}</div></div>'
            for count, label in stats
        )

        def mg_level(m):
            try:
                return int(m.get("Level"))
            except (TypeError, ValueError):
                return None

        has_hierarchy = bool(mg_rows) and any(mg_level(m) is not None for m in mg_rows)

        subs_by_mgname: dict = {}
        for s in subs_rows:
            subs_by_mgname.setdefault(str(s.get("mgParent", "")), []).append(str(s.get("subscriptionId", "")))

        _descendant_sub_ids_cache: dict = {}

        def descendant_sub_ids(mg):
            mg_id = str(mg.get("Id", ""))
            if mg_id in _descendant_sub_ids_cache:
                return _descendant_sub_ids_cache[mg_id]
            ids = set(subs_by_mgname.get(str(mg.get("DisplayName", "")), []))
            level = mg_level(mg)
            level = level if level is not None else 0
            path = str(mg.get("Path", "") or "")
            children = [m for m in mg_rows if mg_level(m) == level + 1 and str(m.get("Path", "")).startswith(f"{path}/")]
            for child in children:
                ids |= descendant_sub_ids(child)
            _descendant_sub_ids_cache[mg_id] = ids
            return ids

        def scope_matches(row, mg):
            mg_id = str(mg.get("Id", ""))
            scope = str(row.get("scope_", "") or "")
            if mg_id and f"/managementgroups/{mg_id}".lower() in scope.lower():
                return True
            sub_id = str(row.get("subscriptionId", "") or "")
            return bool(sub_id) and sub_id in descendant_sub_ids(mg)

        subs_rows_by_mgname: dict = {}
        for s in subs_rows:
            subs_rows_by_mgname.setdefault(str(s.get("mgParent", "")), []).append(s)

        # Resource Groups + individual resources come from the Discovery phase, not Governance
        # (Governance's own "ResourceSummary" sheet is just aggregate type+location counts).
        discovery_data = self.data.get("discovery", {})

        def disc_real_rows(sheet_name):
            return [
                row for row in discovery_data.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        rg_rows_by_sub: dict = {}
        for rg in disc_real_rows("ResourceGroups"):
            rg_rows_by_sub.setdefault(str(rg.get("subscriptionId", "")), []).append(rg)

        _RESOURCE_SHEET_SKIP = {"Subscriptions", "ResourceGroups", "ResourceSummary"}
        resources_by_rg: dict = {}
        for sheet_name in discovery_data.keys():
            if sheet_name in _RESOURCE_SHEET_SKIP:
                continue
            for row in disc_real_rows(sheet_name):
                rg_name = row.get("resourceGroup")
                sub_id = row.get("subscriptionId")
                res_name = row.get("name") or row.get("friendlyName")
                if not rg_name or not sub_id or not res_name:
                    continue
                key = (str(sub_id), str(rg_name).lower())
                resources_by_rg.setdefault(key, []).append({"name": str(res_name), "kind": sheet_name})

        MAX_RESOURCES_PER_RG = 15
        NODE_SIZES = {
            "mg": (200, 78),
            "sub": (170, 72),
            "rg": (160, 64),
            "resource": (150, 50),
            "more": (130, 40),
        }
        DIAGRAM_H_GAP, DIAGRAM_V_GAP = 28, 46
        # MG/Sub form the horizontal "parallel" org-chart row (SVG); RGs/Resources are rendered
        # separately as a plain HTML collapsible tree below (see render_resource_explorer) so an
        # expanded subtree can never overlap the SVG-positioned MG/Sub boxes or their siblings.
        HORIZONTAL_KINDS = {"mg", "sub"}
        _diagram_id_seq = [0]

        def next_diagram_id():
            _diagram_id_seq[0] += 1
            return f"dgnode{_diagram_id_seq[0]}"

        def diagram_children(node, depth_guard=0):
            if depth_guard > 15:
                return []
            kind = node["kind"]
            data = node["data"]
            if kind == "mg":
                level = mg_level(data)
                level = level if level is not None else 0
                path = str(data.get("Path", "") or "")
                children = [
                    {"kind": "mg", "data": m, "children": []}
                    for m in mg_rows
                    if mg_level(m) == level + 1 and str(m.get("Path", "")).startswith(f"{path}/")
                ] + [
                    {"kind": "sub", "data": s, "children": []}
                    for s in subs_rows_by_mgname.get(str(data.get("DisplayName", "")), [])
                ]
            elif kind == "sub":
                sub_id = str(data.get("subscriptionId", ""))
                children = [{"kind": "rg", "data": rg, "children": []} for rg in rg_rows_by_sub.get(sub_id, [])]
            elif kind == "rg":
                sub_id = str(data.get("subscriptionId", ""))
                rg_name = str(data.get("name", ""))
                res_list = resources_by_rg.get((sub_id, rg_name.lower()), [])
                children = [{"kind": "resource", "data": r, "children": []} for r in res_list[:MAX_RESOURCES_PER_RG]]
                if len(res_list) > MAX_RESOURCES_PER_RG:
                    children.append({"kind": "more", "data": {"count": len(res_list) - MAX_RESOURCES_PER_RG}, "children": []})
            else:
                children = []
            for c in children:
                c["children"] = diagram_children(c, depth_guard + 1)
            return children

        def node_width(node):
            return NODE_SIZES.get(node["kind"], (170, 60))[0]

        def node_height(node):
            return NODE_SIZES.get(node["kind"], (170, 60))[1]

        def layout_subtree_width(node):
            w = node_width(node)
            children = [c for c in node["children"] if c["kind"] in HORIZONTAL_KINDS]
            if not children:
                node["_w"] = w
                return w
            total = sum(layout_subtree_width(c) for c in children) + DIAGRAM_H_GAP * (len(children) - 1)
            node["_w"] = max(w, total)
            return node["_w"]

        def assign_x(node, left):
            children = [c for c in node["children"] if c["kind"] in HORIZONTAL_KINDS]
            if not children:
                node["_x"] = left + node["_w"] / 2
                return
            children_total = sum(c["_w"] for c in children) + DIAGRAM_H_GAP * (len(children) - 1)
            cursor = left + (node["_w"] - children_total) / 2
            for c in children:
                assign_x(c, cursor)
                cursor += c["_w"] + DIAGRAM_H_GAP
            node["_x"] = left + node["_w"] / 2

        def assign_ids(node):
            node["_id"] = next_diagram_id()
            for c in node["children"]:
                assign_ids(c)

        def collect_sub_nodes(node, out):
            if node["kind"] == "sub":
                out.append(node)
            for c in node["children"]:
                collect_sub_nodes(c, out)

        def build_diagram(typed_roots):
            # typed_roots are already {"kind": ..., "data": ...} dicts - either "mg" (real MG
            # hierarchy) or "sub" (fallback when no MG read permission exists), so subscriptions
            # can still anchor the tree/resource explorer even without a Management Group scope.
            tree_roots = [{"kind": r["kind"], "data": r["data"], "children": []} for r in typed_roots]
            for root in tree_roots:
                root["children"] = diagram_children(root)
            if not tree_roots:
                return "", []

            for root in tree_roots:
                assign_ids(root)
            sub_nodes = []
            for root in tree_roots:
                collect_sub_nodes(root, sub_nodes)

            cursor = 0
            for i, root in enumerate(tree_roots):
                if i:
                    cursor += DIAGRAM_H_GAP * 2
                layout_subtree_width(root)
                assign_x(root, cursor)
                cursor += root["_w"]
            total_w = cursor

            max_depth = 0
            flat = []

            def collect(node, depth):
                nonlocal max_depth
                max_depth = max(max_depth, depth)
                flat.append((node, depth))
                for c in node["children"]:
                    if c["kind"] in HORIZONTAL_KINDS:
                        collect(c, depth + 1)

            for root in tree_roots:
                collect(root, 0)

            depth_height = {}
            for node, depth in flat:
                depth_height[depth] = max(depth_height.get(depth, 0), node_height(node))

            depth_top = {}
            cursor_y = 24
            for d in range(max_depth + 1):
                depth_top[d] = cursor_y
                cursor_y += depth_height.get(d, 0) + DIAGRAM_V_GAP
            diagram_h = cursor_y + 12

            def center_y(depth):
                return depth_top[depth] + depth_height.get(depth, 0) / 2

            for node, depth in flat:
                node["_y"] = center_y(depth)

            diagram_w = max(total_w, NODE_SIZES["mg"][0]) + 40

            def render_node_box(node):
                kind = node["kind"]
                data = node["data"]
                w, h = node_width(node), node_height(node)
                x, y = node["_x"], node["_y"]
                onclick_attr = ""
                icon = ""
                expand_html = ""
                if kind == "mg":
                    mg_id = str(data.get("Id", ""))
                    p_count = len([p for p in policy_assignments if scope_matches(p, data)])
                    r_count = len([r for r in role_assignments if scope_matches(r, data)])
                    name = str(data.get("DisplayName", mg_id) or mg_id)
                    icon = self._azure_icon("management-groups")
                    onclick_js = (
                        "var el=document.getElementById('mg-scope-" + escape(mg_id) + "'); "
                        "if(el){el.open=true; el.scrollIntoView({behavior:'smooth', block:'center'});}"
                    )
                    onclick_attr = f' onclick="{onclick_js}" style="cursor:pointer;"'
                    box_class = "mg-diagram-node mg-diagram-node-mg"
                    badges = (
                        f'<span class="mg-badge mg-badge-policy" title="{p_count} policy assignments">{p_count}</span>'
                        f'<span class="mg-badge mg-badge-rbac" title="{r_count} role assignments">{r_count}</span>'
                    )
                elif kind == "sub":
                    sub_id = str(data.get("subscriptionId", ""))
                    p_count = len([p for p in policy_assignments if str(p.get("subscriptionId", "")) == sub_id])
                    r_count = len([r for r in role_assignments if str(r.get("subscriptionId", "")) == sub_id])
                    name = str(data.get("name", sub_id) or sub_id)
                    icon = self._azure_icon("subscriptions")
                    box_class = "mg-diagram-node mg-diagram-node-sub"
                    badges = (
                        f'<span class="mg-badge mg-badge-policy" title="{p_count} policy assignments">{p_count}</span>'
                        f'<span class="mg-badge mg-badge-rbac" title="{r_count} role assignments">{r_count}</span>'
                    )
                    if node["children"]:
                        sub_id_esc = escape(sub_id)
                        onclick_js = (
                            "var el=document.getElementById('rg-explorer-" + sub_id_esc + "'); "
                            "if(el){el.open=true; el.scrollIntoView({behavior:'smooth', block:'center'});}"
                        )
                        onclick_attr = f' onclick="{onclick_js}" style="cursor:pointer;"'
                        expand_html = (
                            f'<span class="mg-diagram-expand">'
                            f'<span class="mg-diagram-expand-icon">\u25b8</span> {len(node["children"])} '
                            f'{bi("RGs", "GR", "GRs")}</span>'
                        )
                else:  # "rg"/"resource"/"more" kinds are only used by the HTML resource explorer, not the SVG
                    return ""

                label = name
                avail_px = w - (16 if icon else 0) - 6 - 20
                max_chars = max(6, int(avail_px / 6.6))
                if len(label) > max_chars:
                    label = label[: max_chars - 1] + "\u2026"
                return (
                    f'<foreignObject x="{x - w / 2:.1f}" y="{y - h / 2:.1f}" width="{w}" height="{h}">'
                    f'<div xmlns="http://www.w3.org/1999/xhtml" class="{box_class}" title="{escape(name)}"{onclick_attr}>'
                    f'<div class="mg-diagram-name">{icon}<span>{escape(label)}</span></div>'
                    f'<div class="mg-diagram-badges">{badges}</div>{expand_html}</div></foreignObject>'
                )

            def render_subtree(node):
                parts = [render_node_box(node)]
                for child in node["children"]:
                    if child["kind"] not in HORIZONTAL_KINDS:
                        continue
                    y_bottom = node["_y"] + node_height(node) / 2
                    y_top = child["_y"] - node_height(child) / 2
                    mid_y = (y_bottom + y_top) / 2
                    link = (
                        f'<path d="M {node["_x"]:.1f},{y_bottom:.1f} C {node["_x"]:.1f},{mid_y:.1f} '
                        f'{child["_x"]:.1f},{mid_y:.1f} {child["_x"]:.1f},{y_top:.1f}" class="mg-diagram-link" fill="none" />'
                    )
                    parts.append(link + render_subtree(child))
                return "".join(parts)

            body = "".join(render_subtree(root) for root in tree_roots)
            svg = (
                f'<svg class="mg-diagram-svg" viewBox="0 0 {diagram_w:.0f} {diagram_h:.0f}" '
                f'width="{diagram_w:.0f}" height="{diagram_h:.0f}">{body}</svg>'
            )
            return f'<div class="mg-diagram-wrap">{svg}</div>', sub_nodes

        def render_resource_explorer(sub_nodes):
            items = []
            for sub in sub_nodes:
                if not sub["children"]:
                    continue
                sub_id = str(sub["data"].get("subscriptionId", ""))
                sub_name = str(sub["data"].get("name", sub_id) or sub_id)
                rg_items = []
                for rg in sub["children"]:
                    rg_name = str(rg["data"].get("name", "") or "")
                    loc = str(rg["data"].get("location", "") or "")
                    real = [c for c in rg["children"] if c["kind"] == "resource"]
                    more_c = next((c for c in rg["children"] if c["kind"] == "more"), None)
                    res_count = len(real) + (more_c["data"]["count"] if more_c else 0)
                    res_lines = "".join(
                        f'<li>{self._resource_kind_icon(str(r["data"]["kind"]))}'
                        f'<span class="resource-name">{escape(str(r["data"]["name"]))}</span>'
                        f'<span class="mg-diagram-restype">{escape(str(r["data"]["kind"]))}</span></li>'
                        for r in real
                    )
                    if more_c:
                        res_lines += f'<li class="resource-more">+ {more_c["data"]["count"]} {bi("more", "m\u00e1s", "mais")}</li>'
                    body = (
                        f'<ul class="rgx-resource-list">{res_lines}</ul>' if res_lines
                        else f'<p class="empty-state">{bi("No resources found.", "No se encontraron recursos.", "Nenhum recurso encontrado.")}</p>'
                    )
                    loc_badge = f'<span class="scope-id">{escape(loc)}</span>' if loc else ""
                    rg_items.append(f"""<details class="rg-explorer-item">
                        <summary>{self._azure_icon('resource-groups')} <strong>{escape(rg_name)}</strong> {loc_badge}
                            <span class="mg-badge mg-badge-subs">{res_count} {bi('resources', 'recursos', 'recursos')}</span>
                        </summary>
                        {body}
                    </details>""")
                items.append(f"""<details class="rg-explorer" id="rg-explorer-{escape(sub_id)}">
                    <summary>{self._azure_icon('subscriptions')} <strong>{escape(sub_name)}</strong></summary>
                    <div class="rg-explorer-list">{''.join(rg_items)}</div>
                </details>""")
            return "".join(items)

        if has_hierarchy:
            roots = [{"kind": "mg", "data": m, "children": []} for m in mg_rows if mg_level(m) == 0]
            hierarchy_html, sub_nodes_for_explorer = build_diagram(roots)
        elif subs_rows:
            # No usable Management Group hierarchy (commonly: only subscription-scoped Reader
            # access, no MG read permission) - anchor the tree on each subscription directly so
            # the map and resource explorer still populate from real Sub/RG/resource data.
            roots = [{"kind": "sub", "data": s, "children": []} for s in subs_rows]
            hierarchy_html, sub_nodes_for_explorer = build_diagram(roots)
        else:
            hierarchy_html, sub_nodes_for_explorer = "", []
        resource_explorer_html = render_resource_explorer(sub_nodes_for_explorer)
        if not resource_explorer_html:
            resource_explorer_html = f'<p class="empty-state">{bi("No resource group data found (requires the Discovery phase).", "No se encontraron datos de grupos de recursos (requiere la fase de Discovery).", "Nenhum dado de grupo de recursos encontrado (requer a fase de Discovery).")}</p>'
        if not hierarchy_html:
            hierarchy_html = render_table(mg_rows, [
                ("displayName", "Name", "Nombre", False),
                ("name", "Id", "Id", False),
                ("parent", "Parent", "Padre", False),
            ], "No management group data found.", "No se encontraron datos de grupos de administraci\u00f3n.")
        hierarchy_note_html = ""
        if hierarchy_inferred:
            hierarchy_note_text = bi(
                "Management Group nodes were inferred from each subscription's immediate mgParent value because the tenant-root hierarchy was not available.",
                "Los nodos de Grupos de Administración se infirieron del valor mgParent inmediato de cada suscripción porque la jerarquía del tenant root no estaba disponible.",
                "Os nós de Grupos de Gerenciamento foram inferidos do valor mgParent imediato de cada assinatura porque a hierarquia da raiz do tenant não estava disponível.",
            )
            hierarchy_note_html = f'<p class="table-truncated-note">{hierarchy_note_text}</p>'

        scope_items = []
        if has_hierarchy:
            ordered_mgs = sorted(mg_rows, key=lambda m: (mg_level(m) or 0, str(m.get("DisplayName", ""))))
            for mg in ordered_mgs[:40]:
                mg_id = str(mg.get("Id", ""))
                mg_policies = [p for p in policy_assignments if scope_matches(p, mg)]
                mg_roles = [r for r in role_assignments if scope_matches(r, mg)]
                indent = "\u2014" * (mg_level(mg) or 0)
                if mg_policies:
                    policy_rows = "".join(
                        f"<tr><td>{escape(str(p.get('displayName', '')))}</td><td>{escape(str(p.get('enforcement', '')))}</td><td>{escape(str(p.get('identity_', '')))}</td></tr>"
                        for p in mg_policies[:10]
                    )
                    policy_table = f"""<table class="mini-table"><thead><tr><th>{bi('Policy', 'Pol\u00edtica', 'Pol\u00edtica')}</th><th>{bi('Enforcement', 'Aplicaci\u00f3n', 'Aplica\u00e7\u00e3o')}</th><th>{bi('Identity', 'Identidad', 'Identidade')}</th></tr></thead><tbody>{policy_rows}</tbody></table>"""
                else:
                    policy_table = f'<p class="empty-state">{bi("No policy assignments at this scope.", "Sin asignaciones de pol\u00edticas en este alcance.", "Sem atribui\u00e7\u00f5es de pol\u00edtica neste escopo.")}</p>'
                scope_items.append(f"""<details class="scope-detail" id="mg-scope-{escape(mg_id)}">
                    <summary>{indent} <strong>{escape(str(mg.get('DisplayName', mg_id)))}</strong> <span class="scope-id">({escape(mg_id)})</span>
                        <span class="mg-badge mg-badge-policy">{len(mg_policies)} {bi('policies', 'pol.', 'pol.')}</span>
                        <span class="mg-badge mg-badge-rbac">{len(mg_roles)} {bi('roles', 'roles', 'fun\u00e7\u00f5es')}</span>
                        <span class="mg-badge mg-badge-subs">{mg.get('Subscriptions', 0)} {bi('subs', 'subs', 'assin.')}</span>
                    </summary>
                    <p><strong>{bi('Path', 'Ruta', 'Caminho')}:</strong> {escape(str(mg.get('Path', '')))}</p>
                    {policy_table}
                </details>""")
        scope_insights_html = "".join(scope_items) if scope_items else f'<p class="empty-state">{bi("Scope insights require the Management Group hierarchy (Level/Path fields).", "Los detalles de alcance requieren la jerarqu\u00eda de Grupos de Administraci\u00f3n (campos Level/Path).", "Os detalhes de escopo exigem a hierarquia de Grupos de Gerenciamento (campos Level/Path).")}</p>'

        policy_assignments_table = render_table(policy_assignments, [
            ("displayName", "Name", "Nombre", False),
            ("enforcement", "Enforcement", "Aplicaci\u00f3n", False),
            ("identity_", "Identity", "Identidad", False),
            ("scope_", "Scope", "Alcance", False),
        ], "No policy assignments found.", "No se encontraron asignaciones de pol\u00edticas.")
        custom_policies_table = render_table(custom_policies, [
            ("displayName", "Name", "Nombre", False),
            ("effect", "Effect", "Efecto", False),
            ("category", "Category", "Categor\u00eda", False),
        ], "No custom policy definitions found.", "No se encontraron definiciones de pol\u00edticas personalizadas.")
        policy_compliance_table = render_table(policy_compliance, [
            ("policyAssignment", "Policy Assignment", "Asignaci\u00f3n de Pol\u00edtica", False),
            ("complianceState", "State", "Estado", False),
            ("NonCompliantCount", "Non-Compliant", "No Conformes", True),
        ], "No non-compliant resources found.", "No se encontraron recursos no conformes.")
        role_assignments_table = render_table(role_assignments, [
            ("principalType", "Principal Type", "Tipo de Principal", False),
            ("roleDefId", "Role", "Rol", False),
            ("scope_", "Scope", "Alcance", False),
        ], "No role assignments found.", "No se encontraron asignaciones de roles.")
        custom_roles_table = render_table(custom_roles, [
            ("roleName", "Role Name", "Nombre del Rol", False),
            ("description_", "Description", "Descripci\u00f3n", False),
        ], "No custom role definitions found.", "No se encontraron definiciones de roles personalizados.")
        defender_plans_table = render_table(defender_plans, [
            ("name", "Plan", "Plan", False),
            ("tier", "Tier", "Nivel", False),
            ("subPlan", "Sub-plan", "Sub-plan", False),
        ], "No Defender for Cloud plans found.", "No se encontraron planes de Defender for Cloud.")
        secure_scores_table = render_table(secure_scores, [
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("pct", "Score %", "Puntaje %", True),
        ], "No Secure Score data found.", "No se encontraron datos de Secure Score.")
        resource_summary_table = render_table(resource_summary, [
            ("type", "Resource Type", "Tipo de Recurso", False),
            ("location", "Location", "Ubicaci\u00f3n", False),
            ("Count", "Count", "Cantidad", True),
        ], "No resource summary found.", "No se encontr\u00f3 resumen de recursos.")
        orphaned_table = render_table(orphaned, [
            ("Type", "Type", "Tipo", False),
            ("Name", "Name", "Nombre", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("Detail", "Detail", "Detalle", False),
        ], "No orphaned resources found.", "No se encontraron recursos hu\u00e9rfanos.")
        vnets_table = render_table(vnets, [
            ("name", "VNet", "VNet", False),
            ("addressSpace", "Address Space", "Espacio de Direcciones", False),
            ("subnets", "Subnets", "Subredes", True),
            ("peerings", "Peerings", "Emparejamientos", True),
        ], "No virtual networks found.", "No se encontraron redes virtuales.")
        locks_table = render_table(locks, [
            ("name", "Name", "Nombre", False),
            ("lockLevel", "Level", "Nivel", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "No resource locks found.", "No se encontraron bloqueos de recursos.")
        policy_exemptions_table = render_table(policy_exemptions, [
            ("displayName", "Name", "Nombre", False),
            ("exemptionCategory", "Category", "Categor\u00eda", False),
            ("policyAssignmentId", "Policy Assignment", "Asignaci\u00f3n de Pol\u00edtica", False),
            ("expiresOn", "Expires On", "Expira el", False),
        ], "No policy exemptions found.", "No se encontraron exenciones de pol\u00edtica.")

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow">GOVERNANCE VISUALIZER</div>
                <h2>{bi('Management hierarchy, policy, and RBAC', 'Jerarqu\u00eda de administraci\u00f3n, pol\u00edticas y RBAC', 'Hierarquia de gerenciamento, pol\u00edtica e RBAC')}</h2>
                <p>{bi('Management group tree, policy/RBAC assignments, and per-scope insights in one native view.', 'Árbol de grupos de administraci\u00f3n, asignaciones de pol\u00edticas/RBAC y detalles por alcance en una vista nativa.', 'Árvore de grupos de gerenciamento, atribui\u00e7\u00f5es de pol\u00edtica/RBAC e detalhes por escopo em uma \u00fanica visualiza\u00e7\u00e3o nativa.')}</p>
            </div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title security-title">{bi('Hierarchy map', 'Mapa de jerarqu\u00eda', 'Mapa de hierarquia')}{it('Interactive org-chart of Management Groups and Subscriptions (expand a subscription to see its Resource Groups/resources), built from the Governance export.', 'Organigrama interactivo de Grupos de Administraci\u00f3n y Suscripciones (expande una suscripci\u00f3n para ver sus Grupos de Recursos/recursos), construido desde la exportaci\u00f3n de Governance.', 'Organograma interativo de Grupos de Gerenciamento e Assinaturas (expanda uma assinatura para ver seus Grupos de Recursos/recursos), constru\u00eddo a partir da exporta\u00e7\u00e3o do Governance.')}</div>
            {hierarchy_note_html}
            {hierarchy_html}
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Resource explorer', 'Explorador de recursos', 'Explorador de recursos')}{it('Drill-down list of Resource Groups and individual resources per subscription, sourced from the Discovery phase inventory.', 'Lista detallada de Grupos de Recursos y recursos individuales por suscripción, obtenida del inventario de la fase de Discovery.', 'Lista detalhada de Grupos de Recursos e recursos individuais por assinatura, obtida do inventário da fase de Discovery.')}</div>
            <div class="rg-explorer-wrap">{resource_explorer_html}</div>
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Policy assignments', 'Asignaciones de pol\u00edticas', 'Atribui\u00e7\u00f5es de pol\u00edtica')}</h3>{policy_assignments_table}</div>
            <div class="security-panel"><h3>{bi('Custom policy definitions', 'Definiciones de pol\u00edticas personalizadas', 'Defini\u00e7\u00f5es de pol\u00edtica personalizadas')}</h3>{custom_policies_table}</div>
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Non-compliant resources', 'Recursos no conformes', 'Recursos n\u00e3o conformes')}</h3>{policy_compliance_table}</div>
            <div class="security-panel"><h3>{bi('Resource locks', 'Bloqueos de recursos', 'Bloqueios de recursos')}</h3>{locks_table}</div>
        </div>
        <div class="security-panel">
            <h3>{bi('Policy exemptions', 'Exenciones de pol\u00edtica', 'Isen\u00e7\u00f5es de pol\u00edtica')}</h3>
            {policy_exemptions_table}
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Role assignments', 'Asignaciones de roles', 'Atribui\u00e7\u00f5es de fun\u00e7\u00e3o')}</h3>{role_assignments_table}</div>
            <div class="security-panel"><h3>{bi('Custom role definitions', 'Definiciones de roles personalizados', 'Defini\u00e7\u00f5es de fun\u00e7\u00e3o personalizadas')}</h3>{custom_roles_table}</div>
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Defender for Cloud plans', 'Planes de Defender for Cloud', 'Planos do Defender for Cloud')}</h3>{defender_plans_table}</div>
            <div class="security-panel"><h3>{bi('Secure Score by subscription', 'Secure Score por suscripci\u00f3n', 'Secure Score por assinatura')}{it('Microsoft Defender Secure Score percentage per subscription, from the Defender for Cloud scoring API.', 'Porcentaje del Secure Score de Microsoft Defender por suscripci\u00f3n, de la API de puntuaci\u00f3n de Defender for Cloud.', 'Percentual do Secure Score do Microsoft Defender por assinatura, da API de pontua\u00e7\u00e3o do Defender for Cloud.')}</h3>{secure_scores_table}</div>
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Resources by type & location', 'Recursos por tipo y ubicaci\u00f3n', 'Recursos por tipo e localiza\u00e7\u00e3o')}</h3>{resource_summary_table}</div>
            <div class="security-panel"><h3>{bi('Orphaned resources (cost savings)', 'Recursos hu\u00e9rfanos (ahorro de costos)', 'Recursos \u00f3rf\u00e3os (economia de custos)')}</h3>{orphaned_table}</div>
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Virtual networks', 'Redes virtuales', 'Redes virtuais')}{it('Network topology inventory (address spaces, subnets, DDoS protection) collected during the Discovery phase.', 'Inventario de topología de red (espacios de direcciones, subredes, protección DDoS) recolectado durante la fase de Discovery.', 'Inventário de topologia de rede (espaços de endereços, sub-redes, proteção DDoS) coletado durante a fase de Discovery.')}</div>
            {vnets_table}
        </div>
        <div class="section">
            <div class="section-title security-title">{bi('Scope insights', 'Detalles por alcance', 'Detalhes por escopo')}{it('Policy assignment and RBAC role counts rolled up per Management Group scope - click a row to expand its detail.', 'Recuento de asignaciones de políticas y roles RBAC por alcance de Grupo de Administración - haz clic en una fila para ver su detalle.', 'Contagem de atribuições de política e funções RBAC por escopo de Grupo de Gerenciamento - clique em uma linha para ver seu detalhe.')}</div>
            <div class="scope-insights">{scope_insights_html}</div>
        </div>"""

    def _render_cost_view(self) -> str:
        """Render cost optimization insights: right-sizing, idle/orphaned assets, Advisor and reservation savings."""
        advisor = self.data.get("advisor", {})
        metrics = self.data.get("metrics", {})
        governance = self.data.get("governance", {})
        discovery = self.data.get("discovery", {})

        def real_rows(source, sheet_name):
            return [
                row for row in source.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def as_float(value):
            try:
                return float(value)
            except (TypeError, ValueError):
                return 0.0

        def render_table(rows, columns, empty_en, empty_es, limit=15):
            if not rows:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            MAX_TABLE_ROWS = 300
            row_total = len(rows)
            rows = rows[:MAX_TABLE_ROWS]
            header = "".join(f"<th>{bi(label_en, label_es, LABEL_PT.get(label_en, label_en))}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for i, row in enumerate(rows):
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{_cell_html(row.get(key, ""))}</td>'
                    for key, _, _, numeric in columns
                )
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                row_attr += _resource_link_attr(row)
                body += f"<tr{row_attr}>{cells}</tr>"
            more = ""
            if len(rows) > limit:
                more_text = bi(f"+ {len(rows) - limit} more", f"+ {len(rows) - limit} m\u00e1s", f"+ {len(rows) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            if row_total > MAX_TABLE_ROWS:
                more += f'<p class="table-truncated-note">{bi(f"Showing first {MAX_TABLE_ROWS} of {row_total} rows for performance - see the source export for the full list.", f"Mostrando las primeras {MAX_TABLE_ROWS} de {row_total} filas por rendimiento - consulte el archivo de origen para ver la lista completa.", f"Mostrando as primeiras {MAX_TABLE_ROWS} de {row_total} linhas por desempenho - consulte o arquivo de origem para a lista completa.")}</p>'
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>{more}'

        def waste_priority(assessment):
            label = str(assessment or "").lower()
            if "idle" in label or "zero" in label:
                return 0
            if "underutil" in label or "oversiz" in label or "minimal" in label:
                return 1
            if "no data" in label:
                return 3
            return 2

        def assessment_color(assessment):
            label = str(assessment or "").lower()
            if "idle" in label or "zero" in label:
                return "var(--cp-danger)"
            if "underutil" in label or "oversiz" in label or "minimal" in label:
                return "var(--cp-warning)"
            if "no data" in label:
                return "var(--cp-text-muted)"
            return "var(--cp-success)"

        def render_distribution(rows):
            if not rows:
                return ""
            counts = {}
            for row in rows:
                label = str(row.get("Assessment", "Unknown")) or "Unknown"
                counts[label] = counts.get(label, 0) + 1
            total = len(rows)
            bars = "".join(
                f"""<div class="cost-dist-row">
                    <span class="cost-dist-label">{escape(label)}</span>
                    <div class="cost-dist-track"><div class="cost-dist-fill" style="width:{round(count / total * 100)}%; background:{assessment_color(label)};"></div></div>
                    <span class="cost-dist-count">{count}</span>
                </div>"""
                for label, count in sorted(counts.items(), key=lambda kv: kv[1], reverse=True)
            )
            return f'<div class="cost-dist">{bars}</div>'

        def month_label(value):
            if isinstance(value, datetime):
                return value.strftime("%b %Y")
            return str(value) if value not in (None, "") else ""

        def month_sort_key(value):
            return value if isinstance(value, datetime) else datetime.min

        def render_cost_trend_chart(rows):
            """FinOps Hub-style monthly cost trend: Chart.js line chart (real tooltips) + MoM delta badge + simple next-month forecast."""
            monthly = defaultdict(float)
            for r in rows:
                monthly[r.get("BillingMonth")] += as_float(r.get("Cost"))
            months = sorted(monthly.keys(), key=month_sort_key)
            if not months:
                return ""
            values = [monthly[m] for m in months]
            labels = [month_label(m) for m in months]
            n = len(values)
            forecast_value = None
            if n >= 2:
                window = values[max(0, n - 3):]
                deltas = [window[i] - window[i - 1] for i in range(1, len(window))]
                avg_delta = sum(deltas) / len(deltas) if deltas else 0
                forecast_value = max(values[-1] + avg_delta, 0)
            forecast_badge = ""
            chart_labels = {"en": list(labels), "es": list(labels), "pt": list(labels)}
            actual_series = list(values)
            forecast_series = [None] * n
            if forecast_value is not None:
                forecast_word = {"en": "Forecast", "es": "Pron\u00f3stico", "pt": "Previs\u00e3o"}
                for lang, word in forecast_word.items():
                    chart_labels[lang].append(word)
                actual_series.append(None)
                forecast_series.append(forecast_value)
                forecast_series[n - 1] = values[-1]  # connect the dashed segment to the last real point
                forecast_badge = (
                    f'<span class="finops-trend-delta" style="color:var(--cp-warning);">'
                    f'{bi("Next month forecast", "Pron\u00f3stico pr\u00f3ximo mes", "Previs\u00e3o do pr\u00f3ximo m\u00eas")}: ${forecast_value:,.0f}</span>'
                )
            delta_html = ""
            if n >= 2 and values[-2]:
                delta_pct = (values[-1] - values[-2]) / values[-2] * 100
                arrow, color = ("\u25b2", "var(--cp-danger)") if delta_pct > 0 else ("\u25bc", "var(--cp-success)")
                if abs(delta_pct) < 0.05:
                    arrow, color = "\u25cf", "var(--cp-text-muted)"
                delta_html = (
                    f'<span class="finops-trend-delta" style="color:{color};">{arrow} {abs(delta_pct):.1f}% '
                    f'{bi("vs prior month", "vs. mes anterior", "vs. m\u00eas anterior")}</span>'
                )
            self._chart_data["costTrend"] = {
                "type": "line", "canvasId": "costTrendChart",
                "labels": chart_labels,
                "datasets": [
                    {"label": {"en": "Actual", "es": "Real", "pt": "Real"}, "data": actual_series, "colorVar": "--cp-accent", "fill": True},
                    {"label": {"en": "Forecast", "es": "Pron\u00f3stico", "pt": "Previs\u00e3o"}, "data": forecast_series, "colorVar": "--cp-warning", "dashed": True},
                ] if forecast_value is not None else [
                    {"label": {"en": "Actual", "es": "Real", "pt": "Real"}, "data": actual_series, "colorVar": "--cp-accent", "fill": True},
                ],
                "opts": {"currency": True},
            }
            return f"""<div class="finops-trend-card">
                <div class="finops-trend-head">
                    <div>
                        <span class="finops-trend-total">${values[-1]:,.0f}</span>
                        <span class="finops-trend-total-label">{bi("latest month", "\u00faltimo mes", "\u00faltimo m\u00eas")}</span>
                    </div>
                    {delta_html}
                    {forecast_badge}
                </div>
                <div class="chart-card-body" style="height:260px;"><canvas id="costTrendChart"></canvas></div>
            </div>"""

        def render_cost_breakdown_bars(rows, group_key, empty_en, empty_es, limit=6):
            """Horizontal $ bar breakdown by a dimension (service, subscription), FinOps Hub-style."""
            totals = defaultdict(float)
            for r in rows:
                totals[str(r.get(group_key) or "Unknown")] += as_float(r.get("Cost"))
            if not totals:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            items = sorted(totals.items(), key=lambda kv: kv[1], reverse=True)
            total_all = sum(v for _, v in items) or 1
            top = items[:limit]
            other_total = sum(v for _, v in items[limit:])
            if other_total > 0:
                top.append((bi("Other", "Otros", "Outros"), other_total))
            max_val = max(v for _, v in top) if top else 1
            rows_html = "".join(
                f"""<div class="cost-bar-row">
                    <span class="cost-bar-label">{_cell_html(name)}</span>
                    <div class="cost-bar-track"><div class="cost-bar-fill" style="width:{(v / max_val * 100) if max_val else 0:.1f}%;"></div></div>
                    <span class="cost-bar-value">${v:,.0f}</span>
                    <span class="cost-bar-pct">{(v / total_all * 100):.0f}%</span>
                </div>"""
                for name, v in top
            )
            return f'<div class="cost-bar-chart">{rows_html}</div>'

        def render_cost_breakdown_chart(rows, group_key, canvas_id, col_label, empty_en, empty_es, limit=10):
            """Chart.js horizontal bar (Top N by cost) + full 'Show all' table drill-down."""
            totals = defaultdict(float)
            for r in rows:
                totals[str(r.get(group_key) or "Unknown")] += as_float(r.get("Cost"))
            if not totals:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            items = sorted(totals.items(), key=lambda kv: kv[1], reverse=True)
            total_all = sum(v for _, v in items) or 1
            top = items[:limit]
            self._chart_data[canvas_id] = {
                "type": "hbar", "canvasId": canvas_id,
                "labels": {"en": [n for n, _ in top], "es": [n for n, _ in top], "pt": [n for n, _ in top]},
                "values": [round(v, 2) for _, v in top],
                "colorVars": ["--cp-accent"] * len(top),
                "opts": {"currency": True},
            }
            chart_html = f'<div class="chart-card-body" style="height:{max(180, len(top) * 26)}px;"><canvas id="{canvas_id}"></canvas></div>'
            table_rows = "".join(
                f"<tr><td>{escape(name)}</td><td class='num'>${v:,.0f}</td><td class='num'>{(v / total_all * 100):.1f}%</td></tr>"
                for name, v in items
            )
            table_html = f"""<div class="security-table-wrap"><table>
                <thead><tr><th>{col_label}</th><th>{bi('Cost', 'Costo', 'Custo')}</th><th>{bi('Share', 'Participaci\u00f3n', 'Participa\u00e7\u00e3o')}</th></tr></thead>
                <tbody>{table_rows}</tbody>
            </table></div>"""
            return chart_table_block(chart_html, table_html)

        cost_recs = real_rows(advisor, "Cost")
        reservation_recs = [r for r in real_rows(advisor, "ReservationRecommendations") if r.get("ResourceType") or r.get("SkuName")]
        finops = self.data.get("finops", {})
        actual_cost_rows = [r for r in real_rows(finops, "ActualCost") if r.get("Cost") not in (None, "")]
        extended_sheets = [
            ("UnattachedPublicIPs", bi("Unattached Public IPs", "IPs P\u00fablicas No Adjuntas", "IPs P\u00fablicos N\u00e3o Anexados")),
            ("StoppedVMs", bi("Stopped (not deallocated) VMs", "VMs Detenidas (no desasignadas)", "VMs Paradas (n\u00e3o desalocadas)")),
            ("BackendlessAppGateways", bi("Backendless App Gateways", "App Gateways sin Backend", "App Gateways sem Backend")),
            ("BackendlessLoadBalancers", bi("Backendless Load Balancers", "Load Balancers sin Backend", "Load Balancers sem Backend")),
            ("EmptySqlElasticPools", bi("Empty SQL Elastic Pools", "Grupos El\u00e1sticos SQL Vac\u00edos", "Pools El\u00e1sticos SQL Vazios")),
            ("NonSpotAKSPools", bi("Non-Spot AKS Pools", "Pools AKS sin Spot", "Pools AKS sem Spot")),
            ("VMsWithoutHybridBenefit", bi("VMs w/o Hybrid Benefit", "VMs sin Hybrid Benefit", "VMs sem Hybrid Benefit")),
            ("SqlVMsWithoutHybridBenefit", bi("SQL VMs w/o Hybrid Benefit", "VMs SQL sin Hybrid Benefit", "VMs SQL sem Hybrid Benefit")),
        ]
        extended_recs = []
        for sheet_name, label in extended_sheets:
            for row in real_rows(finops, sheet_name):
                row = dict(row)
                row["CheckType"] = label
                extended_recs.append(row)
        vm_rightsizing = sorted(real_rows(metrics, "VM_RightSizing"), key=lambda r: waste_priority(r.get("Assessment")))
        sql_rightsizing = sorted(real_rows(metrics, "SQL_RightSizing"), key=lambda r: waste_priority(r.get("Assessment")))
        plan_rightsizing = sorted(real_rows(metrics, "AppPlan_RightSizing"), key=lambda r: waste_priority(r.get("Assessment")))
        storage_activity = sorted(real_rows(metrics, "Storage_Activity"), key=lambda r: waste_priority(r.get("Assessment")))
        orphaned = real_rows(governance, "OrphanedResources")
        dealloc_vms = real_rows(discovery, "DeallocatedVMs")
        unattached_disks = real_rows(discovery, "UnattachedDisks")
        budgets = real_rows(finops, "Budgets")
        for b in budgets:
            amount = as_float(b.get("Amount"))
            spend = as_float(b.get("CurrentSpend"))
            b["_pct"] = (spend / amount * 100) if amount else 0
        budgets_over = [b for b in budgets if b.get("_pct", 0) >= 100]

        vm_waste = [r for r in vm_rightsizing if waste_priority(r.get("Assessment")) <= 1]
        sql_waste = [r for r in sql_rightsizing if waste_priority(r.get("Assessment")) <= 1]
        plan_waste = [r for r in plan_rightsizing if waste_priority(r.get("Assessment")) <= 1]
        storage_waste = [r for r in storage_activity if waste_priority(r.get("Assessment")) <= 1]

        advisor_savings = sum(as_float(r.get("annualSavings")) for r in cost_recs)
        reservation_savings = sum(as_float(r.get("NetSavings")) for r in reservation_recs)
        total_savings = advisor_savings + reservation_savings
        currency = next((r.get("Currency") for r in reservation_recs if r.get("Currency")), None) \
            or next((r.get("savingsCurrency") for r in cost_recs if r.get("savingsCurrency")), "USD")
        savings_text = f"${total_savings:,.0f}" if total_savings > 0 else "N/A"

        orphaned_total = len(orphaned) + len(dealloc_vms) + len(unattached_disks)
        total_actual_cost = sum(as_float(r.get("Cost")) for r in actual_cost_rows)
        actual_cost_currency = next((r.get("Currency") for r in actual_cost_rows if r.get("Currency")), "USD")

        stats = [
            (savings_text, bi("Est. Annual Savings", "Ahorro Anual Est.", "Economia Anual Est."), "security-accent"),
            (len(cost_recs), bi("Advisor Cost Recs", "Recom. de Costos", "Recom. de Custos"), ""),
            (len(reservation_recs), bi("Reservation Opportunities", "Oportunidades de Reserva", "Oportunidades de Reserva"), ""),
            (len(vm_waste), bi("Idle/Underutilized VMs", "VMs Inactivas/Subutilizadas", "VMs Ociosas/Subutilizadas"), "security-warning"),
            (len(sql_waste), bi("SQL DBs to Rightsize", "BD SQL para Ajustar", "BDs SQL para Ajustar"), "security-warning"),
            (len(plan_waste), bi("App Plans to Rightsize", "Planes App para Ajustar", "Planos App para Ajustar"), "security-warning"),
            (len(storage_waste), bi("Low-activity Storage", "Almacenamiento de Baja Actividad", "Armazenamento de Baixa Atividade"), "security-warning"),
            (orphaned_total, bi("Orphaned/Unused Resources", "Recursos Hu\u00e9rfanos/No Usados", "Recursos \u00d3rf\u00e3os/N\u00e3o Usados"), "security-danger"),
            (len(extended_recs), bi("Extended FinOps Checks", "Verificaciones FinOps Extendidas", "Verifica\u00e7\u00f5es FinOps Estendidas"), "security-warning"),
        ]
        if budgets:
            stats.append((len(budgets_over), bi("Budgets Over Threshold", "Presupuestos Excedidos", "Or\u00e7amentos Excedidos"), "security-danger" if budgets_over else ""))
        if actual_cost_rows:
            stats.insert(1, (f"{total_actual_cost:,.0f} {actual_cost_currency}", bi("Actual Cost (6mo)", "Costo Real (6 meses)", "Custo Real (6 meses)"), ""))
        stat_html = "".join(
            f'<div class="security-stat"><div class="value {color_class}">{value}</div><div class="label">{label}</div></div>'
            for value, label, color_class in stats
        )

        for r in cost_recs:
            savings = as_float(r.get("annualSavings"))
            r["AnnualSavingsDisplay"] = f"{savings:,.0f} {r.get('savingsCurrency', '') or ''}".strip() if savings else ""
        for r in reservation_recs:
            savings = as_float(r.get("NetSavings"))
            r["NetSavingsDisplay"] = f"{savings:,.0f} {r.get('Currency', '') or ''}".strip() if savings else ""
        for r in plan_rightsizing:
            currency = r.get("Currency") or ""
            r["CurrentMonthlyCostDisplay"] = f"{as_float(r.get('CurrentMonthlyCost')):,.0f} {currency}".strip() if r.get("CurrentMonthlyCost") is not None else ""
            r["EstMonthlySavingsDisplay"] = f"{as_float(r.get('EstMonthlySavings')):,.0f} {currency}".strip() if r.get("EstMonthlySavings") is not None else ""
            r["EstAnnualSavingsDisplay"] = f"{as_float(r.get('EstAnnualSavings')):,.0f} {currency}".strip() if r.get("EstAnnualSavings") is not None else ""

        cost_recs_table = render_table(cost_recs, [
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("impact", "Impact", "Impacto", False),
            ("problem", "Problem", "Problema", False),
            ("solution", "Solution", "Soluci\u00f3n", False),
            ("currentSku", "Current SKU", "SKU Actual", False),
            ("targetSku", "Target SKU", "SKU Objetivo", False),
            ("AnnualSavingsDisplay", "Annual Savings", "Ahorro Anual", True),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "No Advisor cost recommendations found.", "No se encontraron recomendaciones de costos de Advisor.")
        reservation_table = render_table(reservation_recs, [
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceType", "Resource Type", "Tipo de Recurso", False),
            ("SkuName", "SKU", "SKU", False),
            ("Term", "Term", "Plazo", False),
            ("RecommendedQty", "Qty", "Cant.", True),
            ("NetSavingsDisplay", "Net Savings/yr", "Ahorro Neto/a\u00f1o", True),
        ], "No reservation/savings plan opportunities found.", "No se encontraron oportunidades de reservas/planes de ahorro.")
        vm_table = render_table(vm_rightsizing, [
            ("Name", "VM", "VM", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("VMSize", "Size", "Tama\u00f1o", False),
            ("AvgCPU_Pct", "Avg CPU %", "CPU Prom. %", True),
            ("MaxCPU_Pct", "Max CPU %", "CPU M\u00e1x. %", True),
            ("Assessment", "Assessment", "Evaluaci\u00f3n", False),
        ], "No VM right-sizing data found.", "No se encontraron datos de ajuste de tama\u00f1o de VMs.")
        sql_table = render_table(sql_rightsizing, [
            ("Name", "Database", "Base de Datos", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("SKU", "SKU", "SKU", False),
            ("MetricType", "Metric", "M\u00e9trica", False),
            ("AvgUsage_Pct", "Avg Usage %", "Uso Prom. %", True),
            ("Assessment", "Assessment", "Evaluaci\u00f3n", False),
        ], "No SQL right-sizing data found.", "No se encontraron datos de ajuste de tama\u00f1o de SQL.")
        plan_table = render_table(plan_rightsizing, [
            ("Name", "App Plan", "Plan de App", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("SKU", "SKU", "SKU", False),
            ("Workers", "Workers", "Workers", True),
            ("AvgCPU_Pct", "Avg CPU %", "CPU Prom. %", True),
            ("Assessment", "Assessment", "Evaluaci\u00f3n", False),
            ("TargetSKU", "Suggested SKU", "SKU Sugerido", False),
            ("CurrentMonthlyCostDisplay", "Est. Monthly Cost", "Costo Mensual Est.", True),
            ("EstMonthlySavingsDisplay", "Est. Monthly Savings", "Ahorro Mensual Est.", True),
            ("EstAnnualSavingsDisplay", "Est. Annual Savings", "Ahorro Anual Est.", True),
        ], "No App Service Plan right-sizing data found.", "No se encontraron datos de ajuste de Planes de App Service.")
        storage_table = render_table(storage_activity, [
            ("Name", "Storage Account", "Cuenta de Almacenamiento", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("SKU", "SKU", "SKU", False),
            ("AvgDailyTxns", "Avg Daily Txns", "Trans. Diarias Prom.", True),
            ("Assessment", "Assessment", "Evaluaci\u00f3n", False),
        ], "No storage activity data found.", "No se encontraron datos de actividad de almacenamiento.")
        orphaned_table = render_table(orphaned, [
            ("Type", "Type", "Tipo", False),
            ("Name", "Name", "Nombre", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("Detail", "Detail", "Detalle", False),
        ], "No orphaned resources found.", "No se encontraron recursos hu\u00e9rfanos.")
        dealloc_table = render_table(dealloc_vms, [
            ("name", "VM", "VM", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("vmSize", "Size", "Tama\u00f1o", False),
            ("location", "Location", "Ubicaci\u00f3n", False),
        ], "No deallocated VMs found.", "No se encontraron VMs desasignadas.")
        unattached_disks_table = render_table(unattached_disks, [
            ("name", "Disk", "Disco", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("diskSizeGB", "Size (GB)", "Tama\u00f1o (GB)", True),
            ("skuName", "SKU", "SKU", False),
        ], "No unattached disks found.", "No se encontraron discos no adjuntos.")
        actual_cost_sorted = sorted(actual_cost_rows, key=lambda r: as_float(r.get("Cost")), reverse=True)
        for r in actual_cost_sorted:
            r["CostDisplay"] = f"{as_float(r.get('Cost')):,.2f} {r.get('Currency', '') or ''}".strip()
            r["MonthDisplay"] = month_label(r.get("BillingMonth"))

        def render_cost_pivot(rows):
            """Service x Month pivot (with a Total column/row) - far easier to scan than a flat
            Subscription/Service/Month/Cost list once there are more than a handful of rows."""
            services = sorted({r.get("ServiceName") for r in rows if r.get("ServiceName")})
            months = sorted({r.get("BillingMonth") for r in rows if r.get("BillingMonth") is not None}, key=month_sort_key)
            if not services or not months:
                return f'<p class="empty-state">{bi("No actual cost data found (requires Cost Management Reader).", "No se encontraron datos de costo real (requiere el rol Cost Management Reader).", "Nenhum dado de custo real encontrado (requer a fun\u00e7\u00e3o Cost Management Reader).")}</p>'
            cost_map = defaultdict(float)
            totals = defaultdict(float)
            for r in rows:
                svc = r.get("ServiceName")
                mon = r.get("BillingMonth")
                if not svc or mon is None:
                    continue
                c = as_float(r.get("Cost"))
                cost_map[(svc, mon)] += c
                totals[svc] += c
            services_sorted = sorted(services, key=lambda s: totals.get(s, 0), reverse=True)
            currency = next((r.get("Currency") for r in rows if r.get("Currency")), "") or ""
            month_totals = {m: sum(cost_map.get((s, m), 0) for s in services) for m in months}
            grand_total = sum(totals.values())
            header = f"<th>{bi('Service', 'Servicio', 'Servi\u00e7o')}</th>" + "".join(
                f'<th class="num">{escape(month_label(m))}</th>' for m in months
            ) + f'<th class="num">{bi("Total", "Total", "Total")}</th>'
            limit = 15
            body = ""
            for i, svc in enumerate(services_sorted):
                cells = "".join(f'<td class="num">{cost_map.get((svc, m), 0):,.0f}</td>' for m in months)
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                body += f'<tr{row_attr}><td>{escape(svc)}</td>{cells}<td class="num"><strong>{totals[svc]:,.0f}</strong></td></tr>'
            total_cells = "".join(f'<td class="num">{month_totals.get(m, 0):,.0f}</td>' for m in months)
            footer = f'<tr class="cost-pivot-total-row"><td>{bi("Total", "Total", "Total")}</td>{total_cells}<td class="num">{grand_total:,.0f} {escape(currency)}</td></tr>'
            more = ""
            if len(services_sorted) > limit:
                more_text = bi(f"+ {len(services_sorted) - limit} more", f"+ {len(services_sorted) - limit} m\u00e1s", f"+ {len(services_sorted) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            return f'<div class="cost-pivot-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}{footer}</tbody></table></div>{more}'

        def render_actual_cost_section(rows):
            """Service x Month pivot with a multi-select Subscription checkbox filter. The pivot
            is rebuilt client-side (see initCostPivotBlocks() in the page script) from the raw
            per-subscription rows embedded as JSON, so any combination of subscriptions can be
            selected at once - not just one at a time."""
            if not rows:
                return f'<p class="empty-state">{bi("No actual cost data found (requires Cost Management Reader).", "No se encontraron datos de costo real (requiere el rol Cost Management Reader).", "Nenhum dado de custo real encontrado (requer a fun\u00e7\u00e3o Cost Management Reader).")}</p>'
            subs = sorted({r.get("Subscription") for r in rows if r.get("Subscription")})
            if len(subs) <= 1:
                # Only one subscription in scope - a checkbox filter would be pointless, but
                # still label which subscription the numbers below belong to.
                pivot_html = render_cost_pivot(rows)
                if not subs:
                    return pivot_html
                label = (
                    '<div class="cost-pivot-toolbar cost-pivot-static-label">'
                    f'<span class="cost-pivot-toolbar-label">{bi("Subscription", "Suscripci\u00f3n", "Assinatura")}:</span>'
                    f'<span class="cost-pivot-sub-badge">{escape(subs[0])}</span>'
                    '</div>'
                )
                return f'{label}{pivot_html}'

            agg = defaultdict(float)
            for r in rows:
                svc, mon, sub = r.get("ServiceName"), r.get("BillingMonth"), r.get("Subscription") or ""
                if not svc or mon is None:
                    continue
                agg[(sub, svc, month_label(mon))] += as_float(r.get("Cost"))
            months_sorted = sorted({r.get("BillingMonth") for r in rows if r.get("BillingMonth") is not None}, key=month_sort_key)
            currency = next((r.get("Currency") for r in rows if r.get("Currency")), "") or ""
            payload = json.dumps({
                "months": [month_label(m) for m in months_sorted],
                "currency": currency,
                "rows": [{"sub": k[0], "svc": k[1], "month": k[2], "cost": round(v, 2)} for k, v in agg.items()],
            }, ensure_ascii=False)

            checks = "".join(
                f'<label class="cost-pivot-sub-check"><input type="checkbox" value="{escape(s)}" checked> {escape(s)}</label>'
                for s in subs
            )
            toolbar = (
                '<div class="cost-pivot-toolbar">'
                f'<span class="cost-pivot-toolbar-label">{bi("Subscriptions", "Suscripciones", "Assinaturas")}:</span>'
                f'{checks}'
                f'<button type="button" class="cost-pivot-selectall" data-mode="all">{bi("All", "Todas", "Todas")}</button>'
                f'<button type="button" class="cost-pivot-selectall" data-mode="none">{bi("None", "Ninguna", "Nenhuma")}</button>'
                '</div>'
            )
            return f'<div class="cost-pivot-block" data-payload="{escape(payload)}">{toolbar}<div class="cost-pivot-render-target"></div></div>'

        actual_cost_table = render_actual_cost_section(actual_cost_rows)
        cost_trend_chart = render_cost_trend_chart(actual_cost_rows)
        cost_by_service_bars = render_cost_breakdown_chart(
            actual_cost_rows, "ServiceName", "costByServiceChart", bi("Service", "Servicio", "Servi\u00e7o"),
            "No service cost data found.", "No se encontraron datos de costo por servicio.",
        )
        cost_by_subscription_bars = render_cost_breakdown_bars(
            actual_cost_rows, "Subscription",
            "No subscription cost data found.", "No se encontraron datos de costo por suscripci\u00f3n.",
        )
        cost_by_region_bars = render_cost_breakdown_bars(
            actual_cost_rows, "ResourceLocation",
            "No region cost data found.", "No se encontraron datos de costo por regi\u00f3n.",
        )

        def render_budget_bars(rows):
            if not rows:
                return f'<p class="empty-state">{bi("No budgets configured (Microsoft.Consumption/budgets).", "No hay presupuestos configurados (Microsoft.Consumption/budgets).", "Nenhum or\u00e7amento configurado (Microsoft.Consumption/budgets).")}</p>'
            rows_html = ""
            for b in sorted(rows, key=lambda r: r.get("_pct", 0), reverse=True):
                pct = min(b.get("_pct", 0), 100)
                color = "var(--cp-danger)" if b.get("_pct", 0) >= 100 else "var(--cp-warning)" if b.get("_pct", 0) >= 75 else "var(--cp-success)"
                currency = b.get("Currency", "") or ""
                label = f"{b.get('BudgetName', 'Budget')} ({b.get('Subscription', '')})"
                rows_html += f"""<div class="cost-bar-row">
                    <span class="cost-bar-label">{_cell_html(label)}</span>
                    <div class="cost-bar-track"><div class="cost-bar-fill" style="width:{pct:.0f}%; background:{color};"></div></div>
                    <span class="cost-bar-value">{as_float(b.get('CurrentSpend')):,.0f} / {as_float(b.get('Amount')):,.0f} {currency}</span>
                    <span class="cost-bar-pct">{b.get('_pct', 0):.0f}%</span>
                </div>"""
            return f'<div class="cost-bar-chart">{rows_html}</div>'

        budget_bars_html = render_budget_bars(budgets)

        def render_expensive_at_risk(rows):
            """Cross-cut: highest-spend services that also have open action items in other pillars."""
            service_totals = defaultdict(float)
            for r in rows:
                service_totals[str(r.get("ServiceName") or "Unknown")] += as_float(r.get("Cost"))
            if not service_totals:
                return f'<p class="empty-state">{bi("No cost data available to cross-reference.", "No hay datos de costo disponibles para cruzar.", "Nenhum dado de custo disponível para cruzamento.")}</p>'
            top_services = sorted(service_totals.items(), key=lambda kv: kv[1], reverse=True)[:12]
            other_actions = [a for a in self.engine.action_items if a.get("pillar") != "Cost Optimization"]
            sev_weight = {"critical": 4, "high": 3, "medium": 2, "low": 1, "info": 0}
            cross_rows = []
            for name, cost in top_services:
                name_lc = name.lower()
                if not name_lc or name_lc == "unknown":
                    continue
                matches = [
                    a for a in other_actions
                    if name_lc in str(a.get("title", "")).lower()
                    or name_lc in str(a.get("description", "")).lower()
                    or any(name_lc in str(res).lower() for res in (a.get("resources") or []))
                ]
                if not matches:
                    continue
                worst = max(matches, key=lambda a: sev_weight.get(str(a.get("severity", "")).lower(), 0))
                cross_rows.append({
                    "Service": name,
                    "CostDisplay": f"${cost:,.0f}",
                    "OpenFindings": len(matches),
                    "TopSeverity": str(worst.get("severity", "")).upper(),
                    "TopIssue": bi(worst.get("title", ""), worst.get("title_es", worst.get("title", "")), worst.get("title_pt", worst.get("title", ""))),
                })
            if not cross_rows:
                return f'<p class="empty-state">{bi("No expensive services currently have open findings in other pillars.", "Ning\u00fan servicio costoso tiene actualmente hallazgos abiertos en otros pilares.", "Nenhum servi\u00e7o caro possui atualmente descobertas abertas em outros pilares.")}</p>'
            return render_table(cross_rows, [
                ("Service", "Service", "Servicio", False),
                ("CostDisplay", "6-mo Cost", "Costo 6 meses", True),
                ("OpenFindings", "Open Findings", "Hallazgos Abiertos", True),
                ("TopSeverity", "Top Severity", "Severidad M\u00e1xima", False),
                ("TopIssue", "Top Issue", "Problema Principal", False),
            ], "", "")

        expensive_at_risk_table = render_expensive_at_risk(actual_cost_rows)
        extended_recs_table = render_table(extended_recs, [
            ("CheckType", "Check", "Verificaci\u00f3n", False),
            ("name", "Resource", "Recurso", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("location", "Location", "Ubicaci\u00f3n", False),
        ], "No extended FinOps findings.", "No se encontraron hallazgos FinOps extendidos.", limit=20)

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow" style="color: var(--cp-success);">COST OPTIMIZATION</div>
                <h2>{bi('Cost optimization opportunities', 'Oportunidades de optimizaci\u00f3n de costos', 'Oportunidades de otimiza\u00e7\u00e3o de custos')}</h2>
                <p>{bi('Right-sizing, idle/orphaned assets, Advisor cost recommendations, and reservation coverage in one view.', 'Ajuste de tama\u00f1o, activos inactivos/hu\u00e9rfanos, recomendaciones de costos de Advisor y cobertura de reservas en una sola vista.', 'Ajuste de tamanho, ativos ociosos/\u00f3rf\u00e3os, recomenda\u00e7\u00f5es de custo do Advisor e cobertura de reservas em uma \u00fanica vis\u00e3o.')}</p>
            </div>
            <div class="security-score-mark" style="color: var(--cp-success);">{savings_text}</div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-accent);">{bi('Cost trends', 'Tendencias de costos', 'Tendências de custos')}{it('Monthly spend trend from Cost Management, compared against the previous period.', 'Tendencia de gasto mensual de Cost Management, comparada con el período anterior.', 'Tendência de gasto mensal do Cost Management, comparada com o período anterior.')}</div>
            {cost_trend_chart if cost_trend_chart else f'<p class="empty-state">{bi("No actual cost data found (requires Cost Management Reader).", "No se encontraron datos de costo real (requiere el rol Cost Management Reader).", "Nenhum dado de custo real encontrado (requer a fun\u00e7\u00e3o Cost Management Reader).")}</p>'}
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Cost by service', 'Costo por servicio', 'Custo por servi\u00e7o')}</h3>{cost_by_service_bars}</div>
                <div class="security-panel"><h3>{bi('Cost by subscription', 'Costo por suscripci\u00f3n', 'Custo por assinatura')}</h3>{cost_by_subscription_bars}</div>
            </div>
            <div class="security-panel">
                <h3>{bi('Cost by region', 'Costo por regi\u00f3n', 'Custo por regi\u00e3o')}</h3>
                {cost_by_region_bars}
            </div>
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-warning);">{bi('Budgets vs. actual spend', 'Presupuestos vs. gasto real', 'Or\u00e7amentos vs. gasto real')}{it('Configured Azure Budgets compared with actual spend to date, highlighting budgets at risk of being exceeded.', 'Presupuestos de Azure configurados comparados con el gasto real a la fecha, resaltando los que est\u00e1n en riesgo de excederse.', 'Or\u00e7amentos do Azure configurados comparados com o gasto real at\u00e9 o momento, destacando os que est\u00e3o em risco de serem excedidos.')}</div>
            {budget_bars_html}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-danger);">{bi('Expensive services with open risk findings', 'Servicios costosos con hallazgos de riesgo abiertos', 'Servi\u00e7os caros com descobertas de risco abertas')}{it('High-spend services cross-referenced with open Advisor/Defender findings, to prioritize cost and risk together.', 'Servicios de mayor gasto cruzados con hallazgos abiertos de Advisor/Defender, para priorizar costo y riesgo en conjunto.', 'Servi\u00e7os de maior gasto cruzados com descobertas abertas do Advisor/Defender, para priorizar custo e risco juntos.')}</div>
            {expensive_at_risk_table}
        </div>

        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Right-sizing opportunities', 'Oportunidades de ajuste de tama\u00f1o', 'Oportunidades de ajuste de tamanho')}{it('Advisor recommendations to resize or shut down underutilized compute and database resources.', 'Recomendaciones de Advisor para redimensionar o apagar recursos de c\u00f3mputo y bases de datos subutilizados.', 'Recomenda\u00e7\u00f5es do Advisor para redimensionar ou desligar recursos de computa\u00e7\u00e3o e banco de dados subutilizados.')}</div>
            <div class="security-panel" style="margin-bottom: 18px;"><h3>{bi('Virtual Machines', 'M\u00e1quinas Virtuales', 'M\u00e1quinas Virtuais')}</h3>{render_distribution(vm_rightsizing)}{vm_table}</div>
            <div class="security-panel" style="margin-bottom: 18px;"><h3>{bi('SQL Databases', 'Bases de Datos SQL', 'Bancos de Dados SQL')}</h3>{render_distribution(sql_rightsizing)}{sql_table}</div>
            <div class="security-panel" style="margin-bottom: 18px;"><h3>{bi('App Service Plans', 'Planes de App Service', 'Planos do App Service')}</h3>{render_distribution(plan_rightsizing)}{plan_table}</div>
            <div class="security-panel"><h3>{bi('Storage Accounts', 'Cuentas de Almacenamiento', 'Contas de Armazenamento')}</h3>{render_distribution(storage_activity)}{storage_table}</div>
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Idle and orphaned resources', 'Recursos inactivos y hu\u00e9rfanos', 'Recursos ociosos e \u00f3rf\u00e3os')}{it('Resources with no active workload (unattached disks, unused public IPs, stopped-but-billed VMs) that are safe cleanup candidates.', 'Recursos sin carga de trabajo activa (discos sin adjuntar, IPs p\u00fablicas sin uso, VMs detenidas pero facturadas) candidatos seguros para limpieza.', 'Recursos sem carga de trabalho ativa (discos n\u00e3o anexados, IPs p\u00fablicos n\u00e3o utilizados, VMs paradas mas cobradas) candidatos seguros para limpeza.')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Deallocated VMs', 'VMs Desasignadas', 'VMs Desalocadas')}</h3>{dealloc_table}</div>
                <div class="security-panel"><h3>{bi('Unattached Managed Disks', 'Discos Administrados No Adjuntos', 'Discos Gerenciados N\u00e3o Anexados')}</h3>{unattached_disks_table}</div>
            </div>
            <div class="security-panel">
                <h3>{bi('Orphaned resources (all types)', 'Recursos hu\u00e9rfanos (todos los tipos)', 'Recursos \u00f3rf\u00e3os (todos os tipos)')}</h3>
                {orphaned_table}
            </div>
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Actual spend (last 6 months)', 'Gasto real (\u00faltimos 6 meses)', 'Gasto real (\u00faltimos 6 meses)')}{it('Six-month historical spend by service, from Cost Management usage details.', 'Gasto hist\u00f3rico de seis meses por servicio, de los detalles de uso de Cost Management.', 'Gasto hist\u00f3rico de seis meses por servi\u00e7o, dos detalhes de uso do Cost Management.')}</div>
            {actual_cost_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Extended FinOps checks', 'Verificaciones FinOps extendidas', 'Verifica\u00e7\u00f5es FinOps estendidas')}{it('Additional FinOps hygiene checks (cost allocation tagging, anomaly detection) beyond the core recommendations.', 'Verificaciones adicionales de higiene FinOps (etiquetado para asignaci\u00f3n de costos, detecci\u00f3n de anomal\u00edas) m\u00e1s all\u00e1 de las recomendaciones principales.', 'Verifica\u00e7\u00f5es adicionais de higiene FinOps (tags para aloca\u00e7\u00e3o de custos, detec\u00e7\u00e3o de anomalias) al\u00e9m das recomenda\u00e7\u00f5es principais.')}</div>
            {extended_recs_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Advisor cost recommendations', 'Recomendaciones de costos de Advisor', 'Recomenda\u00e7\u00f5es de custo do Advisor')}{it('Cost optimization recommendations generated by Azure Advisor.', 'Recomendaciones de optimizaci\u00f3n de costos generadas por Azure Advisor.', 'Recomenda\u00e7\u00f5es de otimiza\u00e7\u00e3o de custo geradas pelo Azure Advisor.')}</div>
            {cost_recs_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Reservation and Savings Plan opportunities', 'Oportunidades de Reservas y Planes de Ahorro', 'Oportunidades de Reservas e Planos de Economia')}{it('Potential savings from Reserved Instances or Savings Plans, based on your recent usage patterns.', 'Ahorros potenciales de Instancias Reservadas o Planes de Ahorro, basados en tus patrones de uso recientes.', 'Economias potenciais de Instâncias Reservadas ou Planos de Economia, com base em seus padrões de uso recentes.')}</div>
            {reservation_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-success);">{bi('Action items', 'Elementos de acci\u00f3n', 'Itens de a\u00e7\u00e3o')}{it('Prioritized cost actions, ranked by estimated savings impact.', 'Acciones de costo priorizadas, ordenadas por el ahorro estimado.', 'A\u00e7\u00f5es de custo priorizadas, ordenadas pelo impacto de economia estimado.')}</div>
            {self._render_actions_for_pillar("Cost Optimization")}
        </div>"""

    @staticmethod
    def _action_item_id(action: dict) -> str:
        """Stable id derived from pillar+title so dismiss state survives re-running the generator
        against the same scan data (the action_items list itself isn't guaranteed stable across runs)."""
        basis = f"{action.get('pillar', '')}|{action.get('title', '')}"
        return hashlib.md5(basis.encode("utf-8")).hexdigest()[:12]

    def _render_action_card(self, action: dict, index, show_pillar: bool = True) -> str:
        """Render a single action-item card with a dismiss checkbox (dismiss/custom state is
        persisted client-side via localStorage, see the action-item JS in generate())."""
        color = DashboardConfig.SEVERITY_COLORS.get(action["severity"], "var(--cp-text-muted)")
        resources_html = ""
        if action.get("resources"):
            items = []
            for r in action["resources"]:
                if not r:
                    continue
                if isinstance(r, dict):
                    name = escape(str(r.get("name", "")))
                    link_attr = _resource_link_attr({"id": r.get("id")}) if r.get("id") else ""
                else:
                    name = escape(str(r))
                    link_attr = ""
                items.append(f"<li{link_attr}>{name}</li>")
            resources_html = "<ul class='resource-list'>" + "".join(items) + "</ul>"
        severity_badge = bi(str(action['severity']).upper(), str(SEVERITY_ES.get(action['severity'], action['severity'].upper())), str(SEVERITY_PT.get(action['severity'], action['severity'].upper())))
        pillar_tag_html = ""
        if show_pillar:
            pillar_tag_html = f"<span class='pillar-tag'>{bi(str(action['pillar']), str(PILLAR_ES.get(action['pillar'], action['pillar'])), str(PILLAR_PT.get(action['pillar'], action['pillar'])))}</span>"
        title_text = bi(action['title'], action.get('title_es', action['title']), action.get('title_pt', action['title']))
        desc_text = bi(action['description'], action.get('description_es', action['description']), action.get('description_pt', action['description']))
        item_id = self._action_item_id(action)
        subscriptions_attr = f' data-subscriptions="{escape(json.dumps(action["subscriptions"], ensure_ascii=False))}"' if action.get("subscriptions") else ""
        return f"""
            <div class="action-item" data-item-id="{item_id}" data-pillar="{escape(action['pillar'])}" data-severity="{action['severity']}"{subscriptions_attr} style="border-left: 4px solid {color};">
                <div class="action-header">
                    <label class="action-check-wrap" title="Dismiss / Descartar">
                        <input type="checkbox" class="action-check" data-item-id="{item_id}">
                    </label>
                    <span class="action-num">#{index}</span>
                    <span class="severity-badge" style="background:{color};">{severity_badge}</span>
                    {pillar_tag_html}
                    <span class="action-title">{title_text}</span>
                </div>
                <p class="action-desc">{desc_text}</p>
                {resources_html}
            </div>"""

    def _render_actions_for_pillar(self, pillar: str) -> str:
        """Render the action-item cards for a single pillar (used by the dedicated pillar tabs)."""
        items = [a for a in self.engine.action_items if a["pillar"] == pillar]
        if not items:
            return f'<div class="action-list" data-pillar="{escape(pillar)}"><p class="empty-state">{bi("No action items for this pillar.", "Sin elementos de acci\u00f3n para este pilar.", "Sem itens de a\u00e7\u00e3o para este pilar.")}</p></div>'
        cards = "".join(self._render_action_card(action, i, show_pillar=False) for i, action in enumerate(items, 1))
        return f'<div class="action-list" data-pillar="{escape(pillar)}">{cards}</div>'

    def _render_subcategory_table(self, pillar: str) -> str:
        """Render the subcategory weight/health breakdown table for a single pillar (used by the dedicated pillar tabs)."""
        rows = self.engine.breakdowns.get(pillar, {}).get("rows", [])
        if not rows:
            return f'<p class="empty-state">{bi("Score breakdown unavailable.", "Desglose de puntaje no disponible.", "Detalhamento de pontua\u00e7\u00e3o indispon\u00edvel.")}</p>'
        rows_html = "".join(
            f"<tr data-subcat=\"{escape(pillar)}|{escape(r['subcategory'])}\"><td>{bi(str(r['subcategory']), str(SUBCATEGORY_ES.get(r['subcategory'], r['subcategory'])), str(SUBCATEGORY_PT.get(r['subcategory'], r['subcategory'])))}</td><td class='num'>{r['weight']}</td>"
            f"<td class='num sc-ht'>{r['healthy']}/{r['total']}</td><td class='num sc-pct'>{r['pct']}%</td></tr>"
            for r in rows
        )
        return f"""<div class="security-table-wrap"><table>
            <thead><tr><th>{bi('Subcategory', 'Subcategor\u00eda', 'Subcategoria')}</th><th>{bi('Weight', 'Peso', 'Peso')}</th><th>{bi('Healthy/Total', 'Saludable/Total', 'Saud\u00e1vel/Total')}</th><th>{bi('Score', 'Puntaje', 'Pontua\u00e7\u00e3o')}</th></tr></thead>
            <tbody>{rows_html}</tbody>
        </table></div>"""

    def _render_reliability_view(self) -> str:
        """Render reliability posture: Availability Zones, backup coverage, DDoS protection, and Advisor HA recommendations."""
        discovery = self.data.get("discovery", {})
        governance = self.data.get("governance", {})
        advisor = self.data.get("advisor", {})

        def real_rows(source, sheet_name):
            return [
                row for row in source.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def render_table(rows, columns, empty_en, empty_es, limit=15):
            if not rows:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            MAX_TABLE_ROWS = 300
            row_total = len(rows)
            rows = rows[:MAX_TABLE_ROWS]
            header = "".join(f"<th>{bi(label_en, label_es, LABEL_PT.get(label_en, label_en))}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for i, row in enumerate(rows):
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{escape(str(row.get(key, "")))}</td>'
                    for key, _, _, numeric in columns
                )
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                row_attr += _resource_link_attr(row)
                body += f"<tr{row_attr}>{cells}</tr>"
            more = ""
            if len(rows) > limit:
                more_text = bi(f"+ {len(rows) - limit} more", f"+ {len(rows) - limit} m\u00e1s", f"+ {len(rows) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            if row_total > MAX_TABLE_ROWS:
                more += f'<p class="table-truncated-note">{bi(f"Showing first {MAX_TABLE_ROWS} of {row_total} rows for performance - see the source export for the full list.", f"Mostrando las primeras {MAX_TABLE_ROWS} de {row_total} filas por rendimiento - consulte el archivo de origen para ver la lista completa.", f"Mostrando as primeiras {MAX_TABLE_ROWS} de {row_total} linhas por desempenho - consulte o arquivo de origem para a lista completa.")}</p>'
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>{more}'

        def ddos_enabled(v):
            return str(v.get("ddosProtection", v.get("ddos", ""))).lower() not in ("false", "", "none")

        vms = real_rows(discovery, "VMs")
        vms_no_zone = [v for v in vms if not v.get("zone") and not v.get("availabilityZone")]
        backups = real_rows(discovery, "BackupVaults")
        vnets = real_rows(discovery, "VNets") or real_rows(governance, "VNets")
        vnets_no_ddos = [v for v in vnets if not ddos_enabled(v)]
        protected_items = real_rows(discovery, "BackupProtectedItems")
        failed_backups = [
            p for p in protected_items
            if str(p.get("lastBackupStatus", "")).lower() not in ("completed", "success", "succeeded", "")
        ]
        health_events = sorted(
            real_rows(discovery, "ResourceHealthEvents"),
            key=lambda e: str(e.get("status", "")).lower() != "active"
        )
        active_health_events = [e for e in health_events if str(e.get("status", "")).lower() == "active"]
        reliability_recs = sorted(
            real_rows(advisor, "Reliability"),
            key=lambda r: {"high": 0, "medium": 1, "low": 2}.get(str(r.get("impact", "")).lower(), 3)
        )
        high_impact = [r for r in reliability_recs if str(r.get("impact", "")).lower() == "high"]

        score = self.engine.scores.get("Reliability", 0)
        stats = [
            (f'<span data-score="Reliability" data-score-suffix="/100">{score}/100</span>', bi("Reliability Score", "Puntaje de Confiabilidad", "Pontua\u00e7\u00e3o de Confiabilidade"), "security-accent"),
            (len(vms_no_zone), bi("VMs without Availability Zone", "VMs sin Zona de Disponibilidad", "VMs sem Zona de Disponibilidade"), "security-warning" if vms_no_zone else ""),
            (len(backups), bi("Backup Vaults Found", "Vaults de Backup Encontrados", "Vaults de Backup Encontrados"), "" if backups else "security-danger"),
            (len(vnets_no_ddos), bi("VNets without DDoS Protection", "VNets sin Protecci\u00f3n DDoS", "VNets sem Prote\u00e7\u00e3o DDoS"), "security-warning" if vnets_no_ddos else ""),
            (len(reliability_recs), bi("Advisor Reliability Recs", "Recom. de Confiabilidad", "Recom. de Confiabilidade"), ""),
            (len(high_impact), bi("High-Impact Recommendations", "Recomendaciones de Alto Impacto", "Recomenda\u00e7\u00f5es de Alto Impacto"), "security-danger" if high_impact else ""),
            (len(failed_backups), bi("Backup Items Failing/Not Protected", "Elementos de Backup Fallidos/No Protegidos", "Itens de Backup Falhando/N\u00e3o Protegidos"), "security-danger" if failed_backups else ""),
            (len(active_health_events), bi("Active Resource Health Events", "Eventos Activos de Resource Health", "Eventos Ativos de Resource Health"), "security-danger" if active_health_events else ""),
        ]
        stat_html = "".join(
            f'<div class="security-stat"><div class="value {color_class}">{value}</div><div class="label">{label}</div></div>'
            for value, label, color_class in stats
        )

        vms_table = render_table(vms_no_zone, [
            ("name", "VM", "VM", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("location", "Location", "Ubicaci\u00f3n", False),
            ("vmSize", "Size", "Tama\u00f1o", False),
            ("powerState", "Power State", "Estado", False),
        ], "All VMs are deployed across Availability Zones.", "Todas las VMs est\u00e1n desplegadas en Zonas de Disponibilidad.")
        backups_table = render_table(backups, [
            ("name", "Backup Vault", "Vault de Backup", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("location", "Location", "Ubicaci\u00f3n", False),
        ], "No Backup/Recovery Services vaults found.", "No se encontraron vaults de Backup/Recovery Services.")
        vnets_table = render_table(vnets_no_ddos, [
            ("name", "VNet", "VNet", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("addressSpace", "Address Space", "Espacio de Direcciones", False),
        ], "All VNets have DDoS Protection enabled.", "Todas las VNets tienen Protecci\u00f3n DDoS habilitada.")
        recs_table = render_table(reliability_recs, [
            ("impact", "Impact", "Impacto", False),
            ("impactedResource", "Resource", "Recurso", False),
            ("problem", "Problem", "Problema", False),
            ("solution", "Solution", "Soluci\u00f3n", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "No Advisor reliability recommendations found.", "No se encontraron recomendaciones de confiabilidad de Advisor.")
        impact_counts = defaultdict(int)
        for r in reliability_recs:
            impact_counts[str(r.get("impact", "")).lower() or "unknown"] += 1
        impact_present = [i for i in ("high", "medium", "low") if impact_counts.get(i, 0) > 0]
        if impact_present:
            self._chart_data["reliabilityImpact"] = {
                "type": "donut", "canvasId": "reliabilityImpactChart",
                "labels": {lang: [SEVERITY_LABELS.get(i, {}).get(lang, i.capitalize()) for i in impact_present] for lang in ("en", "es", "pt")},
                "values": [impact_counts[i] for i in impact_present],
                "colorVars": [_bare_var(DashboardConfig.SEVERITY_COLORS.get(i)) for i in impact_present],
            }
            recs_chart_html = '<div class="chart-card-body" style="height:200px;"><canvas id="reliabilityImpactChart"></canvas></div>'
            recs_block = chart_table_block(recs_chart_html, recs_table)
        else:
            recs_block = recs_table
        protected_items_table = render_table(protected_items, [
            ("friendlyName", "Item", "Elemento", False),
            ("protectedItemType", "Type", "Tipo", False),
            ("workloadType", "Workload", "Carga de Trabajo", False),
            ("protectionStatus", "Protection Status", "Estado de Protecci\u00f3n", False),
            ("lastBackupStatus", "Last Backup Status", "\u00daltimo Estado de Backup", False),
            ("lastBackupTime", "Last Backup Time", "\u00daltima Hora de Backup", False),
            ("policyName", "Policy", "Pol\u00edtica", False),
        ], "No backup-protected items found (vaults exist but nothing is protected, or Backup Reader access is missing).", "No se encontraron elementos protegidos por backup (existen vaults pero nada est\u00e1 protegido, o falta acceso de Backup Reader).")
        health_events_table = render_table(health_events, [
            ("title_", "Event", "Evento", False),
            ("eventType", "Type", "Tipo", False),
            ("status", "Status", "Estado", False),
            ("level", "Level", "Nivel", False),
            ("impactStartTime", "Impact Start", "Inicio de Impacto", False),
            ("impactMitigationTime", "Mitigated", "Mitigado", False),
        ], "No Azure Resource Health / Service Health events found.", "No se encontraron eventos de Resource Health / Service Health de Azure.")

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow" style="color: var(--cp-link);">RELIABILITY</div>
                <h2>{bi('High availability and disaster recovery', 'Alta disponibilidad y recuperaci\u00f3n ante desastres', 'Alta disponibilidade e recupera\u00e7\u00e3o de desastres')}</h2>
                <p>{bi('Availability Zones, backup coverage, DDoS protection, and Advisor high-availability recommendations in one view.', 'Zonas de Disponibilidad, cobertura de backup, protecci\u00f3n DDoS y recomendaciones de alta disponibilidad de Advisor en una sola vista.', 'Zonas de Disponibilidade, cobertura de backup, prote\u00e7\u00e3o DDoS e recomenda\u00e7\u00f5es de alta disponibilidade do Advisor em uma \u00fanica vis\u00e3o.')}</p>
            </div>
            <div class="security-score-mark" style="color: var(--cp-link);"><span data-score="Reliability" data-score-suffix="/100">{score}/100</span></div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-link);">{bi('VMs without Availability Zones', 'VMs sin Zonas de Disponibilidad', 'VMs sem Zonas de Disponibilidade')}{it('Virtual machines not deployed across Availability Zones, representing a single point of failure risk.', 'M\u00e1quinas virtuales que no est\u00e1n distribuidas en Zonas de Disponibilidad, lo que representa un riesgo de punto \u00fanico de falla.', 'M\u00e1quinas virtuais n\u00e3o distribu\u00eddas em Zonas de Disponibilidade, representando um risco de ponto \u00fanico de falha.')}</div>
            {vms_table}
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Backup vault coverage', 'Cobertura de vaults de backup', 'Cobertura de vaults de backup')}</h3>{backups_table}</div>
            <div class="security-panel"><h3>{bi('VNets without DDoS Protection', 'VNets sin Protecci\u00f3n DDoS', 'VNets sem Prote\u00e7\u00e3o DDoS')}</h3>{vnets_table}</div>
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-link);">{bi('Real backup coverage (protected items)', 'Cobertura real de backup (elementos protegidos)', 'Cobertura real de backup (itens protegidos)')}{it('Resources actually protected by Azure Backup / Recovery Services vaults, versus resources that should be protected.', 'Recursos realmente protegidos por Azure Backup / b\u00f3vedas de Recovery Services, frente a los recursos que deber\u00edan estarlo.', 'Recursos realmente protegidos por Azure Backup / cofres do Recovery Services, versus os recursos que deveriam estar protegidos.')}</div>
            {protected_items_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-link);">{bi('Resource / Service Health events', 'Eventos de Resource Health / Service Health', 'Eventos de Resource Health / Service Health')}{it('Recent Azure Resource Health and Service Health events affecting your resources.', 'Eventos recientes de Resource Health y Service Health de Azure que afectan tus recursos.', 'Eventos recentes de Resource Health e Service Health do Azure que afetam seus recursos.')}</div>
            {health_events_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-link);">{bi('Advisor reliability recommendations', 'Recomendaciones de confiabilidad de Advisor', 'Recomenda\u00e7\u00f5es de confiabilidade do Advisor')}{it('Reliability recommendations generated by Azure Advisor.', 'Recomendaciones de confiabilidad generadas por Azure Advisor.', 'Recomenda\u00e7\u00f5es de confiabilidade geradas pelo Azure Advisor.')}</div>
            {recs_block}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-link);">{bi('Score breakdown', 'Desglose de puntaje', 'Detalhamento da pontua\u00e7\u00e3o')}{it('How the reliability score is calculated, broken down by category and subcategory.', 'C\u00f3mo se calcula el puntaje de confiabilidad, desglosado por categor\u00eda y subcategor\u00eda.', 'Como a pontua\u00e7\u00e3o de confiabilidade \u00e9 calculada, detalhada por categoria e subcategoria.')}</div>
            {self._render_subcategory_table("Reliability")}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-link);">{bi('Action items', 'Elementos de acci\u00f3n', 'Itens de a\u00e7\u00e3o')}{it('Prioritized reliability actions, ranked by score impact.', 'Acciones de confiabilidad priorizadas, ordenadas por impacto en el puntaje.', 'A\u00e7\u00f5es de confiabilidade priorizadas, ordenadas pelo impacto na pontua\u00e7\u00e3o.')}</div>
            {self._render_actions_for_pillar("Reliability")}
        </div>"""

    def _render_performance_view(self) -> str:
        """Render performance efficiency: VM/SQL saturation and Advisor performance recommendations."""
        metrics = self.data.get("metrics", {})
        advisor = self.data.get("advisor", {})

        def real_rows(source, sheet_name):
            return [
                row for row in source.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def render_table(rows, columns, empty_en, empty_es, limit=15):
            if not rows:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            MAX_TABLE_ROWS = 300
            row_total = len(rows)
            rows = rows[:MAX_TABLE_ROWS]
            header = "".join(f"<th>{bi(label_en, label_es, LABEL_PT.get(label_en, label_en))}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for i, row in enumerate(rows):
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{escape(str(row.get(key, "")))}</td>'
                    for key, _, _, numeric in columns
                )
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                row_attr += _resource_link_attr(row)
                body += f"<tr{row_attr}>{cells}</tr>"
            more = ""
            if len(rows) > limit:
                more_text = bi(f"+ {len(rows) - limit} more", f"+ {len(rows) - limit} m\u00e1s", f"+ {len(rows) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            if row_total > MAX_TABLE_ROWS:
                more += f'<p class="table-truncated-note">{bi(f"Showing first {MAX_TABLE_ROWS} of {row_total} rows for performance - see the source export for the full list.", f"Mostrando las primeras {MAX_TABLE_ROWS} de {row_total} filas por rendimiento - consulte el archivo de origen para ver la lista completa.", f"Mostrando as primeiras {MAX_TABLE_ROWS} de {row_total} linhas por desempenho - consulte o arquivo de origem para a lista completa.")}</p>'
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>{more}'

        vm_metrics = real_rows(metrics, "VM_RightSizing")
        saturated_vm = [v for v in vm_metrics if "saturated" in str(v.get("Assessment", "")).lower()]
        sql_metrics = real_rows(metrics, "SQL_RightSizing")
        saturated_sql = [s for s in sql_metrics if "saturated" in str(s.get("Assessment", "")).lower()]
        perf_recs = real_rows(advisor, "Performance")

        score = self.engine.scores.get("Performance Efficiency", 0)
        stats = [
            (f'<span data-score="Performance Efficiency" data-score-suffix="/100">{score}/100</span>', bi("Performance Score", "Puntaje de Rendimiento", "Pontua\u00e7\u00e3o de Desempenho"), "security-accent"),
            (len(saturated_vm), bi("Saturated VMs (>80% CPU)", "VMs Saturadas (>80% CPU)", "VMs Saturadas (>80% CPU)"), "security-danger" if saturated_vm else ""),
            (len(saturated_sql), bi("Saturated SQL DBs", "BD SQL Saturadas", "BDs SQL Saturados"), "security-danger" if saturated_sql else ""),
            (len(perf_recs), bi("Advisor Performance Recs", "Recom. de Rendimiento", "Recom. de Desempenho"), ""),
        ]
        stat_html = "".join(
            f'<div class="security-stat"><div class="value {color_class}">{value}</div><div class="label">{label}</div></div>'
            for value, label, color_class in stats
        )

        vm_table = render_table(saturated_vm, [
            ("Name", "VM", "VM", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("VMSize", "Size", "Tama\u00f1o", False),
            ("AvgCPU_Pct", "Avg CPU %", "CPU Prom. %", True),
            ("MaxCPU_Pct", "Max CPU %", "CPU M\u00e1x. %", True),
        ], "No saturated VMs found.", "No se encontraron VMs saturadas.")
        sql_table = render_table(saturated_sql, [
            ("Name", "Database", "Base de Datos", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("SKU", "SKU", "SKU", False),
            ("MetricType", "Metric", "M\u00e9trica", False),
            ("AvgUsage_Pct", "Avg Usage %", "Uso Prom. %", True),
        ], "No saturated SQL databases found.", "No se encontraron bases de datos SQL saturadas.")
        recs_table = render_table(perf_recs, [
            ("impact", "Impact", "Impacto", False),
            ("impactedResource", "Resource", "Recurso", False),
            ("problem", "Problem", "Problema", False),
            ("solution", "Solution", "Soluci\u00f3n", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "No Advisor performance recommendations found.", "No se encontraron recomendaciones de rendimiento de Advisor.")

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow" style="color: var(--cp-warning);">PERFORMANCE EFFICIENCY</div>
                <h2>{bi('Workload scaling and utilization', 'Escalado y utilizaci\u00f3n de cargas de trabajo', 'Escalonamento e utiliza\u00e7\u00e3o de cargas de trabalho')}</h2>
                <p>{bi('Saturated VMs and SQL databases, plus Advisor performance recommendations in one view.', 'VMs y bases de datos SQL saturadas, m\u00e1s recomendaciones de rendimiento de Advisor en una sola vista.', 'VMs e bancos de dados SQL saturados, al\u00e9m das recomenda\u00e7\u00f5es de desempenho do Advisor em uma \u00fanica vis\u00e3o.')}</p>
            </div>
            <div class="security-score-mark" style="color: var(--cp-warning);"><span data-score="Performance Efficiency" data-score-suffix="/100">{score}/100</span></div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-warning);">{bi('Resource saturation', 'Saturaci\u00f3n de recursos', 'Satura\u00e7\u00e3o de recursos')}{it('CPU, memory and other utilization metrics showing resources running near capacity.', 'M\u00e9tricas de CPU, memoria y otros indicadores de uso que muestran recursos operando cerca del l\u00edmite de capacidad.', 'M\u00e9tricas de CPU, mem\u00f3ria e outros indicadores de uso mostrando recursos operando perto da capacidade m\u00e1xima.')}</div>
            <div class="security-two-column">
                <div class="security-panel"><h3>{bi('Virtual Machines', 'M\u00e1quinas Virtuales', 'M\u00e1quinas Virtuais')}</h3>{vm_table}</div>
                <div class="security-panel"><h3>{bi('SQL Databases', 'Bases de Datos SQL', 'Bancos de Dados SQL')}</h3>{sql_table}</div>
            </div>
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-warning);">{bi('Advisor performance recommendations', 'Recomendaciones de rendimiento de Advisor', 'Recomenda\u00e7\u00f5es de desempenho do Advisor')}{it('Performance recommendations generated by Azure Advisor.', 'Recomendaciones de rendimiento generadas por Azure Advisor.', 'Recomenda\u00e7\u00f5es de desempenho geradas pelo Azure Advisor.')}</div>
            {recs_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-warning);">{bi('Score breakdown', 'Desglose de puntaje', 'Detalhamento da pontua\u00e7\u00e3o')}{it('How the performance score is calculated, broken down by category and subcategory.', 'C\u00f3mo se calcula el puntaje de rendimiento, desglosado por categor\u00eda y subcategor\u00eda.', 'Como a pontua\u00e7\u00e3o de desempenho \u00e9 calculada, detalhada por categoria e subcategoria.')}</div>
            {self._render_subcategory_table("Performance Efficiency")}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-warning);">{bi('Action items', 'Elementos de acci\u00f3n', 'Itens de a\u00e7\u00e3o')}{it('Prioritized performance actions, ranked by score impact.', 'Acciones de rendimiento priorizadas, ordenadas por impacto en el puntaje.', 'A\u00e7\u00f5es de desempenho priorizadas, ordenadas pelo impacto na pontua\u00e7\u00e3o.')}</div>
            {self._render_actions_for_pillar("Performance Efficiency")}
        </div>"""

    def _render_operational_view(self) -> str:
        """Render operational excellence: diagnostics coverage, tagging, resource locks, policy compliance, and Advisor recommendations."""
        metrics = self.data.get("metrics", {})
        discovery = self.data.get("discovery", {})
        governance = self.data.get("governance", {})
        advisor = self.data.get("advisor", {})

        def real_rows(source, sheet_name):
            return [
                row for row in source.get(sheet_name, [])
                if not row.get("Result") and str(row.get("DataStatus", "")).lower() != "nodata"
            ]

        def as_int(value):
            try:
                return int(value or 0)
            except (TypeError, ValueError):
                return 0

        def render_table(rows, columns, empty_en, empty_es, limit=15):
            if not rows:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            MAX_TABLE_ROWS = 300
            row_total = len(rows)
            rows = rows[:MAX_TABLE_ROWS]
            header = "".join(f"<th>{bi(label_en, label_es, LABEL_PT.get(label_en, label_en))}</th>" for _, label_en, label_es, _ in columns)
            body = ""
            for i, row in enumerate(rows):
                cells = "".join(
                    f'<td class="{"num" if numeric else ""}">{escape(str(row.get(key, "")))}</td>'
                    for key, _, _, numeric in columns
                )
                row_attr = ' class="extra-row" hidden' if i >= limit else ""
                row_attr += _resource_link_attr(row)
                body += f"<tr{row_attr}>{cells}</tr>"
            more = ""
            if len(rows) > limit:
                more_text = bi(f"+ {len(rows) - limit} more", f"+ {len(rows) - limit} m\u00e1s", f"+ {len(rows) - limit} mais")
                less_text = bi("Show less", "Mostrar menos", "Mostrar menos")
                more = (
                    '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">'
                    f'<span class="more-label">{more_text}</span><span class="less-label">{less_text}</span></button>'
                )
            if row_total > MAX_TABLE_ROWS:
                more += f'<p class="table-truncated-note">{bi(f"Showing first {MAX_TABLE_ROWS} of {row_total} rows for performance - see the source export for the full list.", f"Mostrando las primeras {MAX_TABLE_ROWS} de {row_total} filas por rendimiento - consulte el archivo de origen para ver la lista completa.", f"Mostrando as primeiras {MAX_TABLE_ROWS} de {row_total} linhas por desempenho - consulte o arquivo de origem para a lista completa.")}</p>'
            return f'<div class="security-table-wrap"><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></div>{more}'

        diag = real_rows(metrics, "DiagnosticsCoverage")
        no_diag = [d for d in diag if str(d.get("HasDiagnostics", d.get("Gap", ""))).lower() in ("false", "no diagnostics configured")]

        TAGGABLE_SHEETS = ["VMs", "AppServices", "AKS", "VNets", "NSGs", "LoadBalancers",
                           "Firewalls", "PublicIPs", "Storage", "Databases", "KeyVaults"]
        untagged = []
        total_taggable = 0
        for sheet in TAGGABLE_SHEETS:
            for r in real_rows(discovery, sheet):
                total_taggable += 1
                tags = r.get("tags")
                if not tags or str(tags).strip().lower() in ("", "none", "{}"):
                    untagged.append({
                        "Type": sheet,
                        "Name": r.get("name", r.get("Name", "")),
                        "Subscription": r.get("Subscription", ""),
                        "ResourceGroup": r.get("resourceGroup", r.get("ResourceGroup", "")),
                    })

        locks = real_rows(governance, "Locks")
        policy_compliance = real_rows(governance, "PolicyCompliance")
        non_compliant = [c for c in policy_compliance if as_int(c.get("NonCompliantCount", 0)) > 0]
        opex_recs = real_rows(advisor, "OperationalExcellence")

        score = self.engine.scores.get("Operational Excellence", 0)
        stats = [
            (f'<span data-score="Operational Excellence" data-score-suffix="/100">{score}/100</span>', bi("OpEx Score", "Puntaje OpEx", "Pontua\u00e7\u00e3o OpEx"), "security-accent"),
            (len(no_diag), bi("Missing Diagnostic Settings", "Sin Configuraci\u00f3n de Diagn\u00f3stico", "Sem Configura\u00e7\u00e3o de Diagn\u00f3stico"), "security-warning" if no_diag else ""),
            (f"{len(untagged)}/{total_taggable}", bi("Untagged Resources", "Recursos sin Etiquetas", "Recursos sem Tags"), "security-warning" if untagged else ""),
            (len(locks), bi("Resource Locks", "Bloqueos de Recursos", "Bloqueios de Recursos"), "" if locks else "security-danger"),
            (len(non_compliant), bi("Non-Compliant Policy Evaluations", "Evaluaciones No Conformes", "Avalia\u00e7\u00f5es N\u00e3o Conformes"), "security-warning" if non_compliant else ""),
            (len(opex_recs), bi("Advisor OpEx Recs", "Recom. OpEx", "Recom. OpEx"), ""),
        ]
        stat_html = "".join(
            f'<div class="security-stat"><div class="value {color_class}">{value}</div><div class="label">{label}</div></div>'
            for value, label, color_class in stats
        )

        diag_table = render_table(no_diag, [
            ("Name", "Resource", "Recurso", False),
            ("Type", "Type", "Tipo", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
            ("Gap", "Gap", "Brecha", False),
        ], "All critical resources have diagnostic settings configured.", "Todos los recursos cr\u00edticos tienen configuraci\u00f3n de diagn\u00f3stico.")
        tags_table = render_table(untagged, [
            ("Type", "Resource Type", "Tipo de Recurso", False),
            ("Name", "Name", "Nombre", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("ResourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "All resources are tagged.", "Todos los recursos tienen etiquetas.")
        locks_table = render_table(locks, [
            ("name", "Name", "Nombre", False),
            ("lockLevel", "Level", "Nivel", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "No resource locks configured.", "No hay bloqueos de recursos configurados.")
        recs_table = render_table(opex_recs, [
            ("impact", "Impact", "Impacto", False),
            ("impactedResource", "Resource", "Recurso", False),
            ("problem", "Problem", "Problema", False),
            ("solution", "Solution", "Soluci\u00f3n", False),
            ("Subscription", "Subscription", "Suscripci\u00f3n", False),
            ("resourceGroup", "Resource Group", "Grupo de Recursos", False),
        ], "No Advisor operational excellence recommendations found.", "No se encontraron recomendaciones de excelencia operativa de Advisor.")

        return f"""
        <div class="security-summary-band">
            <div>
                <div class="security-eyebrow" style="color: var(--cp-accent);">OPERATIONAL EXCELLENCE</div>
                <h2>{bi('Observability, tagging, and governance hygiene', 'Observabilidad, etiquetado e higiene de gobernanza', 'Observabilidade, etiquetagem e higiene de governan\u00e7a')}</h2>
                <p>{bi('Diagnostic settings coverage, tagging compliance, resource locks, and Advisor operational recommendations in one view.', 'Cobertura de configuraci\u00f3n de diagn\u00f3stico, cumplimiento de etiquetado, bloqueos de recursos y recomendaciones operativas de Advisor en una sola vista.', 'Cobertura de configura\u00e7\u00e3o de diagn\u00f3stico, conformidade de etiquetagem, bloqueios de recursos e recomenda\u00e7\u00f5es operacionais do Advisor em uma \u00fanica vis\u00e3o.')}</p>
            </div>
            <div class="security-score-mark" style="color: var(--cp-accent);"><span data-score="Operational Excellence" data-score-suffix="/100">{score}/100</span></div>
        </div>
        <div class="security-stats">{stat_html}</div>

        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-accent);">{bi('Diagnostic settings coverage', 'Cobertura de configuraci\u00f3n de diagn\u00f3stico', 'Cobertura de configura\u00e7\u00e3o de diagn\u00f3stico')}{it('Percentage of resources sending diagnostic logs and metrics to a Log Analytics workspace or storage account.', 'Porcentaje de recursos que env\u00edan registros y m\u00e9tricas de diagn\u00f3stico a un workspace de Log Analytics o cuenta de almacenamiento.', 'Percentual de recursos que enviam logs e m\u00e9tricas de diagn\u00f3stico para um workspace do Log Analytics ou conta de armazenamento.')}</div>
            {diag_table}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-accent);">{bi('Untagged resources', 'Recursos sin etiquetas', 'Recursos sem tags')}{it('Resources missing required tags, which hinders cost allocation, ownership tracking and automation.', 'Recursos sin las etiquetas requeridas, lo que dificulta la asignaci\u00f3n de costos, el seguimiento de propietarios y la automatizaci\u00f3n.', 'Recursos sem as tags obrigat\u00f3rias, o que dificulta a aloca\u00e7\u00e3o de custos, o rastreamento de propriet\u00e1rios e a automa\u00e7\u00e3o.')}</div>
            {tags_table}
        </div>
        <div class="security-two-column">
            <div class="security-panel"><h3>{bi('Resource locks', 'Bloqueos de recursos', 'Bloqueios de recursos')}</h3>{locks_table}</div>
            <div class="security-panel"><h3>{bi('Advisor operational excellence recommendations', 'Recomendaciones de excelencia operativa de Advisor', 'Recomenda\u00e7\u00f5es de excel\u00eancia operacional do Advisor')}</h3>{recs_table}</div>
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-accent);">{bi('Score breakdown', 'Desglose de puntaje', 'Detalhamento da pontua\u00e7\u00e3o')}{it('How the operational excellence score is calculated, broken down by category and subcategory.', 'C\u00f3mo se calcula el puntaje de excelencia operativa, desglosado por categor\u00eda y subcategor\u00eda.', 'Como a pontua\u00e7\u00e3o de excel\u00eancia operacional \u00e9 calculada, detalhada por categoria e subcategoria.')}</div>
            {self._render_subcategory_table("Operational Excellence")}
        </div>
        <div class="section">
            <div class="section-title" style="border-bottom-color: var(--cp-accent);">{bi('Action items', 'Elementos de acci\u00f3n', 'Itens de a\u00e7\u00e3o')}{it('Prioritized operational excellence actions, ranked by score impact.', 'Acciones de excelencia operativa priorizadas, ordenadas por impacto en el puntaje.', 'A\u00e7\u00f5es de excel\u00eancia operacional priorizadas, ordenadas pelo impacto na pontua\u00e7\u00e3o.')}</div>
            {self._render_actions_for_pillar("Operational Excellence")}
        </div>"""

    def _render_progress_banner(self) -> str:
        """Render the 'since last scan' delta banner (overall/pillar score deltas + resolved
        items), driven by ScoreHistoryTracker. Renders a first-run note if there's no history yet."""
        previous = self.progress.get("previous")
        if not previous:
            return f"""<div class="progress-banner progress-banner-empty">
                <span>{bi("This is the first recorded scan for this environment. Future runs will show score trends and resolved items here.", "Este es el primer escaneo registrado para este entorno. Las pr\u00f3ximas ejecuciones mostrar\u00e1n tendencias de puntaje y elementos resueltos aqu\u00ed.", "Este \u00e9 o primeiro scan registrado para este ambiente. As pr\u00f3ximas execu\u00e7\u00f5es mostrar\u00e3o tend\u00eancias de pontua\u00e7\u00e3o e itens resolvidos aqui.")}</span>
            </div>"""

        def delta_badge(delta: Optional[int]) -> str:
            if delta is None:
                return '<span class="progress-delta progress-delta-neutral">–</span>'
            if delta > 0:
                return f'<span class="progress-delta progress-delta-up">▲ +{delta}</span>'
            if delta < 0:
                return f'<span class="progress-delta progress-delta-down">▼ {delta}</span>'
            return '<span class="progress-delta progress-delta-neutral">= 0</span>'

        prev_date = str(previous.get("timestamp", ""))[:16].replace("T", " ")
        overall_delta_html = delta_badge(self.progress.get("overall_delta"))
        pillar_chips = "".join(
            f'<span class="progress-pillar-chip">{bi(str(p), str(PILLAR_ES.get(p, p)), str(PILLAR_PT.get(p, p)))} {delta_badge(d)}</span>'
            for p, d in self.progress.get("pillar_deltas", {}).items()
        )
        resolved = self.progress.get("resolved", [])
        new_count = self.progress.get("new_count", 0)
        resolved_html = ""
        if resolved:
            resolved_items = "".join(
                f'<li><span class="severity-badge" style="background:{DashboardConfig.SEVERITY_COLORS.get(r["severity"], "var(--cp-text-muted)")};">{r["severity"].upper()}</span> {escape(r["title"])} <span class="pillar-tag">{escape(r["pillar"])}</span></li>'
                for r in resolved
            )
            resolved_html = f"""<details class="progress-resolved">
                <summary>{bi(f"{len(resolved)} action item(s) resolved since last scan", f"{len(resolved)} elemento(s) resuelto(s) desde el \u00faltimo escaneo", f"{len(resolved)} item(ns) resolvido(s) desde o \u00faltimo scan")}</summary>
                <ul class="progress-resolved-list">{resolved_items}</ul>
            </details>"""

        new_html = ""
        if new_count:
            new_html = f'<span class="progress-new-count">{bi(f"+{new_count} new since last scan", f"+{new_count} nuevos desde el \u00faltimo escaneo", f"+{new_count} novos desde o \u00faltimo scan")}</span>'

        return f"""<div class="progress-banner">
            <div class="progress-banner-row">
                <span class="progress-title">{bi("Since last scan", "Desde el \u00faltimo escaneo", "Desde o \u00faltimo scan")} ({escape(prev_date)})</span>
                <span class="progress-overall">{bi("Overall score", "Puntaje general", "Pontua\u00e7\u00e3o geral")}: {overall_delta_html}</span>
                {new_html}
            </div>
            <div class="progress-pillar-row">{pillar_chips}</div>
            {resolved_html}
        </div>"""

    def _render_score_breakdown(self) -> str:
        """Render the 'why' behind each pillar score per the Advisor Score model."""
        b = self.engine.breakdowns
        sections = []
        pillar_label = lambda p: bi(str(p), str(PILLAR_ES.get(p, p)), str(PILLAR_PT.get(p, p)))
        insufficient_note = bi(
            "Insufficient data — no discovered resource inventory (01_Discovery/ResourceSummary) to compute this pillar's score.",
            "Datos insuficientes: no hay inventario de recursos descubiertos (01_Discovery/ResourceSummary) para calcular el puntaje de este pilar.",
            "Dados insuficientes: não há inventário de recursos descobertos (01_Discovery/ResourceSummary) para calcular a pontuação deste pilar."
        )

        def subcategory_card(pillar):
            info = b.get(pillar, {})
            if not info.get("available"):
                return f"""<div class="breakdown-card">
                <h4>{pillar_label(pillar)}</h4>
                <p class="breakdown-note">{insufficient_note}</p>
            </div>"""
            def pct_bar_cell(pct):
                color = "var(--cp-success)" if pct >= 80 else "var(--cp-warning)" if pct >= 50 else "var(--cp-danger)"
                return f"<td class='num sc-pct'><div class='inline-bar-wrap'><div class='inline-bar-track'><div class='inline-bar-fill' style='width:{pct}%; background:{color};'></div></div><span>{pct}%</span></div></td>"

            rows_html = "".join(
                f"<tr data-subcat=\"{escape(pillar)}|{escape(r['subcategory'])}\"><td>{bi(str(r['subcategory']), str(SUBCATEGORY_ES.get(r['subcategory'], r['subcategory'])), str(SUBCATEGORY_PT.get(r['subcategory'], r['subcategory'])))}</td><td class='num'>{r['weight']}</td>"
                f"<td class='num sc-ht'>{r['healthy']}/{r['total']}</td>{pct_bar_cell(r['pct'])}</tr>"
                for r in info["rows"]
            )
            return f"""<div class="breakdown-card">
                <h4>{pillar_label(pillar)} <span class="breakdown-score" data-score="{pillar}" data-score-suffix="/100">{self.engine.scores[pillar]}/100</span></h4>
                <div class="breakdown-table-wrap">
                    <table class="subcategory-table">
                        <colgroup><col class="subcat-name-col"><col class="subcat-num-col"><col class="subcat-num-col"><col class="subcat-score-col"></colgroup>
                        <thead><tr><th>{bi('Subcategory', 'Subcategor\u00eda', 'Subcategoria')}</th><th>{bi('Weight', 'Peso', 'Peso')}</th><th>{bi('Healthy/Total', 'Saludable/Total', 'Saud\u00e1vel/Total')}</th><th>{bi('Score', 'Puntaje', 'Pontua\u00e7\u00e3o')}</th></tr></thead>
                        <tbody>{rows_html}</tbody>
                    </table>
                </div>
            </div>"""

        sections.append(subcategory_card("Reliability"))
        sections.append(subcategory_card("Performance Efficiency"))
        sections.append(subcategory_card("Operational Excellence"))

        security = b.get("Security", {})
        if security.get("available"):
            findings_html = "".join(
                f"<tr><td>{bi(label_en, label_es, label_pt)}</td><td class='num'>{count}</td></tr>"
                for label_en, label_es, label_pt, count in security.get("findings", [])
            )
            scored = f"<span data-secure-scored>{security.get('subscriptions_scored', 0)}</span>"
            pct = f"<strong data-secure-pct>{security.get('secure_score_pct')}%</strong>"
            note = bi(
                f"Microsoft Defender Secure Score average across {scored} subscription(s): {pct}.",
                f"Promedio del Secure Score de Microsoft Defender en {scored} suscripci\u00f3n(es): {pct}.",
                f"M\u00e9dia do Secure Score da Microsoft Defender em {scored} assinatura(s): {pct}."
            )
            sections.append(f"""<div class="breakdown-card">
                <h4>{pillar_label('Security')} <span class="breakdown-score" data-score="Security" data-score-suffix="/100">{self.engine.scores['Security']}/100</span></h4>
                <p class="breakdown-note">{note}</p>
                <p class="breakdown-note secure-gap-note" data-pillar-note="Security" hidden></p>
                <table>
                    <thead><tr><th>{bi('Finding Category', 'Categor\u00eda de Hallazgo', 'Categoria da Descoberta')}</th><th>{bi('Count', 'Cantidad', 'Quantidade')}</th></tr></thead>
                    <tbody>{findings_html}</tbody>
                </table>
            </div>""")
        else:
            note = bi(
                "Insufficient data \u2014 no Defender Secure Score records found in 04_Governance/SecureScores.",
                "Datos insuficientes: no se encontraron registros de Secure Score de Defender en 04_Governance/SecureScores.",
                "Dados insuficientes: nenhum registro de Secure Score do Defender encontrado em 04_Governance/SecureScores."
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
                source_score = round_half_up((source_healthy / cost["total"]) * 100)
                source_rows.append(
                    f"<tr data-cost-source=\"{escape(name)}\"><td>{bi(str(name), str(COST_SOURCE_ES.get(name, name)), str(COST_SOURCE_PT.get(name, name)))}</td>"
                    f"<td class='num cs-count'>{source['count']} / {affected}</td>"
                    f"<td class='num cs-healthy'>{source_healthy}/{cost['total']}</td>"
                    f"<td class='num cs-score'>{source_score}%</td></tr>"
                )
            combined_label = bi("Combined (deduplicated)", "Combinado (sin duplicados)", "Combinado (sem duplicatas)")
            combined_findings = sum(source["count"] for source in cost.get("sources", []))
            source_rows.append(
                f"<tr data-cost-combined><td><strong>{combined_label}</strong></td>"
                f"<td class='num'><strong class='cs-count'>{combined_findings} / {cost['impacted']}</strong></td>"
                f"<td class='num'><strong class='cs-healthy'>{cost['healthy']}/{cost['total']}</strong></td>"
                f"<td class='num'><strong data-score=\"Cost Optimization\" data-score-suffix=\"%\">{self.engine.scores['Cost Optimization']}%</strong></td></tr>"
            )
            sources_html = "".join(source_rows)
            note = bi(
                "Findings/Affected shows finding count first and unique affected scopes or resources second. The combined row deduplicates identifiers.",
                "Hallazgos/Afectados muestra primero los hallazgos y luego los ámbitos o recursos únicos afectados. La fila combinada elimina duplicados.",
                "Descobertas/Afetados mostra primeiro a contagem de descobertas e depois os escopos ou recursos únicos afetados. A linha combinada remove duplicatas."
            )
            sections.append(f"""<div class="breakdown-card cost-breakdown-card">
                <h4>{pillar_label('Cost Optimization')} <span class="breakdown-score" data-score="Cost Optimization" data-score-suffix="/100">{self.engine.scores['Cost Optimization']}/100</span></h4>
                <p class="breakdown-note">{note}</p>
                <div class="breakdown-table-wrap">
                    <table>
                        <colgroup><col class="cost-source-col"><col class="cost-findings-col"><col><col></colgroup>
                        <thead><tr><th>{bi('Finding Source', 'Fuente del Hallazgo', 'Fonte da Descoberta')}</th><th>{bi('Findings/Affected', 'Hallazgos/Afectados', 'Descobertas/Afetados')}</th><th>{bi('Healthy/Total', 'Saludable/Total', 'Saudável/Total')}</th><th>{bi('Score', 'Puntaje', 'Pontuação')}</th></tr></thead>
                        <tbody>{sources_html}</tbody>
                    </table>
                </div>
            </div>""")
        else:
            note = bi(
                "Insufficient data — no discovered resource inventory to compute this pillar's score.",
                "Datos insuficientes: no hay inventario de recursos descubiertos para calcular el puntaje de este pilar.",
                "Dados insuficientes: não há inventário de recursos descobertos para calcular a pontuação deste pilar."
            )
            sections.append(f"""<div class="breakdown-card">
                <h4>{pillar_label('Cost Optimization')}</h4>
                <p class="breakdown-note">{note}</p>
            </div>""")

        return "".join(sections)
    
    def generate(self) -> str:
        overall = self.engine.get_overall_score()
        overall_color = 'var(--cp-success)' if overall >= 80 else 'var(--cp-warning)' if overall >= 50 else 'var(--cp-danger)'
        scores = self.engine.scores
        progress_banner_html = self._render_progress_banner()
        score_breakdown_html = self._render_score_breakdown()
        governance_view_html = self._render_governance_view()
        security_view_html = self._render_security_view()
        cost_view_html = self._render_cost_view()
        reliability_view_html = self._render_reliability_view()
        performance_view_html = self._render_performance_view()
        operational_view_html = self._render_operational_view()
        action_items = sorted(self.engine.action_items, 
                            key=lambda x: {"critical": 0, "high": 1, "medium": 2, "low": 3, "info": 4}.get(x["severity"], 5))
        
        # Count stats
        total_resources = sum(int(r.get("Count", 0) or 0) for r in self.data.get("discovery", {}).get("ResourceSummary", []))
        total_subs = len(self.data.get("discovery", {}).get("Subscriptions", []) or self.data.get("governance", {}).get("Subscriptions", []))
        total_actions = len(action_items)
        critical_actions = len([a for a in action_items if a["severity"] in ("critical", "high")])

        # Dashboard-wide subscription filter (checkbox bar shown under the tab strip, visible
        # on every tab) - drives both the Cost tab's pivot (via its own checkboxes, kept in
        # sync by initGlobalSubscriptionFilter()) and row-hiding on any other table across all
        # tabs that carries a "Subscription" column.
        sub_id_by_name = {
            s.get("name"): str(s.get("subscriptionId") or "") for s in (
                self.data.get("discovery", {}).get("Subscriptions", [])
                or self.data.get("governance", {}).get("Subscriptions", [])
            ) if s.get("name")
        }
        all_sub_names = sorted(sub_id_by_name)
        if len(all_sub_names) > 1:
            global_sub_checks = "".join(
                f'<label class="global-sub-check"><input type="checkbox" value="{escape(s)}" data-sub-id="{escape(sub_id_by_name[s])}" checked> {escape(s)}</label>'
                for s in all_sub_names
            )
            global_sub_filter_html = f"""<div class="global-sub-filter-bar" id="globalSubFilterBar">
    <span class="global-sub-filter-label">{bi('Filter by subscription', 'Filtrar por suscripci\u00f3n', 'Filtrar por assinatura')}:</span>
    {global_sub_checks}
    <button type="button" class="cost-pivot-selectall" data-mode="all">{bi("All", "Todas", "Todas")}</button>
    <button type="button" class="cost-pivot-selectall" data-mode="none">{bi("None", "Ninguna", "Nenhuma")}</button>
    <span class="global-sub-filter-count"></span>
    <span class="global-sub-filter-note">{bi("Filters tables, action items, recommendations, and pillar scores.", "Filtra tablas, elementos de acci\u00f3n, recomendaciones y puntajes por pilar.", "Filtra tabelas, itens de a\u00e7\u00e3o, recomenda\u00e7\u00f5es e pontua\u00e7\u00f5es por pilar.")}</span>
</div>"""
        else:
            global_sub_filter_html = ""

        # Chart.js chart configs (canvasId -> config), consumed by initAllCharts() in the
        # embedded <script> - centralizes all chart definitions instead of scattering
        # per-section inline scripts. Alias to the shared instance dict so other
        # _render_*_view() methods (already called earlier, e.g. security_view_html)
        # can also contribute entries.
        chart_data = self._chart_data

        # Build action items HTML
        actions_html = "".join(self._render_action_card(action, i, show_pillar=True) for i, action in enumerate(action_items, 1))

        # Action items summary charts: by severity (donut) and by pillar (horizontal bar)
        severity_order = ["critical", "high", "medium", "low", "info"]
        severity_counts = defaultdict(int)
        for a in action_items:
            severity_counts[a.get("severity", "info")] += 1
        severity_present = [s for s in severity_order if severity_counts.get(s, 0) > 0]
        if severity_present:
            chart_data["actionSeverity"] = {
                "type": "donut", "canvasId": "actionSeverityChart",
                "keys": severity_present,
                "labels": {
                    lang: [SEVERITY_LABELS.get(s, {}).get(lang, s.capitalize()) for s in severity_present]
                    for lang in ("en", "es", "pt")
                },
                "values": [severity_counts[s] for s in severity_present],
                "colorVars": [_bare_var(DashboardConfig.SEVERITY_COLORS.get(s)) for s in severity_present],
            }
        pillar_action_counts = defaultdict(int)
        for a in action_items:
            pillar_action_counts[a.get("pillar", "")] += 1
        pillars_with_actions = [p for p in scores.keys() if pillar_action_counts.get(p, 0) > 0]
        if pillars_with_actions:
            chart_data["actionByPillar"] = {
                "type": "hbar", "canvasId": "actionByPillarChart",
                "keys": pillars_with_actions,
                "labels": {
                    "en": pillars_with_actions,
                    "es": [PILLAR_ES.get(p, p) for p in pillars_with_actions],
                    "pt": [PILLAR_PT.get(p, p) for p in pillars_with_actions],
                },
                "values": [pillar_action_counts[p] for p in pillars_with_actions],
                "colorVars": [PILLAR_COLOR_VARS[p] for p in pillars_with_actions],
            }

        # Build pillar score cards
        pillar_cards = ""
        for pillar, sc in scores.items():
            color = DashboardConfig.PILLAR_COLORS[pillar]
            pillar_cards += f"""
            <div class="pillar-card" data-pillar-card="{pillar}">
                <div class="pillar-score" data-score="{pillar}" style="color:{color};">{sc}</div>
                <div class="pillar-bar">
                    <div class="pillar-bar-fill" style="width:{sc}%; background:{color};"></div>
                </div>
                <div class="pillar-name">{bi(pillar, PILLAR_ES.get(pillar, pillar), PILLAR_PT.get(pillar, pillar))}</div>
                {'<div class="pillar-note" data-pillar-note="Security" hidden></div>' if pillar == 'Security' else ''}
            </div>"""

        # WAF Pillar Scores comparison chart (Overview, below the pillar cards)
        chart_data["pillarScores"] = {
            "type": "hbar", "canvasId": "pillarScoresChart",
            "keys": list(scores.keys()),
            "labels": {
                "en": list(scores.keys()),
                "es": [PILLAR_ES.get(p, p) for p in scores.keys()],
                "pt": [PILLAR_PT.get(p, p) for p in scores.keys()],
            },
            "values": list(scores.values()),
            "colorVars": [PILLAR_COLOR_VARS[p] for p in scores.keys()],
            "opts": {"max": 100, "suffix": "/100"},
        }

        # Build resource summary table
        resource_summary = self.data.get("discovery", {}).get("ResourceSummary", [])
        top_resources = sorted(resource_summary, key=lambda x: int(x.get("Count", 0) or 0), reverse=True)[:20]
        resource_rows = ""
        for r in top_resources:
            rtype = str(r.get("type", "")).split("/")[-1] if "/" in str(r.get("type", "")) else r.get("type", "")
            resource_rows += f"<tr><td>{r.get('type', '')}</td><td>{r.get('location', '')}</td><td class='num'>{r.get('Count', 0)}</td></tr>"

        def render_bar_chart(items, empty_en, empty_es, max_bars=10):
            """Horizontal bar chart reusing the same .cost-bar-* CSS already used for budgets/cost
            breakdowns elsewhere - items is a list of (label, count) tuples, already summed."""
            if not items:
                return f'<p class="empty-state">{bi_empty(empty_en, empty_es)}</p>'
            top = sorted(items, key=lambda kv: kv[1], reverse=True)[:max_bars]
            max_val = max(v for _, v in top) or 1
            rows_html = "".join(
                f"""<div class="cost-bar-row">
                    <span class="cost-bar-label">{escape(str(label))}</span>
                    <div class="cost-bar-track"><div class="cost-bar-fill" style="width:{(v / max_val * 100):.0f}%;"></div></div>
                    <span class="cost-bar-value">{v:,}</span>
                </div>"""
                for label, v in top
            )
            return f'<div class="cost-bar-chart">{rows_html}</div>'

        resource_type_totals = defaultdict(int)
        for r in resource_summary:
            rtype = str(r.get("type", "")).split("/")[-1] if "/" in str(r.get("type", "")) else str(r.get("type", ""))
            resource_type_totals[rtype or "Unknown"] += int(r.get("Count", 0) or 0)
        resource_types_chart = render_bar_chart(
            list(resource_type_totals.items()),
            "No resource data found.", "No se encontraron datos de recursos.",
        )

        # Advisor summary
        advisor_summary = self.data.get("advisor", {}).get("SummaryByCategory", [])
        advisor_rows = ""
        for a in advisor_summary:
            advisor_rows += f"<tr><td>{a.get('category', '')}</td><td>{a.get('impact', '')}</td><td class='num'>{a.get('Count', 0)}</td></tr>"

        advisor_category_totals = defaultdict(int)
        for a in advisor_summary:
            advisor_category_totals[str(a.get("category", "") or "Unknown")] += int(a.get("Count", 0) or 0)
        advisor_summary_chart = render_bar_chart(
            list(advisor_category_totals.items()),
            "No Advisor data found.", "No se encontraron datos de Advisor.",
        )

        chart_data_json = json.dumps(chart_data)
        score_model_json = json.dumps(self.engine.score_model_payload(), ensure_ascii=False).replace("</", "<\\/")
        chartjs_lib = self._vendor_js("chart.umd.min.js")

        html = f"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Azure Well-Architected Framework - Consolidated Dashboard</title>
<link rel="icon" type="image/png" media="(prefers-color-scheme: light)" href="data:image/png;base64,{self._favicon_png_base64('dark')}">
<link rel="icon" type="image/png" media="(prefers-color-scheme: dark)" href="data:image/png;base64,{self._favicon_png_base64('white')}">
<script>
    (() => {{
        const param = new URLSearchParams(window.location.search).get("scoutTheme");
        let saved = null;
        try {{ saved = localStorage.getItem("wafDashboardTheme"); }} catch (e) {{}}
        const theme = ["light", "dark"].includes(param)
            ? param
            : (["light", "dark"].includes(saved) ? saved : "light");
        document.documentElement.setAttribute("data-theme", theme);
    }})();
</script>
<style>
:root {{
    color-scheme: light;
    --cp-bg: #eeeeee;
    --cp-bg-elevated: #ffffff;
    --cp-surface: #ffffff;
    --cp-surface-soft: #eeeeee;
    --cp-border: rgba(0, 0, 0, 0.08);
    --cp-border-strong: rgba(0, 0, 0, 0.16);
    --cp-text: #000000;
    --cp-text-muted: #666666;
    --cp-text-soft: #979797;
    --cp-accent: #0088cc;
    --cp-accent-hover: #269fdb;
    --cp-accent-soft: rgba(0, 136, 204, 0.08);
    --cp-accent-fg: #ffffff;
    --cp-success: #3f8f5c;
    --cp-critical: #d0102b;
    --cp-danger: #ff8f5f;
    --cp-warning: #666666;
    --cp-link: #0088cc;
    --cp-shadow: 0 2px 8px rgba(0, 0, 0, 0.04), 0 12px 32px rgba(0, 0, 0, 0.06);
    --cp-overlay: rgba(255, 255, 255, 0.75);
    --cp-panel: rgba(255, 255, 255, 0.72);
    --cp-panel-strong: rgba(255, 255, 255, 0.94);
    --cp-sheen: rgba(255, 255, 255, 0.6);
    --cp-highlight: rgba(0, 136, 204, 0.1);
    --overview-header-start: #eeeeee;
    --overview-header-end: #eeeeee;
    --cp-radius-lg: 22px;
    --cp-radius-md: 16px;
    --cp-radius-sm: 12px;
    --cp-ease: cubic-bezier(0.25, 0.1, 0.25, 1);
}}
html[data-theme="dark"] {{
    color-scheme: dark;
    --cp-bg: #000000;
    --cp-bg-elevated: #12192f;
    --cp-surface: #12192f;
    --cp-surface-soft: #1c1c1e;
    --cp-border: rgba(255, 255, 255, 0.08);
    --cp-border-strong: rgba(255, 255, 255, 0.16);
    --cp-text: #eeeeee;
    --cp-text-muted: #979797;
    --cp-text-soft: #666666;
    --cp-accent: #33a3dd;
    --cp-accent-hover: #4fb1e5;
    --cp-accent-soft: rgba(51, 163, 221, 0.16);
    --cp-accent-fg: #ffffff;
    --cp-success: #5aa876;
    --cp-critical: #e8404f;
    --cp-danger: #ff8f5f;
    --cp-warning: #979797;
    --cp-link: #33a3dd;
    --cp-shadow: 0 2px 8px rgba(0, 0, 0, 0.3), 0 12px 32px rgba(0, 0, 0, 0.4);
    --cp-overlay: rgba(18, 25, 47, 0.75);
    --cp-panel: rgba(18, 25, 47, 0.72);
    --cp-panel-strong: rgba(18, 25, 47, 0.94);
    --cp-sheen: rgba(255, 255, 255, 0.04);
    --cp-highlight: rgba(51, 163, 221, 0.16);
    --overview-header-start: #12192f;
    --overview-header-end: #12192f;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Segoe UI", Helvetica, Arial, sans-serif;
        background: var(--cp-bg); color: var(--cp-text);
    line-height: 1.47059; -webkit-font-smoothing: antialiased; letter-spacing: -0.01em;
}}
.dashboard {{ max-width: 1400px; margin: 0 auto; padding: 32px 24px 24px; }}

/* Header */
.header {{
    background: linear-gradient(135deg, #063466 0%, #00152e 100%);
    color: #ffffff; padding: 40px 40px 32px; border-radius: var(--cp-radius-lg);
    border: 1px solid rgba(255, 255, 255, 0.08); box-shadow: var(--cp-shadow);
    margin-bottom: 24px; position: relative;
}}
.header h1 {{ font-size: 30px; font-weight: 700; letter-spacing: -0.02em; margin: 0; min-width: 0; }}
.header-brand {{ display: flex; align-items: center; gap: 24px; margin-bottom: 12px; }}
.header-logo {{ width: 96px; height: 96px; flex-shrink: 0; }}
.header-brand .az-icon.header-logo {{ width: 96px; height: 96px; }}
.header-title-sub {{ display: block; font-size: 22px; font-weight: 700; margin-top: 2px; }}
.header .subtitle {{ color: rgba(255, 255, 255, 0.78); font-size: 14px; font-weight: 400; }}

/* Dashboard views (segmented control) */
.view-tabs {{
    display: flex; flex-wrap: wrap; justify-content: space-evenly; width: 100%; gap: 10px; padding: 4px;
    margin: 0 auto 24px; background: var(--cp-surface-soft); border: 1px solid var(--cp-border); border-radius: 980px;
    position: relative;
}}
.view-tab-thumb {{
    position: absolute; top: 4px; left: 0; height: calc(100% - 8px); width: 0;
    background: var(--cp-surface); border-radius: 980px; box-shadow: 0 1px 4px rgba(0,0,0,0.12);
    transition: transform 0.4s var(--cp-ease), width 0.4s var(--cp-ease), top 0.4s var(--cp-ease), height 0.4s var(--cp-ease);
    will-change: transform, width, top, height; z-index: 0; pointer-events: none;
}}
.view-tab {{
    border: 0; background: transparent; color: var(--cp-text-muted); cursor: pointer;
    padding: 8px 20px; border-radius: 980px; font: inherit; font-size: 13px; font-weight: 590;
    transition: color 0.25s var(--cp-ease); position: relative; z-index: 1;
}}
.view-tab[aria-selected="true"] {{ color: var(--cp-text); }}
.dashboard-view {{ transition: opacity 0.22s var(--cp-ease), transform 0.22s var(--cp-ease); opacity: 1; transform: translateY(0); }}
.dashboard-view.view-fade-out {{ opacity: 0; transform: translateY(6px); }}
.dashboard-view.view-fade-in {{ opacity: 0; transform: translateY(6px); }}
.dashboard-view[hidden] {{ display: none; }}

/* Dedicated security view */
.security-summary-band {{
    display: flex; justify-content: space-between; align-items: center; gap: 24px;
    background: var(--cp-surface); color: var(--cp-text); padding: 28px 32px;
    border: 1px solid var(--cp-border); border-radius: var(--cp-radius-lg); box-shadow: var(--cp-shadow);
    margin-bottom: 18px;
}}
.security-summary-band h2 {{ font-size: 26px; font-weight: 600; letter-spacing: -0.02em; margin: 2px 0 6px; }}
.security-summary-band p {{ color: var(--cp-text-muted); font-size: 13px; }}
.security-eyebrow {{ color: var(--cp-danger); font-size: 10px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; }}
.security-score-mark {{ font-size: 44px; font-weight: 700; letter-spacing: -0.02em; color: var(--cp-danger); white-space: nowrap; }}
.security-stats {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; margin-bottom: 24px; }}
.security-stat {{
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md);
    padding: 18px; min-width: 0; box-shadow: var(--cp-shadow); transition: transform 0.2s var(--cp-ease);
}}
.security-stat:hover {{ transform: translateY(-2px); }}
.security-stat .value {{ font-size: 27px; font-weight: 700; letter-spacing: -0.02em; }}
.security-stat .label {{ color: var(--cp-text-muted); font-size: 11px; margin-top: 4px; }}
.security-accent {{ color: var(--cp-link); }}
.security-danger {{ color: var(--cp-danger); }}
.security-warning {{ color: var(--cp-warning); }}
.security-title {{ border-bottom-color: var(--cp-danger); }}
.security-two-column {{ display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 18px; margin-bottom: 24px; }}
.security-panel {{
    background: var(--cp-surface); padding: 20px; min-width: 0;
    border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md); box-shadow: var(--cp-shadow);
}}
.security-panel h3 {{ font-size: 14px; font-weight: 590; margin-bottom: 12px; }}
.security-table-wrap {{ width: 100%; overflow-x: auto; }}
.security-table-wrap table {{ min-width: 540px; }}

/* Service x Month cost pivot (replaces a long flat Service/Month/Cost list) */
.cost-pivot-toolbar {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px 12px; margin-bottom: 12px; font-size: 13px; }}
.cost-pivot-toolbar-label {{ color: var(--cp-text-muted); font-weight: 600; }}
.cost-pivot-sub-check {{
    display: inline-flex; align-items: center; gap: 6px; padding: 5px 10px; border-radius: 999px;
    border: 1px solid var(--cp-border); background: var(--cp-surface); font-size: 12.5px; cursor: pointer;
}}
.cost-pivot-sub-check input {{ margin: 0; }}
.cost-pivot-selectall {{
    padding: 5px 10px; border-radius: 999px; border: 1px solid var(--cp-border); background: transparent;
    color: var(--cp-text-muted); font-size: 12.5px; cursor: pointer;
}}
.cost-pivot-selectall:hover {{ color: var(--cp-accent); border-color: var(--cp-accent); }}
.cost-pivot-sub-badge {{
    display: inline-flex; align-items: center; padding: 5px 10px; border-radius: 999px;
    border: 1px solid var(--cp-border); background: var(--cp-surface); font-size: 12.5px; font-weight: 600;
}}

/* Dashboard-wide subscription filter (under the tab strip, visible on every tab) */
.global-sub-filter-bar {{
    display: flex; flex-wrap: wrap; align-items: center; gap: 8px 12px;
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md);
    box-shadow: var(--cp-shadow); padding: 10px 20px; margin: 0 0 18px;
}}
.global-sub-filter-label {{ color: var(--cp-text-muted); font-size: 12.5px; font-weight: 700; }}
.global-sub-filter-count {{ font-size: 11px; font-weight: 700; color: var(--cp-accent); }}
.global-sub-filter-note {{ flex-basis: 100%; font-size: 11px; color: var(--cp-text-muted); }}
.is-sub-hidden {{ display: none !important; }}
.action-scope {{ margin-top: 6px; font-size: 11px; font-weight: 600; color: var(--cp-accent); }}
.global-sub-check {{
    display: inline-flex; align-items: center; gap: 6px; padding: 5px 10px; border-radius: 999px;
    border: 1px solid var(--cp-border); background: var(--cp-bg); font-size: 12.5px; cursor: pointer;
}}
.global-sub-check input {{ margin: 0; }}

/* Rows whose data carries an ARM resource id can be double-clicked to open in the Azure Portal */
.security-table-wrap tbody tr[data-resource-id] {{ cursor: pointer; }}
.security-table-wrap tbody tr[data-resource-id]:hover {{ background: var(--cp-bg); }}
.cost-pivot-wrap {{ width: 100%; overflow-x: auto; }}
.cost-pivot-wrap table {{ min-width: 540px; width: 100%; border-collapse: collapse; }}
.cost-pivot-wrap th, .cost-pivot-wrap td {{ padding: 8px 12px; border-bottom: 1px solid var(--cp-border); text-align: left; }}
.cost-pivot-wrap th.num, .cost-pivot-wrap td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
.cost-pivot-total-row td {{ font-weight: 700; background: var(--cp-bg); border-top: 2px solid var(--cp-border); }}

/* Per-table filter toolbar (auto-injected by initTableFilters()) */
.table-filter-bar {{ display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin-bottom: 10px; }}
.table-filter-select, .table-filter-search {{
    font-size: 12px; padding: 5px 10px; border-radius: 8px; border: 1px solid var(--cp-border);
    background: var(--cp-surface); color: var(--cp-text); min-width: 130px;
}}
.table-filter-search {{ flex: 1 1 160px; min-width: 160px; }}
.table-filter-count {{ font-size: 11px; color: var(--cp-text-muted); margin-left: auto; white-space: nowrap; }}
.table-filter-empty {{ padding: 14px 0; font-size: 12px; color: var(--cp-text-muted); display: none; }}

/* Azure-Portal-style info tooltip: small "i" circle, bubble only shown on hover/focus */
.info-tip {{
    display: inline-flex; align-items: center; justify-content: center; position: relative;
    width: 16px; height: 16px; border-radius: 50%; border: 1px solid var(--cp-text-muted);
    color: var(--cp-text-muted); font-size: 10px; font-weight: 700; font-style: italic;
    margin-left: 8px; cursor: help; vertical-align: middle; flex: 0 0 auto;
}}
.info-tip:hover, .info-tip:focus {{ border-color: var(--cp-accent); color: var(--cp-accent); outline: none; }}
.info-tip-bubble {{
    position: absolute; bottom: calc(100% + 8px); left: 50%; transform: translateX(-50%) translateY(4px);
    background: var(--cp-text); color: var(--cp-bg); padding: 8px 12px; border-radius: 8px;
    font-size: 12px; font-weight: 400; font-style: normal; line-height: 1.4; white-space: normal;
    width: max-content; max-width: 260px; text-align: left;
    opacity: 0; visibility: hidden; pointer-events: none;
    transition: opacity 0.15s var(--cp-ease), transform 0.15s var(--cp-ease); z-index: 20;
    box-shadow: 0 6px 20px rgba(0,0,0,0.25);
}}
.info-tip-bubble::after {{
    content: ""; position: absolute; top: 100%; left: 50%; transform: translateX(-50%);
    border: 6px solid transparent; border-top-color: var(--cp-text);
}}
.info-tip:hover .info-tip-bubble, .info-tip:focus .info-tip-bubble {{
    opacity: 1; visibility: visible; transform: translateX(-50%) translateY(0);
}}
.empty-state {{ color: var(--cp-text-muted); font-size: 12px; padding: 16px 0; }}
.source-status {{ display: inline-block; padding: 3px 10px; border-radius: 980px; font-size: 10px; font-weight: 700; }}
.source-available, .source-nodata {{ background: rgba(48, 209, 88, 0.12); color: var(--cp-success); }}
.source-partial {{ background: rgba(255, 149, 0, 0.14); color: var(--cp-warning); }}
.source-forbidden, .source-unavailable, .source-error {{ background: rgba(255, 59, 48, 0.12); color: var(--cp-danger); }}
.source-skipped {{ background: var(--cp-surface-soft); color: var(--cp-text-muted); }}
.security-success {{ color: var(--cp-success); }}
.cost-dist {{ display: flex; flex-direction: column; gap: 6px; margin-bottom: 14px; }}
.cost-dist-row {{ display: grid; grid-template-columns: minmax(90px, auto) 1fr 24px; align-items: center; gap: 8px; font-size: 11px; }}
.cost-dist-label {{ color: var(--cp-text-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.cost-dist-track {{ height: 6px; border-radius: 980px; background: var(--cp-surface-soft); overflow: hidden; }}
.cost-dist-fill {{ height: 100%; border-radius: 980px; }}
.cost-dist-count {{ text-align: right; font-weight: 600; }}
.finops-trend-card {{
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md);
    box-shadow: var(--cp-shadow); padding: 20px 22px 8px; margin-bottom: 18px;
}}
.finops-trend-head {{ display: flex; align-items: baseline; justify-content: space-between; flex-wrap: wrap; gap: 8px; margin-bottom: 4px; }}
.finops-trend-total {{ font-size: 30px; font-weight: 700; letter-spacing: -0.02em; color: var(--cp-accent); }}
.finops-trend-total-label {{ color: var(--cp-text-muted); font-size: 12px; margin-left: 8px; }}
.finops-trend-delta {{ font-size: 13px; font-weight: 600; }}
.finops-trend-svg {{ width: 100%; height: auto; aspect-ratio: 1000 / 260; display: block; }}
.trend-grid-line {{ stroke: var(--cp-border); stroke-width: 1; stroke-dasharray: 4 4; }}
.trend-dot {{ fill: var(--cp-accent); stroke: var(--cp-surface); stroke-width: 2; }}
.trend-value-label {{ font-size: 15px; fill: var(--cp-text); font-weight: 600; }}
.trend-month-label {{ font-size: 15px; fill: var(--cp-text-muted); }}
.cost-bar-chart {{ display: flex; flex-direction: column; gap: 10px; }}
.cost-bar-row {{ display: grid; grid-template-columns: minmax(90px, 0.9fr) 2fr auto 34px; align-items: center; gap: 10px; font-size: 12px; }}
.cost-bar-label {{ color: var(--cp-text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.cost-bar-track {{ height: 8px; border-radius: 980px; background: var(--cp-surface-soft); overflow: hidden; }}
.cost-bar-fill {{ height: 100%; border-radius: 980px; background: linear-gradient(90deg, var(--cp-accent), var(--cp-success)); }}
.inline-bar-wrap {{ display: flex; align-items: center; justify-content: flex-end; gap: 8px; }}
.inline-bar-track {{ width: 60px; height: 6px; border-radius: 980px; background: var(--cp-border); overflow: hidden; }}
.inline-bar-fill {{ height: 100%; border-radius: 980px; }}
.cost-bar-value {{ font-weight: 700; white-space: nowrap; }}
.cost-bar-pct {{ color: var(--cp-text-muted); text-align: right; }}

/* Chart.js integration - chart cards + chart/table toggle pattern */
.chart-card {{ background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md); padding: 16px 18px; box-shadow: var(--cp-shadow); }}
.chart-card-body {{ position: relative; width: 100%; }}
.chart-summary-row {{ display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }}
.chart-summary-row .chart-card {{ min-width: 260px; }}
.chart-table-block {{ display: flex; flex-direction: column; gap: 10px; }}
.chart-table-toggle {{ display: inline-flex; gap: 4px; background: var(--cp-surface-soft); border-radius: 980px; padding: 3px; align-self: flex-start; }}
.ctt-btn {{ border: none; background: transparent; color: var(--cp-text-muted); font-size: 11px; font-weight: 600; padding: 5px 14px; border-radius: 980px; cursor: pointer; transition: background 0.15s var(--cp-ease), color 0.15s var(--cp-ease); }}
.ctt-btn.is-active {{ background: var(--cp-accent); color: var(--cp-accent-fg); }}
.ctt-panel[hidden] {{ display: none; }}
.chart-view-details-link {{ font-size: 11px; color: var(--cp-link); cursor: pointer; background: none; border: none; padding: 0; text-decoration: underline; align-self: flex-start; }}
.chart-mini-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; }}
.chart-mini-title {{ font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; color: var(--cp-text-muted); margin-bottom: 6px; }}
.control-progress-list {{ display: flex; flex-direction: column; gap: 10px; max-height: 320px; overflow-y: auto; }}
.control-progress-row {{ display: grid; grid-template-columns: minmax(120px, 1.2fr) 2fr auto; align-items: center; gap: 10px; font-size: 12px; }}
.control-progress-label {{ color: var(--cp-text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.control-progress-track {{ height: 8px; border-radius: 980px; background: var(--cp-surface-soft); overflow: hidden; }}
.control-progress-fill {{ height: 100%; border-radius: 980px; }}
.control-progress-value {{ font-weight: 700; white-space: nowrap; color: var(--cp-text-muted); }}
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
    background: var(--cp-surface); border-radius: var(--cp-radius-lg); padding: 32px;
    border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow);
}}
.overall-score {{
    display: flex; flex-direction: column; align-items: center; justify-content: center;
}}
.score-circle {{
    width: 140px; height: 140px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 48px; font-weight: 700;
    background: conic-gradient(
        var(--overall-color) calc(var(--overall-deg) * 1deg),
        var(--cp-border) calc(var(--overall-deg) * 1deg)
    );
    position: relative;
}}
.score-circle::after {{
    content: attr(data-value); position: absolute;
    width: 114px; height: 114px; border-radius: 50%;
    background: var(--cp-surface); display: flex; align-items: center; justify-content: center;
    font-size: 40px; font-weight: 700; letter-spacing: -0.02em;
    color: var(--overall-color);
}}
.score-label {{ margin-top: 12px; font-size: 13px; color: var(--cp-text-muted); font-weight: 590; }}

.pillars-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;
}}
.pillar-card {{
    background: var(--cp-surface-soft); border-radius: var(--cp-radius-md); padding: 18px;
    border: 1px solid var(--cp-border); transition: transform 0.2s var(--cp-ease), box-shadow 0.2s var(--cp-ease);
}}
.pillar-card:hover {{ transform: translateY(-2px); box-shadow: var(--cp-shadow); }}
.pillar-score {{ font-size: 32px; font-weight: 700; letter-spacing: -0.02em; }}
.pillar-bar {{ height: 5px; background: var(--cp-border); border-radius: 980px; margin: 10px 0; overflow: hidden; }}
.pillar-bar-fill {{ height: 100%; border-radius: 980px; transition: width 0.6s var(--cp-ease); }}
.pillar-name {{ font-size: 12px; font-weight: 590; color: var(--cp-text-muted); }}

@media (max-width: 700px) {{
    .dashboard {{ padding: 16px; }}
    .header {{ padding: 72px 20px 24px; }}
    .header h1 {{ padding-right: 0; font-size: 24px; }}
    .header-title-sub {{ font-size: 18px; }}
    .header-logo {{ width: 40px; height: 40px; }}
    .header-brand .az-icon.header-logo {{ width: 40px; height: 40px; }}
    .header-controls {{ top: 16px; right: 16px; }}
    .score-section {{ grid-template-columns: minmax(0, 1fr); padding: 20px; }}
    .pillars-grid {{ grid-template-columns: repeat(auto-fit, minmax(min(200px, 100%), 1fr)); min-width: 0; }}
    .pillar-card {{ min-width: 0; }}
}}

/* Score Breakdown */
.methodology-note {{
    background: var(--cp-accent-soft); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md);
    padding: 14px 18px; font-size: 12px; color: var(--cp-text); margin-bottom: 16px;
}}
.methodology-note a {{ color: var(--cp-link); }}
.breakdown-grid {{
    display: grid; grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 16px;
}}
@media (min-width: 701px) {{ .breakdown-grid {{ grid-auto-rows: 1fr; }} }}
@media (max-width: 1100px) {{ .breakdown-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}
@media (max-width: 700px) {{ .breakdown-grid {{ grid-template-columns: 1fr; }} }}
.breakdown-card {{
    background: var(--cp-surface); border-radius: var(--cp-radius-md); padding: 18px;
    border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow); min-width: 0;
    transition: transform 0.2s var(--cp-ease);
}}
.breakdown-card:hover {{ transform: translateY(-2px); }}
.breakdown-card h4 {{ font-size: 13px; font-weight: 590; margin-bottom: 10px; display: flex; justify-content: space-between; }}
.breakdown-score {{ color: var(--cp-link); }}
.breakdown-note {{ font-size: 12px; color: var(--cp-text-muted); }}
.pillar-note {{ margin-top: 8px; font-size: 10.5px; line-height: 1.4; color: var(--cp-text-muted); }}
.secure-gap-note {{ margin: 6px 0; padding: 8px 10px; border-left: 3px solid var(--cp-warning); background: var(--cp-surface-soft); border-radius: 6px; }}
.breakdown-table-wrap {{ width: 100%; overflow-x: auto; }}
.cost-breakdown-card table {{ table-layout: fixed; }}
.cost-breakdown-card th,
.cost-breakdown-card td {{ overflow-wrap: anywhere; white-space: normal; }}
.cost-source-col {{ width: 34%; }}
.cost-findings-col {{ width: 26%; }}
.subcategory-table {{ table-layout: fixed; }}
.subcategory-table td {{ overflow-wrap: break-word; white-space: normal; padding: 8px 6px; }}
.subcategory-table th {{ overflow-wrap: normal; word-break: normal; white-space: normal; padding: 8px 6px; font-size: 10px; }}
.subcat-name-col {{ width: 36%; }}
.subcat-num-col {{ width: 15%; }}
.subcat-score-col {{ width: 34%; }}
.inline-bar-wrap {{ min-width: 0; }}
.inline-bar-track {{ flex: 0 0 auto; }}
.subcategory-table .inline-bar-wrap {{ gap: 4px; }}
.subcategory-table .inline-bar-track {{ width: 30px; }}
.subcategory-table .inline-bar-wrap span {{ font-size: 10px; white-space: nowrap; }}

/* Governance view (native hierarchy tree, scope insights) */
.resource-rec-intro {{ color: var(--cp-text-muted); font-size: 12px; margin: -6px 0 16px; max-width: 900px; }}
.resource-rec-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 14px; }}
.resource-rec-card {{
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-top: 3px solid var(--cp-link);
    border-radius: var(--cp-radius-md); padding: 18px; min-width: 0; box-shadow: var(--cp-shadow);
}}
.resource-rec-card.resource-rec-high {{ border-top-color: var(--cp-danger); }}
.resource-rec-card.resource-rec-medium {{ border-top-color: var(--cp-warning); }}
.resource-rec-card.resource-rec-low {{ border-top-color: var(--cp-link); }}
.resource-rec-heading {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 10px; }}
.resource-rec-heading h3 {{ font-size: 15px; font-weight: 650; margin-top: 3px; }}
.resource-rec-priority {{ color: var(--cp-text-muted); font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 0; }}
.resource-rec-high .resource-rec-priority {{ color: var(--cp-danger); }}
.resource-rec-medium .resource-rec-priority {{ color: var(--cp-warning); }}
.resource-rec-count {{
    background: var(--cp-surface-soft); color: var(--cp-text-muted); border-radius: 6px;
    padding: 4px 7px; font-size: 10px; font-weight: 650; white-space: nowrap;
}}
.resource-rec-card > p {{ color: var(--cp-text-muted); font-size: 12px; margin-bottom: 14px; }}
.resource-rec-implementation {{
    display: grid; gap: 2px; margin-bottom: 13px; padding: 9px 10px;
    background: var(--cp-surface-soft); border-radius: 7px;
}}
.resource-rec-implementation span {{ color: var(--cp-text-muted); font-size: 9px; font-weight: 700; text-transform: uppercase; }}
.resource-rec-implementation strong {{ color: var(--cp-text); font-size: 12px; }}
.resource-rec-implementation small {{ color: var(--cp-text-soft); font-size: 9px; }}
.resource-rec-evidence-label {{ font-size: 10px; font-weight: 700; margin-bottom: 5px; }}
.resource-rec-card ul {{ padding-left: 17px; margin-bottom: 14px; }}
.resource-rec-card li {{ color: var(--cp-text-muted); font-size: 11px; margin: 4px 0; }}
.resource-rec-card li::marker {{ color: var(--cp-accent); }}
.resource-rec-card .resource-rec-more {{ color: var(--cp-text-soft); }}
.resource-rec-link {{ color: var(--cp-link); font-size: 11px; font-weight: 650; text-decoration: none; }}
.resource-rec-link:hover {{ text-decoration: underline; }}
.mg-tree, .mg-tree ul {{ list-style: none; padding-left: 30px; }}
.mg-tree {{ padding-left: 0; }}
.mg-tree li {{ margin: 8px 0; }}
.mg-node-details > summary {{ list-style: none; cursor: pointer; }}
.mg-node-details > summary::-webkit-details-marker {{ display: none; }}
.mg-node {{
    display: inline-flex; align-items: center; gap: 12px; background: var(--cp-surface);
    border: 1px solid var(--cp-border); border-radius: 980px; padding: 10px 18px; font-size: 15px; font-weight: 590;
}}
.mg-name {{ white-space: nowrap; display: inline-flex; align-items: center; gap: 8px; }}
.az-icon {{ width: 22px; height: 22px; flex-shrink: 0; vertical-align: middle; }}
.mg-badge .az-icon {{ width: 14px; height: 14px; }}
.mg-sub-leaf .az-icon {{ width: 18px; height: 18px; }}
.mg-badge {{
    display: inline-flex; align-items: center; justify-content: center; gap: 4px; min-width: 26px; height: 26px;
    border-radius: 980px; font-size: 12px; font-weight: 700; padding: 0 10px;
}}
.mg-badge-policy {{ background: rgba(0, 120, 212, 0.14); color: var(--cp-link); }}
.mg-badge-rbac {{ background: rgba(255, 149, 0, 0.14); color: var(--cp-warning); }}
.mg-badge-subs {{ background: rgba(48, 209, 88, 0.12); color: var(--cp-success); }}
.mg-sub-leaf {{ font-size: 14px; color: var(--cp-text-muted); padding: 6px 0; list-style: none; display: flex; align-items: center; gap: 8px; }}
.mg-diagram-wrap {{ overflow-x: auto; padding: 8px 0 18px; text-align: center; }}
.mg-diagram-svg {{ display: inline-block; }}
.mg-diagram-link {{ stroke: var(--cp-border); stroke-width: 2; }}
.mg-diagram-node {{
    box-sizing: border-box; width: 100%; height: 100%; background: var(--cp-surface); border: 1px solid var(--cp-border);
    border-top-width: 3px; border-radius: var(--cp-radius-md); box-shadow: var(--cp-shadow); overflow: hidden;
    display: flex; flex-direction: column; justify-content: center; align-items: center; gap: 4px; padding: 8px 10px;
    font-family: inherit;
}}
.mg-diagram-node-mg {{ border-top-color: var(--cp-link); }}
.mg-diagram-node-sub {{ border-top-color: var(--cp-success); }}
.mg-diagram-node-rg {{ border-top-color: var(--cp-warning); }}
.mg-diagram-node-resource {{ border-top-color: var(--cp-border); }}
.mg-diagram-node-more {{
    border-style: dashed; border-top-width: 1px; color: var(--cp-text-muted); font-size: 12px; font-weight: 590;
    box-shadow: none; background: transparent;
}}
.mg-diagram-name {{
    display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 600; width: 100%; flex: 0 0 auto;
    overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
}}
.mg-diagram-name .az-icon {{ width: 16px; height: 16px; flex-shrink: 0; }}
.mg-diagram-emoji {{ font-size: 13px; line-height: 1; flex-shrink: 0; }}
.mg-diagram-badges {{ display: flex; gap: 6px; flex: 0 0 auto; }}
.mg-diagram-restype {{
    font-size: 10px; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: var(--cp-text-muted);
}}
.mg-diagram-expand {{
    all: unset; cursor: pointer; font-size: 11px; font-weight: 600; color: var(--cp-link);
    display: flex; align-items: center; gap: 3px; flex: 0 0 auto;
}}
.mg-diagram-expand-icon {{ display: inline-block; font-size: 9px; transition: transform 0.15s var(--cp-ease); }}
.mg-diagram-expand[aria-expanded="true"] .mg-diagram-expand-icon {{ transform: rotate(90deg); }}
.scope-insights {{ display: flex; flex-direction: column; gap: 6px; }}
.scope-detail {{
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md); padding: 10px 14px;
}}
.scope-detail summary {{ cursor: pointer; font-size: 12px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
.scope-detail .scope-id {{ color: var(--cp-text-muted); font-weight: 400; }}
.scope-detail p {{ font-size: 12px; margin: 8px 0; }}
.rg-explorer-wrap {{ display: flex; flex-direction: column; gap: 8px; }}
.rg-explorer {{
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md); padding: 10px 14px;
}}
.rg-explorer summary {{ cursor: pointer; font-size: 13px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
.rg-explorer summary .az-icon {{ width: 18px; height: 18px; }}
.rg-explorer .scope-id {{ color: var(--cp-text-muted); font-weight: 400; }}
.rg-explorer-list {{
    margin-top: 10px; padding-left: 16px; border-left: 2px solid var(--cp-border);
    display: flex; flex-direction: column; gap: 8px;
}}
.rg-explorer-item {{ border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md); padding: 8px 12px; }}
.rg-explorer-item summary {{ cursor: pointer; font-size: 12.5px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
.rgx-resource-list {{
    list-style: none; margin: 8px 0 0; padding-left: 16px; border-left: 2px dashed var(--cp-border);
    display: flex; flex-direction: column; gap: 5px;
}}
.rgx-resource-list li {{ font-size: 12px; display: flex; align-items: center; gap: 8px; }}
.rgx-resource-list .resource-name {{ flex: 1 1 auto; }}
.rgx-resource-list .resource-icon {{ width: 16px; height: 16px; flex: 0 0 auto; }}
.rgx-resource-list .resource-more {{ color: var(--cp-text-muted); font-style: italic; }}
.mini-table {{ width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 6px; }}
.mini-table th, .mini-table td {{ border: 1px solid var(--cp-border); padding: 4px 8px; text-align: left; }}
.table-more-note {{
    display: block; width: 100%; text-align: left; font-size: 11px; color: var(--cp-accent); font-weight: 600;
    margin-top: 6px; padding: 4px 0; background: none; border: none; cursor: pointer; font-family: inherit;
}}
.table-more-note:hover {{ text-decoration: underline; }}
.table-more-note .less-label {{ display: none; }}
.table-more-note.expanded .more-label {{ display: none; }}
.table-more-note.expanded .less-label {{ display: inline; }}
.table-truncated-note {{ font-size: 11px; color: var(--cp-text-muted); font-style: italic; margin-top: 6px; }}
@media (max-width: 700px) {{
    .resource-rec-grid {{ grid-template-columns: minmax(0, 1fr); }}
    .resource-rec-heading {{ align-items: flex-start; }}
}}

/* Score continuity ("since last scan") */
.progress-banner {{
    background: var(--cp-surface); border-radius: var(--cp-radius-md); padding: 16px 20px; margin-bottom: 20px;
    border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow);
}}
.progress-banner-empty {{ display: flex; align-items: center; gap: 10px; color: var(--cp-text-muted); font-size: 13px; }}
.progress-banner-row {{ display: flex; align-items: center; flex-wrap: wrap; gap: 10px 16px; font-size: 14px; }}
.progress-title {{ font-weight: 600; }}
.progress-overall {{ color: var(--cp-text-muted); }}
.progress-delta {{ font-weight: 700; padding: 2px 8px; border-radius: 100px; font-size: 13px; }}
.progress-delta-up {{ color: var(--cp-success); background: rgba(29, 154, 74, 0.12); }}
.progress-delta-down {{ color: var(--cp-danger); background: rgba(255, 59, 48, 0.12); }}
.progress-delta-neutral {{ color: var(--cp-text-muted); background: var(--cp-surface-soft); }}
.progress-new-count {{ font-size: 12px; color: var(--cp-warning); font-weight: 600; }}
.progress-pillar-row {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }}
.progress-pillar-chip {{
    display: inline-flex; align-items: center; gap: 6px; font-size: 12px; padding: 4px 10px;
    border-radius: 100px; background: var(--cp-surface-soft); border: 1px solid var(--cp-border);
}}
.progress-resolved {{ margin-top: 12px; }}
.progress-resolved summary {{ cursor: pointer; font-size: 13px; font-weight: 600; color: var(--cp-success); }}
.progress-resolved-list {{ margin: 10px 0 0; padding-left: 4px; list-style: none; font-size: 13px; }}
.progress-resolved-list li {{ padding: 6px 0; border-top: 1px solid var(--cp-border); display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
.progress-resolved-list li:first-child {{ border-top: none; }}

/* Stats */
.stats-row {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px; margin-bottom: 24px;
}}
.stat-box {{
    background: var(--cp-surface); border-radius: var(--cp-radius-md); padding: 24px;
    border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow); text-align: center;
    transition: transform 0.2s var(--cp-ease);
}}
.stat-box:hover {{ transform: translateY(-2px); }}
.stat-box .value {{ font-size: 36px; font-weight: 700; letter-spacing: -0.02em; color: var(--cp-link); }}
.stat-box .label {{ font-size: 12px; color: var(--cp-text-muted); margin-top: 4px; }}

/* Action Items */
.section {{ margin-bottom: 24px; }}
.section-title {{
    font-size: 19px; font-weight: 600; letter-spacing: -0.01em; margin-bottom: 16px;
    padding-bottom: 10px; border-bottom: 2px solid var(--cp-accent);
}}
.action-item {{
    background: var(--cp-surface); border-radius: var(--cp-radius-md); padding: 18px; margin-bottom: 12px;
    border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow);
    transition: transform 0.2s var(--cp-ease);
}}
.action-item:hover {{ transform: translateY(-1px); }}
.action-header {{ display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }}
.action-num {{ font-weight: 700; color: var(--cp-text-soft); font-size: 12px; }}
.severity-badge {{
    color: var(--cp-accent-fg); padding: 3px 10px; border-radius: 980px;
    font-size: 10px; font-weight: 700; text-transform: uppercase;
}}
.pillar-tag {{
    background: var(--cp-surface-soft); padding: 3px 10px; border-radius: 980px;
    font-size: 10px; color: var(--cp-text-muted);
}}
.action-title {{ font-weight: 590; font-size: 14px; }}
.action-desc {{ color: var(--cp-text-muted); margin-top: 8px; font-size: 13px; }}
.resource-list {{
    margin-top: 8px; padding-left: 20px;
    font-size: 11px; color: var(--cp-text-muted); max-height: 120px; overflow-y: auto;
}}
.resource-list li {{ margin: 2px 0; }}
.resource-list li[data-resource-id] {{ cursor: pointer; }}
.resource-list li[data-resource-id]:hover {{ color: var(--cp-link); text-decoration: underline; }}

/* Action item toolbar / add form / dismiss */
.section-title-row {{ display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; margin-bottom: 16px; padding-bottom: 10px; border-bottom: 2px solid var(--cp-accent); }}
.action-toolbar {{ display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }}
.btn-add-action {{
    background: var(--cp-accent); color: var(--cp-accent-fg); border: 0; border-radius: 980px;
    padding: 7px 16px; font-size: 12.5px; font-weight: 600; cursor: pointer; transition: opacity 0.15s var(--cp-ease);
}}
.btn-add-action:hover {{ opacity: 0.85; }}
.btn-link {{
    background: transparent; border: 0; color: var(--cp-link); font-size: 12.5px; font-weight: 590;
    cursor: pointer; padding: 4px 2px;
}}
.btn-link:hover {{ text-decoration: underline; }}
.action-check-wrap {{ display: inline-flex; align-items: center; cursor: pointer; }}
.action-check {{ width: 16px; height: 16px; cursor: pointer; accent-color: var(--cp-accent); }}
.action-item.is-dismissed {{ opacity: 0.5; }}
.action-item.is-dismissed .action-title {{ text-decoration: line-through; }}
.action-item[data-custom="true"] {{ border-style: dashed; }}
.action-form {{
    background: var(--cp-surface-soft); border: 1px solid var(--cp-border); border-radius: var(--cp-radius-md);
    padding: 16px; margin-bottom: 16px; display: none; grid-template-columns: 1fr 1fr; gap: 12px;
}}
.action-form.is-open {{ display: grid; }}
.action-form-field {{ display: flex; flex-direction: column; gap: 4px; }}
.action-form-field.full {{ grid-column: 1 / -1; }}
.action-form-field label {{ font-size: 11px; font-weight: 600; color: var(--cp-text-muted); text-transform: uppercase; }}
.action-form-field input, .action-form-field select, .action-form-field textarea {{
    background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: 8px;
    padding: 8px 10px; font: inherit; font-size: 13px; color: var(--cp-text); resize: vertical;
}}
.action-form-actions {{ grid-column: 1 / -1; display: flex; gap: 10px; justify-content: flex-end; }}
.action-form-actions button {{ border-radius: 980px; padding: 7px 18px; font-size: 12.5px; font-weight: 600; cursor: pointer; border: 0; }}
.severity-filter-bar {{ display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 16px; }}
.severity-filter-label {{ font-size: 12px; font-weight: 600; color: var(--cp-text-muted); margin-right: 2px; }}
.severity-chip {{
    --chip-color: var(--cp-text-muted);
    background: var(--cp-surface); border: 1px solid var(--cp-border); color: var(--cp-text);
    border-radius: 980px; padding: 5px 14px; font-size: 12px; font-weight: 600; cursor: pointer;
    transition: all 0.15s var(--cp-ease); position: relative;
}}
.severity-chip::before {{
    content: ''; display: inline-block; width: 8px; height: 8px; border-radius: 50%;
    background: var(--chip-color); margin-right: 6px; vertical-align: middle;
}}
.severity-chip[data-severity="all"]::before {{ display: none; }}
.severity-chip:hover {{ border-color: var(--chip-color); }}
.severity-chip.is-active {{ background: var(--chip-color); border-color: var(--chip-color); color: var(--cp-accent-fg); }}
.severity-chip[data-severity="all"].is-active {{ background: var(--cp-accent); border-color: var(--cp-accent); }}
.action-item.is-severity-hidden {{ display: none !important; }}
.action-form-save {{ background: var(--cp-accent); color: var(--cp-accent-fg); }}
.action-form-cancel {{ background: var(--cp-surface); border: 1px solid var(--cp-border); color: var(--cp-text); }}

/* Tables */
.data-grid {{
    display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px;
}}
@media (max-width: 900px) {{ .data-grid {{ grid-template-columns: 1fr; }} }}
.table-card {{
    background: var(--cp-surface); border-radius: var(--cp-radius-md); padding: 22px;
    border: 1px solid var(--cp-border); box-shadow: var(--cp-shadow); min-width: 0;
}}
.table-card h3 {{ font-size: 14px; font-weight: 590; margin-bottom: 12px; color: var(--cp-text); }}
.table-card table {{ table-layout: fixed; }}
.table-card th,
.table-card td {{ overflow-wrap: anywhere; white-space: normal; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
th {{ background: transparent; padding: 8px; text-align: left; font-weight: 590; color: var(--cp-text-muted); border-bottom: 1px solid var(--cp-border-strong); }}
td {{ padding: 8px; border-bottom: 1px solid var(--cp-border); }}
tr:hover td {{ background: var(--cp-accent-soft); }}
.num {{ text-align: right; font-weight: 590; }}

/* Footer */
.footer {{
    text-align: center; padding: 32px 24px; color: var(--cp-text-soft); font-size: 11px;
    border-top: 1px solid var(--cp-border); margin-top: 32px;
}}

/* Header controls and trilingual toggle (default: English shown, Spanish/Portuguese hidden) */
.i18n-es, .i18n-pt {{ display: none; }}
body.lang-es .i18n-en, body.lang-es .i18n-pt {{ display: none; }}
body.lang-es .i18n-es {{ display: inline; }}
body.lang-pt .i18n-en, body.lang-pt .i18n-es {{ display: none; }}
body.lang-pt .i18n-pt {{ display: inline; }}
.header-controls {{
    position: absolute; top: 24px; right: 28px;
    display: flex; gap: 8px;
}}
.header-control {{
    background: rgba(255, 255, 255, 0.95); color: #1d1d1f;
    border: 1px solid rgba(255, 255, 255, 0.5); border-radius: 980px;
    height: 34px; min-width: 38px; padding: 0 14px;
    font-size: 12px; font-weight: 590; cursor: pointer;
    box-shadow: var(--cp-shadow); transition: transform 0.2s var(--cp-ease), background 0.2s var(--cp-ease);
}}
.header-control:hover {{ transform: translateY(-1px); background: #ffffff; }}
.theme-toggle {{ font-size: 18px; line-height: 1; padding: 0; }}
</style>
</head>
<body>
<div class="dashboard">

<!-- Header -->
<div class="header">
    <div class="header-controls">
        <button id="themeToggleBtn" class="header-control theme-toggle" onclick="toggleTheme()" type="button" aria-label="Switch to night mode" title="Switch to night mode">☾</button>
        <button id="langToggleBtn" class="header-control lang-toggle" onclick="toggleLang()" type="button">ES 🇪🇸</button>
    </div>
    <div class="header-brand">
        {self._azure_icon("azure-logo", "az-icon header-logo")}
        <h1>{bi("Azure Well-Architected Framework", "Azure Well-Architected Framework", "Azure Well-Architected Framework")}<span class="header-title-sub">{bi("Discovery Report", "Informe de Descubrimiento", "Relatório de Descoberta")}</span></h1>
    </div>
    <div class="subtitle">{bi("Generated", "Generado", "Gerado")}: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {bi("Consolidated view across all discovery phases", "Vista consolidada de todas las fases de descubrimiento", "Visão consolidada de todas as fases de descoberta")}</div>
</div>

<div class="view-tabs" id="viewTabsBar" role="tablist" aria-label="Dashboard views">
    <span id="viewTabThumb" class="view-tab-thumb" aria-hidden="true"></span>
    <button id="overviewTab" class="view-tab" role="tab" aria-selected="true" aria-controls="overviewView" onclick="setDashboardView('overview')" type="button">{bi('Overview', 'Resumen', 'Visão geral')}</button>
    <button id="reliabilityTab" class="view-tab" role="tab" aria-selected="false" aria-controls="reliabilityView" onclick="setDashboardView('reliability')" type="button">{bi('Reliability', 'Confiabilidad', 'Confiabilidade')}</button>
    <button id="securityTab" class="view-tab" role="tab" aria-selected="false" aria-controls="securityView" onclick="setDashboardView('security')" type="button">{bi('Security', 'Seguridad', 'Segurança')}</button>
    <button id="costTab" class="view-tab" role="tab" aria-selected="false" aria-controls="costView" onclick="setDashboardView('cost')" type="button">{bi('Cost Opt', 'Optim. Costos', 'Otim. Custos')}</button>
    <button id="operationalTab" class="view-tab" role="tab" aria-selected="false" aria-controls="operationalView" onclick="setDashboardView('operational')" type="button">{bi('OpEx', 'Exc. Operativa', 'Exc. Operacional')}</button>
    <button id="performanceTab" class="view-tab" role="tab" aria-selected="false" aria-controls="performanceView" onclick="setDashboardView('performance')" type="button">{bi('Performance', 'Rendimiento', 'Desempenho')}</button>
    <button id="governanceTab" class="view-tab" role="tab" aria-selected="false" aria-controls="governanceView" onclick="setDashboardView('governance')" type="button">{bi('Governance', 'Gobernanza', 'Governança')}</button>
</div>

{global_sub_filter_html}

<main id="overviewView" class="dashboard-view" role="tabpanel" aria-labelledby="overviewTab">
<!-- Score continuity since last scan -->
{progress_banner_html}

<!-- Overall Score + Pillar Breakdown -->
<div class="score-section">
    <div class="overall-score">
        <div class="score-circle" id="overallScoreCircle" data-value="{overall}" style="--overall-deg:{overall * 3.6}; --overall-color:{overall_color};"></div>
        <div class="score-label">{bi("Overall WAF Score", "Puntaje General WAF", "Pontuação Geral WAF")}</div>
    </div>
    <div class="pillars-grid">
        {pillar_cards}
    </div>
</div>

<!-- WAF Pillar Scores comparison chart -->
<div class="section chart-card">
    <div class="section-title">{bi("WAF Pillar Scores", "Puntajes por Pilar WAF", "Pontuações por Pilar WAF")}{it("Side-by-side comparison of all pillar scores.", "Comparaci\u00f3n lado a lado de los puntajes de todos los pilares.", "Comparação lado a lado das pontuações de todos os pilares.")}</div>
    <div class="chart-card-body" style="height:220px;"><canvas id="pillarScoresChart"></canvas></div>
</div>

<!-- Stats Row -->
<div class="stats-row">
    <div class="stat-box"><div class="value">{total_subs}</div><div class="label">{bi("Subscriptions", "Suscripciones", "Assinaturas")}</div></div>
    <div class="stat-box"><div class="value">{total_resources:,}</div><div class="label">{bi("Total Resources", "Recursos Totales", "Total de Recursos")}</div></div>
    <div class="stat-box"><div class="value" id="totalActionsStat">{total_actions}</div><div class="label">{bi("Action Items", "Elementos de Acci\u00f3n", "Itens de Ação")}</div></div>
    <div class="stat-box"><div class="value" id="criticalActionsStat" style="color:var(--cp-danger);">{critical_actions}</div><div class="label">{bi("Critical/High Priority", "Prioridad Cr\u00edtica/Alta", "Prioridade Crítica/Alta")}</div></div>
</div>

<!-- Score Methodology -->
<div class="section">
    <div class="section-title">{bi("Score Methodology (Microsoft Advisor Score model)", "Metodolog\u00eda de Puntuaci\u00f3n (modelo de Azure Advisor Score)", "Metodologia de Pontuação (modelo do Azure Advisor Score)")}{it("Explains how each pillar score (0-100) is calculated from Advisor recommendation subcategories and weights.", "Explica c\u00f3mo se calcula el puntaje de cada pilar (0-100) a partir de las subcategor\u00edas y pesos de las recomendaciones de Advisor.", "Explica como a pontuação de cada pilar (0-100) é calculada a partir das subcategorias e pesos das recomendações do Advisor.")}</div>
    <div class="methodology-note">
        {bi("Pillar scores follow the official", "Los puntajes de cada pilar siguen la f\u00f3rmula oficial de", "As pontuações de cada pilar seguem a fórmula oficial do")} <a href="https://learn.microsoft.com/en-us/azure/advisor/advisor-score#calculation-of-advisor-score" target="_blank">Azure Advisor Score</a>
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
            "recursos saludables en lugar de ponderaci\u00f3n por costo de lista, ya que esta herramienta no llama a la API de precios de Azure.",
            "e os pesos de subcategoria em vez de um sistema de pontos arbitrário. Segurança usa diretamente o Secure Score do Microsoft Defender. "
            "Confiabilidade, Desempenho e Excelência Operacional usam os pesos de subcategoria publicados pela Microsoft — as recomendações são "
            "mapeadas para subcategorias por correspondência de palavras-chave (o Azure Resource Graph não expõe a tag interna de subcategoria do Advisor), "
            "e o inventário de recursos descobertos é usado como o conjunto compartilhado de 'total de recursos aplicáveis'. Custo usa uma proporção de "
            "recursos saudáveis em vez de ponderação por custo de varejo, já que esta ferramenta não chama a API de Preços de Varejo do Azure."
        )}
    </div>
    <div class="breakdown-grid">
        {score_breakdown_html}
    </div>
</div>

<!-- Evidence-based remediation recommendations -->
<div class="section resource-recommendations">
    <div class="section-title security-title">{bi('Evidence-based remediation priorities', 'Prioridades de remediación basadas en evidencia', 'Prioridades de remediação baseadas em evidências')}{it('Highest-priority observed findings, balanced across WAF pillars. A service is suggested only when the finding directly identifies a relevant implementation path.', 'Hallazgos observados de mayor prioridad, balanceados entre pilares WAF. Solo se sugiere un servicio cuando el hallazgo identifica directamente una ruta de implementación relevante.', 'Descobertas observadas de maior prioridade, equilibradas entre os pilares WAF. Um serviço só é sugerido quando a descoberta identifica diretamente um caminho de implementação relevante.')}</div>
    <p class="resource-rec-intro">{bi('Each card starts from an observed action item and preserves its source, affected scope, and WAF pillar. Suggested services are implementation options, not automatic deployment decisions; architecture, licensing, cost, and tradeoffs still require validation.', 'Cada tarjeta parte de una acción observada y conserva su fuente, alcance afectado y pilar WAF. Los servicios sugeridos son opciones de implementación, no decisiones automáticas de despliegue; aún deben validarse arquitectura, licenciamiento, costo y compromisos.', 'Cada cartão parte de uma ação observada e preserva sua fonte, escopo afetado e pilar WAF. Os serviços sugeridos são opções de implementação, não decisões automáticas de implantação; arquitetura, licenciamento, custo e compensações ainda exigem validação.')}</p>
    {self._render_azure_resource_recommendations()}
</div>

<!-- Action Items -->
<div class="section">
    <div class="chart-summary-row">
        <div class="chart-card" style="flex:1;">
            <div class="section-title" style="font-size:14px;">{bi('Recommendations by Severity', 'Recomendaciones por Severidad', 'Recomendações por Severidade')}</div>
            <div class="chart-card-body" style="height:180px;">
                {f'<canvas id="actionSeverityChart"></canvas>' if severity_present else f'<p class="empty-state">{bi_empty("No action items found.", "No se encontraron elementos de acci\u00f3n.")}</p>'}
            </div>
        </div>
        <div class="chart-card" style="flex:1;">
            <div class="section-title" style="font-size:14px;">{bi('Recommendations by WAF Pillar', 'Recomendaciones por Pilar WAF', 'Recomendações por Pilar WAF')}</div>
            <div class="chart-card-body" style="height:180px;">
                {f'<canvas id="actionByPillarChart"></canvas>' if pillars_with_actions else f'<p class="empty-state">{bi_empty("No action items found.", "No se encontraron elementos de acci\u00f3n.")}</p>'}
            </div>
        </div>
    </div>
    <div class="section-title-row">
        <div class="section-title" style="border-bottom:none; padding-bottom:0; margin-bottom:0;">{bi('Action Items', 'Elementos de Acci\u00f3n', 'Itens de Ação')} (<span id="actionItemsCount">{total_actions}</span>){it('Consolidated, prioritized list of action items across all pillars - add, resolve, or dismiss items directly from here.', 'Lista consolidada y priorizada de elementos de acci\u00f3n de todos los pilares: agrega, resuelve o descarta elementos directamente aqu\u00ed.', 'Lista consolidada e priorizada de itens de ação de todos os pilares - adicione, resolva ou descarte itens diretamente aqui.')}</div>
        <div class="action-toolbar">
            <button id="addActionBtn" class="btn-add-action" type="button">+ {bi('Add action item', 'Agregar elemento', 'Adicionar item')}</button>
            <button id="toggleDismissedBtn" class="btn-link" type="button"
                    data-show-en="Show dismissed" data-show-es="Mostrar descartados" data-show-pt="Mostrar descartados"
                    data-hide-en="Hide dismissed" data-hide-es="Ocultar descartados" data-hide-pt="Ocultar descartados">
                <span class="i18n-en toggle-label">Show dismissed</span><span class="i18n-es toggle-label">Mostrar descartados</span><span class="i18n-pt toggle-label">Mostrar descartados</span> (<span class="dismissed-count">0</span>)
            </button>
        </div>
    </div>
    <div id="severityFilterBar" class="severity-filter-bar">
        <span class="severity-filter-label">{bi('Filter by severity:', 'Filtrar por severidad:', 'Filtrar por severidade:')}</span>
        <button type="button" class="severity-chip is-active" data-severity="all">{bi('All', 'Todos', 'Todos')}</button>
        <button type="button" class="severity-chip" data-severity="critical" style="--chip-color: var(--cp-critical);">{bi('Critical', 'Cr\u00edtico', 'Crítico')}</button>
        <button type="button" class="severity-chip" data-severity="high" style="--chip-color: var(--cp-danger);">{bi('High', 'Alto', 'Alto')}</button>
        <button type="button" class="severity-chip" data-severity="medium" style="--chip-color: var(--cp-warning);">{bi('Medium', 'Medio', 'Médio')}</button>
        <button type="button" class="severity-chip" data-severity="low" style="--chip-color: var(--cp-link);">{bi('Low', 'Bajo', 'Baixo')}</button>
        <button type="button" class="severity-chip" data-severity="info" style="--chip-color: var(--cp-success);">{bi('Info', 'Info', 'Info')}</button>
    </div>
    <div id="addActionForm" class="action-form">
        <div class="action-form-field full">
            <label>{bi('Title', 'T\u00edtulo', 'Título')}</label>
            <input type="text" id="newActionTitle" placeholder="e.g. Rotate storage account access keys">
        </div>
        <div class="action-form-field full">
            <label>{bi('Description', 'Descripci\u00f3n', 'Descrição')}</label>
            <textarea id="newActionDesc" rows="2"></textarea>
        </div>
        <div class="action-form-field">
            <label>{bi('Pillar', 'Pilar', 'Pilar')}</label>
            <select id="newActionPillar">
                <option value="Reliability">Reliability</option>
                <option value="Security">Security</option>
                <option value="Cost Optimization">Cost Optimization</option>
                <option value="Operational Excellence">Operational Excellence</option>
                <option value="Performance Efficiency">Performance Efficiency</option>
                <option value="Governance">Governance</option>
            </select>
        </div>
        <div class="action-form-field">
            <label>{bi('Severity', 'Severidad', 'Severidade')}</label>
            <select id="newActionSeverity">
                <option value="critical">Critical</option>
                <option value="high">High</option>
                <option value="medium" selected>Medium</option>
                <option value="low">Low</option>
                <option value="info">Info</option>
            </select>
        </div>
        <div class="action-form-field full">
            <label>{bi('Resources (comma-separated)', 'Recursos (separados por coma)', 'Recursos (separados por vírgula)')}</label>
            <input type="text" id="newActionResources" placeholder="vm-prod-01, storageacct123">
        </div>
        <div class="action-form-actions">
            <button type="button" id="cancelActionBtn" class="action-form-cancel">{bi('Cancel', 'Cancelar', 'Cancelar')}</button>
            <button type="button" id="saveActionBtn" class="action-form-save">{bi('Save', 'Guardar', 'Salvar')}</button>
        </div>
    </div>
    <div id="actionsListAll" class="action-list" data-pillar="all">
    {actions_html}
    </div>
</div>

<!-- Data Summary Tables -->
<div class="data-grid">
    <div class="table-card">
        <h3>{bi("Top Resource Types", "Principales Tipos de Recursos", "Principais Tipos de Recursos")}</h3>
        {resource_types_chart}
        <table>
            <thead><tr><th>{bi("Resource Type", "Tipo de Recurso", "Tipo de Recurso")}</th><th>{bi("Location", "Ubicaci\u00f3n", "Localização")}</th><th>{bi("Count", "Cantidad", "Quantidade")}</th></tr></thead>
            <tbody>{resource_rows}</tbody>
        </table>
    </div>
    <div class="table-card">
        <h3>{bi("Advisor Summary", "Resumen de Advisor", "Resumo do Advisor")}</h3>
        {advisor_summary_chart}
        <table>
            <thead><tr><th>{bi("Category", "Categor\u00eda", "Categoria")}</th><th>{bi("Impact", "Impacto", "Impacto")}</th><th>{bi("Count", "Cantidad", "Quantidade")}</th></tr></thead>
            <tbody>{advisor_rows}</tbody>
        </table>
    </div>
</div>
</main>

<main id="securityView" class="dashboard-view" role="tabpanel" aria-labelledby="securityTab" hidden>
    {security_view_html}
</main>

<main id="governanceView" class="dashboard-view" role="tabpanel" aria-labelledby="governanceTab" hidden>
    {governance_view_html}
</main>

<main id="costView" class="dashboard-view" role="tabpanel" aria-labelledby="costTab" hidden>
    {cost_view_html}
</main>

<main id="reliabilityView" class="dashboard-view" role="tabpanel" aria-labelledby="reliabilityTab" hidden>
    {reliability_view_html}
</main>

<main id="performanceView" class="dashboard-view" role="tabpanel" aria-labelledby="performanceTab" hidden>
    {performance_view_html}
</main>

<main id="operationalView" class="dashboard-view" role="tabpanel" aria-labelledby="operationalTab" hidden>
    {operational_view_html}
</main>

<!-- Footer -->
<div class="footer">
    {bi("Azure Well-Architected Framework Discovery Report | Generated by Azure Governance Discovery Toolkit", "Informe de Descubrimiento Azure Well-Architected Framework | Generado por Azure Governance Discovery Toolkit", "Relatório de Descoberta Azure Well-Architected Framework | Gerado pelo Azure Governance Discovery Toolkit")}<br>
    {bi("This report consolidates findings from Resource Discovery, Azure Advisor, Metrics Analysis, and Governance Visualization.", "Este informe consolida los hallazgos de Resource Discovery, Azure Advisor, An\u00e1lisis de M\u00e9tricas y Visualizaci\u00f3n de Gobernanza.", "Este relatório consolida os achados do Resource Discovery, Azure Advisor, Análise de Métricas e Visualização de Governança.")}
</div>

</div>
<script>
{chartjs_lib}
</script>
<script>
window.WAF_CHART_DATA = {chart_data_json};
window.WAF_SCORE_MODEL = {score_model_json};
function cssVar(name) {{ return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }}
function resolveColorVars(vars_) {{ return (vars_ || []).map(cssVar); }}
function currentChartLang() {{
    return document.body.classList.contains('lang-es') ? 'es' : (document.body.classList.contains('lang-pt') ? 'pt' : 'en');
}}
function labelsFor(i18nLabels) {{
    if (!i18nLabels) return [];
    return i18nLabels[currentChartLang()] || i18nLabels.en || [];
}}
window.__wafCharts = {{}};
function registerChart(id, chart, meta) {{
    window.__wafCharts[id] = {{ chart: chart, meta: meta || {{}} }};
    return chart;
}}
function destroyChart(id) {{
    var entry = window.__wafCharts[id];
    if (entry && entry.chart) {{ entry.chart.destroy(); }}
    delete window.__wafCharts[id];
}}
function baseChartOptions(extra) {{
    var c = {{ text: cssVar('--cp-text'), muted: cssVar('--cp-text-muted'), border: cssVar('--cp-border'), surface: cssVar('--cp-surface') }};
    var opts = {{
        responsive: true, maintainAspectRatio: false,
        animation: {{ duration: 250 }},
        plugins: {{
            legend: {{ display: false, labels: {{ color: c.text, font: {{ size: 11 }} }} }},
            tooltip: {{
                backgroundColor: c.surface, titleColor: c.text, bodyColor: c.text,
                borderColor: c.border, borderWidth: 1, padding: 10, cornerRadius: 8, displayColors: true
            }}
        }},
        scales: {{}}
    }};
    if (extra) {{
        Object.keys(extra).forEach(function(k) {{
            if (k === 'plugins' && extra.plugins) {{
                Object.keys(extra.plugins).forEach(function(pk) {{
                    opts.plugins[pk] = Object.assign({{}}, opts.plugins[pk], extra.plugins[pk]);
                }});
            }} else {{
                opts[k] = extra[k];
            }}
        }});
    }}
    return opts;
}}
function createHorizontalBarChart(canvasId, i18nLabels, values, colorVars, opts) {{
    var ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    opts = opts || {{}};
    var labels = labelsFor(i18nLabels);
    var c = {{ border: cssVar('--cp-border'), muted: cssVar('--cp-text-muted'), text: cssVar('--cp-text') }};
    var chart = new Chart(ctx, {{
        type: 'bar',
        data: {{ labels: labels, datasets: [{{ data: values, backgroundColor: resolveColorVars(colorVars), borderRadius: 6, maxBarThickness: 22 }}] }},
        options: baseChartOptions({{
            indexAxis: 'y',
            scales: {{
                x: {{
                    beginAtZero: true, max: opts.max, grid: {{ color: c.border }}, ticks: {{
                        color: c.muted,
                        callback: opts.currency ? function(v) {{ return '$' + Number(v).toLocaleString(); }} : undefined
                    }}
                }},
                y: {{ grid: {{ display: false }}, ticks: {{ color: c.text }} }}
            }},
            plugins: {{
                tooltip: {{ callbacks: opts.currency ? {{ label: function(ctx) {{ return '$' + Number(ctx.parsed.x).toLocaleString(); }} }} : undefined }}
            }}
        }})
    }});
    if (opts.onClick) {{
        ctx.onclick = function(evt) {{
            var points = chart.getElementsAtEventForMode(evt, 'nearest', {{ intersect: true }}, true);
            if (points.length) opts.onClick(points[0].index, labels[points[0].index]);
        }};
    }}
    return registerChart(canvasId, chart, {{ type: 'hbar', i18nLabels: i18nLabels, colorVars: colorVars }});
}}
function createDonutChart(canvasId, i18nLabels, values, colorVars, opts) {{
    var ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    opts = opts || {{}};
    var labels = labelsFor(i18nLabels);
    var c = {{ text: cssVar('--cp-text'), surface: cssVar('--cp-surface') }};
    var chart = new Chart(ctx, {{
        type: 'doughnut',
        data: {{ labels: labels, datasets: [{{ data: values, backgroundColor: resolveColorVars(colorVars), borderColor: c.surface, borderWidth: 2 }}] }},
        options: baseChartOptions({{
            cutout: '65%',
            plugins: {{ legend: {{ display: true, position: 'right', labels: {{ color: c.text, font: {{ size: 11 }}, boxWidth: 10, padding: 8 }} }} }}
        }})
    }});
    if (opts.onClick) {{
        ctx.onclick = function(evt) {{
            var points = chart.getElementsAtEventForMode(evt, 'nearest', {{ intersect: true }}, true);
            if (points.length) opts.onClick(points[0].index, labels[points[0].index]);
        }};
    }}
    return registerChart(canvasId, chart, {{ type: 'donut', i18nLabels: i18nLabels, colorVars: colorVars }});
}}
function createLineChart(canvasId, i18nLabels, datasets, opts) {{
    var ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    opts = opts || {{}};
    var labels = labelsFor(i18nLabels);
    var lang = currentChartLang();
    var c = {{ border: cssVar('--cp-border'), muted: cssVar('--cp-text-muted'), accent: cssVar('--cp-accent'), text: cssVar('--cp-text') }};
    var chartDatasets = (datasets || []).map(function(ds) {{
        var color = ds.colorVar ? cssVar(ds.colorVar) : (ds.color || c.accent);
        return {{
            label: ds.label ? (ds.label[lang] || ds.label.en) : '',
            data: ds.data, borderColor: color, backgroundColor: color + '33',
            fill: !!ds.fill, tension: 0.3,
            borderDash: ds.dashed ? [6, 4] : undefined,
            pointRadius: 3, pointBackgroundColor: color
        }};
    }});
    var chart = new Chart(ctx, {{
        type: 'line',
        data: {{ labels: labels, datasets: chartDatasets }},
        options: baseChartOptions({{
            scales: {{
                x: {{ grid: {{ color: c.border }}, ticks: {{ color: c.muted }} }},
                y: {{
                    beginAtZero: true, grid: {{ color: c.border }}, ticks: {{
                        color: c.muted,
                        callback: opts.currency ? function(v) {{ return '$' + Number(v).toLocaleString(); }} : undefined
                    }}
                }}
            }},
            plugins: {{
                legend: {{ display: (datasets || []).length > 1, labels: {{ color: c.text }} }},
                tooltip: {{ callbacks: opts.currency ? {{ label: function(ctx) {{ return (ctx.dataset.label ? ctx.dataset.label + ': ' : '') + '$' + Number(ctx.parsed.y).toLocaleString(); }} }} : undefined }}
            }}
        }})
    }});
    return registerChart(canvasId, chart, {{ type: 'line', i18nLabels: i18nLabels, datasets: datasets }});
}}
function createStackedBarChart(canvasId, i18nLabels, series, opts) {{
    var ctx = document.getElementById(canvasId);
    if (!ctx || typeof Chart === 'undefined') return null;
    opts = opts || {{}};
    var labels = labelsFor(i18nLabels);
    var lang = currentChartLang();
    var c = {{ border: cssVar('--cp-border'), muted: cssVar('--cp-text-muted'), text: cssVar('--cp-text') }};
    var chartDatasets = (series || []).map(function(s) {{
        return {{ label: s.label ? (s.label[lang] || s.label.en) : '', data: s.data, backgroundColor: resolveColorVars([s.colorVar])[0], borderRadius: 4 }};
    }});
    var chart = new Chart(ctx, {{
        type: 'bar',
        data: {{ labels: labels, datasets: chartDatasets }},
        options: baseChartOptions({{
            indexAxis: opts.horizontal ? 'y' : 'x',
            scales: {{
                x: {{ stacked: true, grid: {{ color: c.border }}, ticks: {{ color: c.muted }} }},
                y: {{ stacked: true, grid: {{ color: c.border }}, ticks: {{ color: c.text }} }}
            }},
            plugins: {{ legend: {{ display: true, labels: {{ color: c.text, font: {{ size: 11 }} }} }} }}
        }})
    }});
    return registerChart(canvasId, chart, {{ type: 'stacked', i18nLabels: i18nLabels, series: series }});
}}
function updateAllChartsTheme() {{
    Object.keys(window.__wafCharts || {{}}).forEach(function(id) {{
        var entry = window.__wafCharts[id];
        var chart = entry.chart, meta = entry.meta || {{}};
        if (!chart) return;
        var c = {{ text: cssVar('--cp-text'), muted: cssVar('--cp-text-muted'), border: cssVar('--cp-border'), surface: cssVar('--cp-surface') }};
        if (chart.options.scales) {{
            Object.keys(chart.options.scales).forEach(function(axis) {{
                var scale = chart.options.scales[axis];
                if (scale.grid) scale.grid.color = c.border;
                if (scale.ticks) scale.ticks.color = (axis === 'y') ? c.text : c.muted;
            }});
        }}
        if (chart.options.plugins && chart.options.plugins.legend && chart.options.plugins.legend.labels) {{
            chart.options.plugins.legend.labels.color = c.text;
        }}
        if (chart.options.plugins && chart.options.plugins.tooltip) {{
            chart.options.plugins.tooltip.backgroundColor = c.surface;
            chart.options.plugins.tooltip.titleColor = c.text;
            chart.options.plugins.tooltip.bodyColor = c.text;
            chart.options.plugins.tooltip.borderColor = c.border;
        }}
        if (meta.colorVars && chart.data.datasets[0]) {{
            chart.data.datasets[0].backgroundColor = resolveColorVars(meta.colorVars);
        }}
        if (chart.config.type === 'doughnut' && chart.data.datasets[0]) {{
            chart.data.datasets[0].borderColor = c.surface;
        }}
        chart.update('none');
    }});
}}
function refreshAllChartLabels() {{
    var lang = currentChartLang();
    Object.keys(window.__wafCharts || {{}}).forEach(function(id) {{
        var entry = window.__wafCharts[id];
        var chart = entry.chart, meta = entry.meta || {{}};
        if (!chart || !meta.i18nLabels) return;
        chart.data.labels = meta.i18nLabels[lang] || meta.i18nLabels.en;
        if (meta.datasets) {{
            chart.data.datasets.forEach(function(ds, idx) {{
                if (meta.datasets[idx] && meta.datasets[idx].label) {{
                    ds.label = meta.datasets[idx].label[lang] || meta.datasets[idx].label.en;
                }}
            }});
        }}
        if (meta.series) {{
            chart.data.datasets.forEach(function(ds, idx) {{
                if (meta.series[idx] && meta.series[idx].label) {{
                    ds.label = meta.series[idx].label[lang] || meta.series[idx].label.en;
                }}
            }});
        }}
        chart.update('none');
    }});
}}
function isEffectivelyHidden(el) {{
    while (el) {{
        if (el.hidden || getComputedStyle(el).display === 'none') return true;
        el = el.parentElement;
    }}
    return false;
}}
function initAllCharts() {{
    if (typeof Chart === 'undefined' || !window.WAF_CHART_DATA) return;
    Object.keys(window.WAF_CHART_DATA).forEach(function(key) {{
        var cfg = window.WAF_CHART_DATA[key];
        if (!cfg || window.__wafCharts[cfg.canvasId]) return;
        var canvas = document.getElementById(cfg.canvasId);
        // Chart.js permanently zero-sizes a canvas created while its container chain is
        // hidden (display:none tab panels) - defer creation until the tab becomes visible
        // (setDashboardView calls initAllCharts() again on every view switch).
        if (!canvas || isEffectivelyHidden(canvas)) return;
        if (cfg.type === 'hbar') createHorizontalBarChart(cfg.canvasId, cfg.labels, cfg.values, cfg.colorVars, cfg.opts);
        else if (cfg.type === 'donut') createDonutChart(cfg.canvasId, cfg.labels, cfg.values, cfg.colorVars, cfg.opts);
        else if (cfg.type === 'line') createLineChart(cfg.canvasId, cfg.labels, cfg.datasets, cfg.opts);
        else if (cfg.type === 'stacked') createStackedBarChart(cfg.canvasId, cfg.labels, cfg.series, cfg.opts);
    }});
}}
document.addEventListener('click', function(e) {{
    var btn = e.target.closest('.ctt-btn');
    if (!btn) return;
    var block = btn.closest('.chart-table-block');
    if (!block) return;
    var view = btn.getAttribute('data-view');
    block.querySelectorAll('.ctt-btn').forEach(function(b) {{ b.classList.toggle('is-active', b === btn); }});
    var chartPanel = block.querySelector('.ctt-chart');
    var tablePanel = block.querySelector('.ctt-table');
    if (chartPanel) chartPanel.hidden = view !== 'chart';
    if (tablePanel) tablePanel.hidden = view !== 'table';
}});
</script>
<script>
function updateThemeButton() {{
    var btn = document.getElementById('themeToggleBtn');
    if (!btn) return;
    var dark = document.documentElement.getAttribute('data-theme') === 'dark';
    var spanish = document.body.classList.contains('lang-es');
    var portuguese = document.body.classList.contains('lang-pt');
    var label = dark
        ? (spanish ? 'Cambiar a modo claro' : portuguese ? 'Mudar para modo claro' : 'Switch to light mode')
        : (spanish ? 'Cambiar a modo nocturno' : portuguese ? 'Mudar para modo noturno' : 'Switch to night mode');
    btn.textContent = dark ? '☀' : '☾';
    btn.setAttribute('aria-label', label);
    btn.setAttribute('title', label);
    btn.setAttribute('aria-pressed', String(dark));
}}
function setTheme(theme, persist) {{
    var normalized = theme === 'dark' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', normalized);
    if (persist !== false) {{
        try {{ localStorage.setItem('wafDashboardTheme', normalized); }} catch (e) {{}}
    }}
    updateThemeButton();
    if (typeof updateAllChartsTheme === 'function') updateAllChartsTheme();
}}
function toggleTheme() {{
    setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
}}
var LANG_CYCLE = ['en', 'es', 'pt'];
var LANG_NEXT_LABEL = {{ en: 'ES 🇪🇸', es: 'PT 🇧🇷', pt: 'EN 🇬🇧' }};
function setLang(lang) {{
    if (LANG_CYCLE.indexOf(lang) === -1) {{ lang = 'en'; }}
    document.body.classList.toggle('lang-es', lang === 'es');
    document.body.classList.toggle('lang-pt', lang === 'pt');
    document.documentElement.lang = lang;
    var btn = document.getElementById('langToggleBtn');
    if (btn) {{ btn.textContent = LANG_NEXT_LABEL[lang]; }}
    updateThemeButton();
    try {{ localStorage.setItem('wafDashboardLang', lang); }} catch (e) {{}}
    moveTabThumb(currentDashboardView, true);
    if (typeof refreshAllChartLabels === 'function') refreshAllChartLabels();
}}
function toggleLang() {{
    var current = document.body.classList.contains('lang-es') ? 'es' : (document.body.classList.contains('lang-pt') ? 'pt' : 'en');
    var next = LANG_CYCLE[(LANG_CYCLE.indexOf(current) + 1) % LANG_CYCLE.length];
    setLang(next);
}}
function toggleTableRows(btn) {{
    var wrap = btn.previousElementSibling;
    if (!wrap) return;
    var expanded = btn.classList.toggle('expanded');
    var rows = wrap.querySelectorAll('.extra-row');
    for (var i = 0; i < rows.length; i++) {{ rows[i].hidden = !expanded; }}
}}

/* Double-click any element carrying a data-resource-id (ARM resource id) to open it
   directly in the Azure Portal. Single delegated listener - works for table rows AND
   action-item resource-list entries, current or future. */
document.addEventListener('dblclick', function(e) {{
    var el = e.target.closest('[data-resource-id]');
    if (!el) return;
    var id = el.getAttribute('data-resource-id');
    if (!id) return;
    window.open('https://portal.azure.com/#resource' + id + '/overview', '_blank', 'noopener');
}});

/* ---- Per-table filters: auto-detects categorical columns (Subscription, Resource Group,
   Severity, etc.) in every rendered table and injects a dropdown per column plus a free-text
   search box - no per-table Python changes needed, this works for any current/future table. */
function initTableFilters() {{
    // Hidden (non-active) tab panels have display:none, and under that condition
    // .innerText stops respecting the nested .i18n-en/.i18n-es/.i18n-pt display:none
    // rules in most browsers (falls back to concatenating all of them) - so headers/
    // cells must be read with textContent + an explicit pick of the active language span.
    var spanish = document.body.classList.contains('lang-es');
    var portuguese = document.body.classList.contains('lang-pt');
    var activeI18nClass = portuguese ? 'i18n-pt' : (spanish ? 'i18n-es' : 'i18n-en');
    function visibleText(el) {{
        var span = el.querySelector('.' + activeI18nClass);
        return (span ? span.textContent : el.textContent).trim();
    }}

    var wraps = document.querySelectorAll('.security-table-wrap');
    wraps.forEach(function(wrap) {{
        var table = wrap.querySelector('table');
        if (!table) return;
        var theadRow = table.querySelector('thead tr');
        var tbody = table.querySelector('tbody');
        if (!theadRow || !tbody) return;
        var rows = Array.prototype.slice.call(tbody.rows);
        if (rows.length < 2) return;
        var ths = Array.prototype.slice.call(theadRow.cells);
        var colCount = ths.length;

        var colValues = ths.map(function() {{ return {{}}; }});
        rows.forEach(function(row) {{
            for (var c = 0; c < colCount && c < row.cells.length; c++) {{
                var txt = visibleText(row.cells[c]);
                if (txt) colValues[c][txt] = true;
            }}
        }});

        var filterCols = [];
        var subColIndex = -1;
        var subHeaderNames = ['Subscription', 'Suscripción', 'Assinatura'];
        for (var c = 0; c < colCount; c++) {{
            var headerLabel = visibleText(ths[c]);
            if (subHeaderNames.indexOf(headerLabel) !== -1) {{ subColIndex = c; continue; }}
            var distinct = Object.keys(colValues[c]);
            if (distinct.length >= 2 && distinct.length <= 20 && distinct.length < rows.length * 0.6) {{
                filterCols.push({{ index: c, label: headerLabel, values: distinct.sort() }});
            }}
        }}
        if (filterCols.length === 0 && rows.length < 6) return;

        var bar = document.createElement('div');
        bar.className = 'table-filter-bar';
        var selects = [];
        filterCols.slice(0, 3).forEach(function(col) {{
            var select = document.createElement('select');
            select.className = 'table-filter-select';
            select.setAttribute('data-col', String(col.index));
            var allOpt = document.createElement('option');
            allOpt.value = '';
            allOpt.textContent = col.label + ' (' + col.values.length + ')';
            select.appendChild(allOpt);
            col.values.forEach(function(v) {{
                var opt = document.createElement('option');
                opt.value = v; opt.textContent = v;
                select.appendChild(opt);
            }});
            select.addEventListener('change', applyFilters);
            bar.appendChild(select);
            selects.push(select);
        }});

        var search = document.createElement('input');
        search.type = 'search';
        search.className = 'table-filter-search';
        search.placeholder = spanish ? 'Buscar en esta tabla...' : portuguese ? 'Pesquisar nesta tabela...' : 'Search this table...';
        search.addEventListener('input', applyFilters);
        bar.appendChild(search);

        var countLabel = document.createElement('span');
        countLabel.className = 'table-filter-count';
        bar.appendChild(countLabel);
        wrap.parentNode.insertBefore(bar, wrap);

        var emptyMsg = document.createElement('div');
        emptyMsg.className = 'table-filter-empty';
        emptyMsg.textContent = spanish ? 'Ninguna fila coincide con los filtros.' : portuguese ? 'Nenhuma linha corresponde aos filtros.' : 'No rows match these filters.';
        wrap.parentNode.insertBefore(emptyMsg, wrap.nextSibling);

        function applyFilters() {{
            var searchTerm = search.value.trim().toLowerCase();
            var anySelectActive = selects.some(function(s) {{ return s.value !== ''; }});
            var filtersActive = !!searchTerm || anySelectActive;
            var visible = 0;
            rows.forEach(function(row) {{
                if (row.dataset.globalSubHidden === '1') {{ row.style.display = 'none'; return; }}
                if (!filtersActive) {{
                    row.style.display = '';
                    if (row.classList.contains('extra-row')) {{ row.hidden = true; }} else {{ visible++; }}
                    return;
                }}
                var matches = true;
                for (var i = 0; i < selects.length; i++) {{
                    var col = Number(selects[i].getAttribute('data-col'));
                    var want = selects[i].value;
                    if (want && row.cells[col] && visibleText(row.cells[col]) !== want) {{ matches = false; break; }}
                }}
                if (matches && searchTerm) {{ matches = row.textContent.toLowerCase().indexOf(searchTerm) !== -1; }}
                if (matches) {{ row.hidden = false; row.style.display = ''; visible++; }} else {{ row.style.display = 'none'; }}
            }});
            countLabel.textContent = filtersActive ? (visible + ' / ' + rows.length) : '';
            emptyMsg.style.display = (filtersActive && visible === 0) ? 'block' : 'none';
            var moreBtn = wrap.querySelector('.table-more-note');
            if (moreBtn) {{ moreBtn.style.display = filtersActive ? 'none' : ''; }}
        }}
        wrap.__applyLocalFilters = applyFilters;
        wrap.__subColIndex = subColIndex;
        applyFilters();
    }});
}}

/* ---- Dashboard-wide subscription filter (#globalSubFilterBar, visible on every tab):
   (1) syncs each Cost-tab pivot block's own checkboxes to match and re-renders it, and
   (2) hides/shows rows in every OTHER table (any tab) that has a Subscription column,
   combined with that table's own local filters via row.dataset.globalSubHidden. */
function initGlobalSubscriptionFilter() {{
    var bar = document.getElementById('globalSubFilterBar');
    if (!bar) return;
    var checks = Array.prototype.slice.call(bar.querySelectorAll('.global-sub-check input'));
    if (checks.length === 0) return;

    var spanish = document.body.classList.contains('lang-es');
    var portuguese = document.body.classList.contains('lang-pt');
    var activeI18nClass = portuguese ? 'i18n-pt' : (spanish ? 'i18n-es' : 'i18n-en');
    function visibleText(el) {{
        var span = el.querySelector('.' + activeI18nClass);
        return (span ? span.textContent : el.textContent).trim();
    }}
    var subHeaderNames = ['Subscription', 'Suscripci\u00f3n', 'Assinatura'];

    // Tables may show a subscription by display name or by GUID: every known name/id is either
    // selected or not, while values that aren't a known subscription are never hidden.
    var known = {{}};
    checks.forEach(function(c) {{
        known[c.value.toLowerCase()] = true;
        if (c.dataset.subId) known[c.dataset.subId.toLowerCase()] = true;
    }});

    function apply() {{
        var selected = {{}};
        checks.forEach(function(c) {{
            if (!c.checked) return;
            selected[c.value.toLowerCase()] = true;
            if (c.dataset.subId) selected[c.dataset.subId.toLowerCase()] = true;
        }});
        var selectedCount = checks.filter(function(c) {{ return c.checked; }}).length;
        activeSubscriptions = selectedCount === checks.length ? null : selected;
        var countEl = bar.querySelector('.global-sub-filter-count');
        if (countEl) countEl.textContent = selectedCount === checks.length ? '' : selectedCount + ' / ' + checks.length;

        document.querySelectorAll('.cost-pivot-block').forEach(function(block) {{
            block.querySelectorAll('.cost-pivot-sub-check input').forEach(function(c) {{
                var want = !!selected[c.value.toLowerCase()];
                if (c.checked !== want) {{
                    c.checked = want;
                    c.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            }});
        }});

        document.querySelectorAll('.security-table-wrap').forEach(function(wrap) {{
            var table = wrap.querySelector('table');
            if (!table) return;
            var theadRow = table.querySelector('thead tr');
            var tbody = table.querySelector('tbody');
            if (!theadRow || !tbody) return;
            var subCol = (typeof wrap.__subColIndex === 'number') ? wrap.__subColIndex : -1;
            if (subCol === -1) {{
                var ths = Array.prototype.slice.call(theadRow.cells);
                for (var i = 0; i < ths.length; i++) {{
                    if (subHeaderNames.indexOf(visibleText(ths[i])) !== -1) {{ subCol = i; break; }}
                }}
            }}
            if (subCol === -1) return;
            var rows = Array.prototype.slice.call(tbody.rows);
            rows.forEach(function(row) {{
                if (!row.cells[subCol]) return;
                var val = visibleText(row.cells[subCol]).toLowerCase();
                row.dataset.globalSubHidden = (!val || !known[val] || selected[val]) ? '' : '1';
            }});
            if (wrap.__applyLocalFilters) {{
                wrap.__applyLocalFilters();
            }} else {{
                rows.forEach(function(row) {{
                    if (row.dataset.globalSubHidden === '1') {{ row.style.display = 'none'; }}
                    else if (!row.classList.contains('extra-row')) {{ row.style.display = ''; }}
                }});
            }}
        }});
        refreshScores();
        applyDismissedState();
    }}

    checks.forEach(function(c) {{ c.addEventListener('change', apply); }});
    bar.querySelectorAll('.cost-pivot-selectall').forEach(function(btn) {{
        btn.addEventListener('click', function() {{
            var mode = btn.getAttribute('data-mode');
            checks.forEach(function(c) {{ c.checked = (mode === 'all'); }});
            apply();
        }});
    }});
}}
var DASHBOARD_VIEWS = ['overview', 'reliability', 'security', 'cost', 'operational', 'performance', 'governance'];
function escapeHtml(s) {{
    return String(s).replace(/[&<>"']/g, function(c) {{
        return {{ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }}[c];
    }});
}}
function initCostPivotBlocks() {{
    document.querySelectorAll('.cost-pivot-block[data-payload]').forEach(function(block) {{
        if (block.dataset.wired) return;
        block.dataset.wired = '1';
        var payload;
        try {{ payload = JSON.parse(block.getAttribute('data-payload')); }} catch (e) {{ return; }}
        var target = block.querySelector('.cost-pivot-render-target');
        var checks = Array.prototype.slice.call(block.querySelectorAll('.cost-pivot-sub-check input'));
        var spanish = document.body.classList.contains('lang-es');
        var portuguese = document.body.classList.contains('lang-pt');
        var t = {{
            service: portuguese ? 'Servi\u00e7o' : (spanish ? 'Servicio' : 'Service'),
            total: 'Total',
            noMatch: portuguese ? 'Nenhuma linha corresponde \u00e0 sele\u00e7\u00e3o.' : (spanish ? 'Ninguna fila coincide con la selecci\u00f3n.' : 'No rows match this selection.'),
            more: portuguese ? 'mais' : (spanish ? 'm\u00e1s' : 'more'),
            showLess: portuguese ? 'Mostrar menos' : (spanish ? 'Mostrar menos' : 'Show less'),
        }};

        function render() {{
            var selected = checks.map(function(c) {{ return c.checked ? c.value : null; }}).filter(Boolean);
            var rows = payload.rows.filter(function(r) {{ return selected.indexOf(r.sub) !== -1; }});
            var totalsByService = {{}}, cellsByServiceMonth = {{}};
            rows.forEach(function(r) {{
                var key = r.svc + '|' + r.month;
                cellsByServiceMonth[key] = (cellsByServiceMonth[key] || 0) + r.cost;
                totalsByService[r.svc] = (totalsByService[r.svc] || 0) + r.cost;
            }});
            var services = Object.keys(totalsByService).sort(function(a, b) {{ return totalsByService[b] - totalsByService[a]; }});
            if (services.length === 0) {{
                target.innerHTML = '<p class="empty-state">' + t.noMatch + '</p>';
                return;
            }}
            var header = '<th>' + t.service + '</th>' + payload.months.map(function(m) {{
                return '<th class="num">' + escapeHtml(m) + '</th>';
            }}).join('') + '<th class="num">' + t.total + '</th>';
            var limit = 15;
            var body = services.map(function(svc, i) {{
                var cells = payload.months.map(function(m) {{
                    return '<td class="num">' + Math.round(cellsByServiceMonth[svc + '|' + m] || 0).toLocaleString() + '</td>';
                }}).join('');
                var attr = i >= limit ? ' class="extra-row" hidden' : '';
                return '<tr' + attr + '><td>' + escapeHtml(svc) + '</td>' + cells +
                    '<td class="num"><strong>' + Math.round(totalsByService[svc]).toLocaleString() + '</strong></td></tr>';
            }}).join('');
            var monthTotals = payload.months.map(function(m) {{
                return services.reduce(function(sum, svc) {{ return sum + (cellsByServiceMonth[svc + '|' + m] || 0); }}, 0);
            }});
            var grandTotal = services.reduce(function(sum, svc) {{ return sum + totalsByService[svc]; }}, 0);
            var footer = '<tr class="cost-pivot-total-row"><td>' + t.total + '</td>' +
                monthTotals.map(function(v) {{ return '<td class="num">' + Math.round(v).toLocaleString() + '</td>'; }}).join('') +
                '<td class="num">' + Math.round(grandTotal).toLocaleString() + ' ' + escapeHtml(payload.currency || '') + '</td></tr>';
            var more = '';
            if (services.length > limit) {{
                more = '<button type="button" class="table-more-note" onclick="toggleTableRows(this)">' +
                    '<span class="more-label">+ ' + (services.length - limit) + ' ' + t.more + '</span>' +
                    '<span class="less-label">' + t.showLess + '</span></button>';
            }}
            target.innerHTML = '<div class="cost-pivot-wrap"><table><thead><tr>' + header + '</tr></thead><tbody>' + body + footer + '</tbody></table></div>' + more;
        }}

        checks.forEach(function(c) {{ c.addEventListener('change', render); }});
        block.querySelectorAll('.cost-pivot-selectall').forEach(function(btn) {{
            btn.addEventListener('click', function() {{
                var mode = btn.getAttribute('data-mode');
                checks.forEach(function(c) {{ c.checked = (mode === 'all'); }});
                render();
            }});
        }});
        render();
    }});
}}
var currentDashboardView = 'overview';
function moveTabThumb(view, skipAnimation) {{
    var thumb = document.getElementById('viewTabThumb');
    var tab = document.getElementById(view + 'Tab');
    var bar = document.getElementById('viewTabsBar');
    if (!thumb || !tab || !bar) return;
    if (skipAnimation) {{ thumb.style.transition = 'none'; }}
    var barRect = bar.getBoundingClientRect();
    var tabRect = tab.getBoundingClientRect();
    thumb.style.top = (tabRect.top - barRect.top) + 'px';
    thumb.style.height = tabRect.height + 'px';
    thumb.style.width = tabRect.width + 'px';
    thumb.style.transform = 'translateX(' + (tabRect.left - barRect.left) + 'px)';
    if (skipAnimation) {{
        void thumb.offsetHeight;
        thumb.style.transition = '';
    }}
}}
function setDashboardView(view, animate) {{
    var normalized = DASHBOARD_VIEWS.indexOf(view) !== -1 ? view : 'overview';
    animate = animate !== false;
    for (var i = 0; i < DASHBOARD_VIEWS.length; i++) {{
        var tab = document.getElementById(DASHBOARD_VIEWS[i] + 'Tab');
        if (tab) {{ tab.setAttribute('aria-selected', String(DASHBOARD_VIEWS[i] === normalized)); }}
    }}
    var prevPanel = document.getElementById(currentDashboardView + 'View');
    var nextPanel = document.getElementById(normalized + 'View');
    if (!animate || !prevPanel || prevPanel === nextPanel || !nextPanel) {{
        for (var j = 0; j < DASHBOARD_VIEWS.length; j++) {{
            var panel = document.getElementById(DASHBOARD_VIEWS[j] + 'View');
            if (panel) {{ panel.hidden = DASHBOARD_VIEWS[j] !== normalized; }}
        }}
        if (normalized === 'governance') {{ recalcAllDiagramBoxes(); }}
    }} else {{
        prevPanel.classList.add('view-fade-out');
        window.setTimeout(function() {{
            prevPanel.hidden = true;
            prevPanel.classList.remove('view-fade-out');
            nextPanel.hidden = false;
            nextPanel.classList.add('view-fade-in');
            if (normalized === 'governance') {{ recalcAllDiagramBoxes(); }}
            requestAnimationFrame(function() {{
                requestAnimationFrame(function() {{ nextPanel.classList.remove('view-fade-in'); }});
            }});
        }}, 180);
    }}
    currentDashboardView = normalized;
    moveTabThumb(normalized, !animate);
    try {{ localStorage.setItem('wafDashboardView', normalized); }} catch (e) {{}}
    window.setTimeout(function() {{
        if (typeof initAllCharts === 'function') initAllCharts();
        Object.keys(window.__wafCharts || {{}}).forEach(function(id) {{
            var entry = window.__wafCharts[id];
            if (entry && entry.chart) {{ entry.chart.resize(); }}
        }});
    }}, animate ? 200 : 0);
}}
window.addEventListener('resize', function() {{ moveTabThumb(currentDashboardView, true); }});

/* ---- Governance hierarchy diagram: lazy-expand Resource Groups / Resources ---- */
function recalcDiagramBox(svg) {{
    if (!svg) return;
    try {{
        var bbox = svg.getBBox();
        if (!bbox.width || !bbox.height) return;
        var pad = 16;
        var minX = bbox.x - pad;
        var minY = bbox.y - pad;
        var w = bbox.width + pad * 2;
        var h = bbox.height + pad * 2;
        svg.setAttribute('viewBox', minX + ' ' + minY + ' ' + w + ' ' + h);
        svg.setAttribute('width', Math.round(w));
        svg.setAttribute('height', Math.round(h));
    }} catch (e) {{}}
}}
function recalcAllDiagramBoxes() {{
    var svgs = document.querySelectorAll('.mg-diagram-svg');
    for (var i = 0; i < svgs.length; i++) {{ recalcDiagramBox(svgs[i]); }}
}}

/* ---- Action items: dismiss (persisted) + manually-added custom items (client-side only, localStorage) ---- */
var ACTION_DISMISSED_KEY = 'wafDismissedActions';
var ACTION_CUSTOM_KEY = 'wafCustomActions';
var SEVERITY_COLORS_JS = {{"critical": "var(--cp-critical)", "high": "var(--cp-danger)", "medium": "var(--cp-warning)", "low": "var(--cp-link)", "info": "var(--cp-success)"}};
var PILLAR_NAMES_ES = {{"Reliability": "Confiabilidad", "Security": "Seguridad", "Cost Optimization": "Optimizaci\u00f3n de Costos", "Operational Excellence": "Excelencia Operativa", "Performance Efficiency": "Eficiencia de Rendimiento", "Governance": "Gobernanza"}};
var PILLAR_NAMES_PT = {{"Reliability": "Confiabilidade", "Security": "Seguran\u00e7a", "Cost Optimization": "Otimiza\u00e7\u00e3o de Custos", "Operational Excellence": "Excel\u00eancia Operacional", "Performance Efficiency": "Efici\u00eancia de Desempenho", "Governance": "Governan\u00e7a"}};
var showDismissedActions = false;

function loadJsonList(key) {{
    try {{ return JSON.parse(localStorage.getItem(key) || '[]'); }} catch (e) {{ return []; }}
}}
function saveJsonList(key, arr) {{
    try {{ localStorage.setItem(key, JSON.stringify(arr)); }} catch (e) {{}}
}}

function updateActionCounters() {{
    var allCards = document.querySelectorAll('#actionsListAll .action-item');
    var visibleCount = 0, criticalCount = 0;
    allCards.forEach(function(card) {{
        if (card.style.display === 'none') return;
        visibleCount++;
        var badge = card.querySelector('.severity-badge');
        if (badge && /HIGH|CRITICAL/i.test(badge.textContent)) criticalCount++;
    }});
    var totalEl = document.getElementById('actionItemsCount');
    var statTotal = document.getElementById('totalActionsStat');
    var statCritical = document.getElementById('criticalActionsStat');
    if (totalEl) totalEl.textContent = visibleCount;
    if (statTotal) statTotal.textContent = visibleCount;
    if (statCritical) statCritical.textContent = criticalCount;
    var dismissed = loadJsonList(ACTION_DISMISSED_KEY);
    document.querySelectorAll('.dismissed-count').forEach(function(el) {{ el.textContent = dismissed.length; }});
}}

var activeSeverities = null; /* null = show all severities */
var activeSubscriptions = null; /* null = all subscriptions; otherwise {{lowercased name or id: true}} */

/* An item/card is hidden by the subscription filter only when it lists subscriptions and none are selected
   (items with an unknown subscription list stay visible). */
function cardSubscriptionHidden(card) {{
    if (!activeSubscriptions) return false;
    var raw = card.getAttribute('data-subscriptions');
    if (!raw) return false;
    var subs;
    try {{ subs = JSON.parse(raw); }} catch (e) {{ return false; }}
    subs = Array.isArray(subs) ? subs : Object.keys(subs);
    if (!subs.length) return false;
    return !subs.some(function(s) {{ return !!activeSubscriptions[String(s).toLowerCase()]; }});
}}

/* ---- Pillar scores for the selected subscriptions (mirrors ScoreModel in the generator) ---- */
function wafRound(value) {{ return Math.floor(value + 0.5); }}
function wafTotal(bySubscription, selection) {{
    var total = 0;
    var map = bySubscription || {{}};
    Object.keys(map).forEach(function(sub) {{
        if (!selection || sub === '' || selection[sub.toLowerCase()]) total += map[sub];
    }});
    return total;
}}
function wafPillar(spec, selection) {{
    if (spec.kind === 'advisor') {{
        var pool = wafTotal(spec.pool, selection);
        if (pool <= 0) return {{ available: false }};
        var rows = [], weighted = 0, weightTotal = 0;
        Object.keys(spec.weights).forEach(function(subcategory) {{
            var weight = spec.weights[subcategory];
            var impacted = spec.forcedZero.indexOf(subcategory) !== -1 ? pool : Math.min(wafTotal(spec.impacted[subcategory], selection), pool);
            var healthy = Math.max(0, pool - impacted), pct = healthy / pool * 100;
            weighted += pct * weight;
            weightTotal += weight;
            rows.push({{ subcategory: subcategory, healthy: healthy, total: pool, pct: wafRound(pct * 10) / 10 }});
        }});
        return {{ available: true, score: wafRound(weighted / weightTotal), rows: rows }};
    }}
    if (spec.kind === 'ratio') {{
        var costPool = wafTotal(spec.pool, selection);
        var sources = spec.sources.map(function(s) {{
            return {{ name: s.name, count: wafTotal(s.count, selection), affected: wafTotal(s.affected, selection) }};
        }});
        if (costPool <= 0) return {{ available: false, sources: sources }};
        var impactedCost = Math.min(wafTotal(spec.impacted, selection), costPool);
        var healthyCost = Math.max(0, costPool - impactedCost);
        return {{ available: true, score: wafRound(healthyCost / costPool * 100), sources: sources, impacted: impactedCost, healthy: healthyCost, total: costPool }};
    }}
    var picked = spec.scores.filter(function(pair) {{
        return !selection || pair[0] === '' || selection[pair[0].toLowerCase()];
    }}).map(function(pair) {{ return pair[1]; }});
    if (!picked.length) return {{ available: false }};
    var average = picked.reduce(function(a, b) {{ return a + b; }}, 0) / picked.length;
    return {{ available: true, score: wafRound(average), pct: wafRound(average * 10) / 10, scored: picked.length }};
}}
function wafScoreColor(value) {{
    return value >= 80 ? 'var(--cp-success)' : value >= 50 ? 'var(--cp-warning)' : 'var(--cp-danger)';
}}
function refreshScores() {{
    var model = window.WAF_SCORE_MODEL;
    if (!model) return;
    var results = {{}}, available = [];
    Object.keys(model.pillars).forEach(function(pillar) {{
        results[pillar] = wafPillar(model.pillars[pillar], activeSubscriptions);
        if (results[pillar].available) available.push(results[pillar].score);
    }});
    var overall = available.length ? wafRound(available.reduce(function(a, b) {{ return a + b; }}, 0) / available.length) : null;

    var circle = document.getElementById('overallScoreCircle');
    if (circle) {{
        circle.setAttribute('data-value', overall === null ? 'N/A' : overall);
        circle.style.setProperty('--overall-deg', (overall || 0) * 3.6);
        circle.style.setProperty('--overall-color', wafScoreColor(overall || 0));
    }}
    Object.keys(results).forEach(function(pillar) {{
        var r = results[pillar];
        document.querySelectorAll('[data-score="' + pillar + '"]').forEach(function(el) {{
            el.textContent = r.available ? r.score + (el.getAttribute('data-score-suffix') || '') : 'N/A';
        }});
        document.querySelectorAll('[data-pillar-card="' + pillar + '"] .pillar-bar-fill').forEach(function(fill) {{
            fill.style.width = (r.available ? r.score : 0) + '%';
        }});
    }});
    var chartCfg = window.WAF_CHART_DATA && window.WAF_CHART_DATA.pillarScores;
    if (chartCfg && chartCfg.keys) {{
        chartCfg.values = chartCfg.keys.map(function(key) {{ return results[key] && results[key].available ? results[key].score : 0; }});
        var chartEntry = window.__wafCharts && window.__wafCharts.pillarScoresChart;
        if (chartEntry && chartEntry.chart) {{
            chartEntry.chart.data.datasets[0].data = chartCfg.values;
            chartEntry.chart.update();
        }}
    }}
    document.querySelectorAll('tr[data-subcat]').forEach(function(tr) {{
        var parts = tr.getAttribute('data-subcat').split('|');
        var r = results[parts[0]];
        var row = r && r.available ? r.rows.filter(function(x) {{ return x.subcategory === parts[1]; }})[0] : null;
        var healthyCell = tr.querySelector('.sc-ht');
        if (healthyCell) healthyCell.textContent = row ? row.healthy + '/' + row.total : 'N/A';
        var pctCell = tr.querySelector('.sc-pct');
        if (!pctCell) return;
        var fill = pctCell.querySelector('.inline-bar-fill');
        var label = row ? row.pct.toFixed(1) + '%' : 'N/A';
        if (fill) {{
            fill.style.width = (row ? row.pct : 0) + '%';
            fill.style.background = wafScoreColor(row ? row.pct : 0);
            var span = pctCell.querySelector('span');
            if (span) span.textContent = label;
        }} else {{
            pctCell.textContent = label;
        }}
    }});
    var cost = results['Cost Optimization'];
    document.querySelectorAll('tr[data-cost-source]').forEach(function(tr) {{
        var name = tr.getAttribute('data-cost-source');
        var source = cost && cost.sources ? cost.sources.filter(function(s) {{ return s.name === name; }})[0] : null;
        var ok = !!(cost && cost.available && source);
        var affected = ok ? Math.min(source.affected, cost.total) : 0;
        var healthy = ok ? Math.max(0, cost.total - affected) : 0;
        tr.querySelector('.cs-count').textContent = ok ? source.count + ' / ' + affected : 'N/A';
        tr.querySelector('.cs-healthy').textContent = ok ? healthy + '/' + cost.total : 'N/A';
        tr.querySelector('.cs-score').textContent = ok ? wafRound(healthy / cost.total * 100) + '%' : 'N/A';
    }});
    document.querySelectorAll('tr[data-cost-combined]').forEach(function(tr) {{
        var ok = !!(cost && cost.available);
        var findings = ok ? cost.sources.reduce(function(sum, s) {{ return sum + s.count; }}, 0) : 0;
        tr.querySelector('.cs-count').textContent = ok ? findings + ' / ' + cost.impacted : 'N/A';
        tr.querySelector('.cs-healthy').textContent = ok ? cost.healthy + '/' + cost.total : 'N/A';
    }});
    var secure = results['Security'];
    document.querySelectorAll('[data-secure-pct]').forEach(function(el) {{
        el.textContent = secure && secure.available ? secure.pct.toFixed(1) + '%' : 'N/A';
    }});
    document.querySelectorAll('[data-secure-scored]').forEach(function(el) {{
        el.textContent = secure && secure.available ? secure.scored : 0;
    }});
    refreshSecurityGapNote(model);
}}

/* Security averages only the subscriptions that have a Defender Secure Score. For the selected
   subscriptions that don't, say so instead of leaving a bare N/A. */
function refreshSecurityGapNote(model) {{
    var spec = model.pillars.Security;
    var nodes = document.querySelectorAll('[data-pillar-note="Security"]');
    if (!spec || !nodes.length) return;
    var scored = {{}};
    spec.scores.forEach(function(pair) {{ if (pair[0] !== '') scored[pair[0].toLowerCase()] = true; }});
    var gaps = model.subscriptions.filter(function(name) {{
        var key = name.toLowerCase();
        return (!activeSubscriptions || activeSubscriptions[key]) && !scored[key];
    }});
    if (!gaps.length) {{
        nodes.forEach(function(node) {{ node.hidden = true; node.innerHTML = ''; }});
        return;
    }}
    var names = gaps.map(escapeHtml).join(', ');
    var lead = {{
        en: 'No Defender Secure Score for ' + names + ' (left out of the Security average).',
        es: 'Sin Secure Score de Defender para ' + names + ' (excluidas del promedio de Security).',
        pt: 'Sem Secure Score do Defender para ' + names + ' (fora da m\u00e9dia de Security).'
    }};
    var html = ['en', 'es', 'pt'].map(function(lang) {{ return '<span class="i18n-' + lang + '">' + lead[lang] + '</span>'; }}).join('');
    nodes.forEach(function(node) {{ node.innerHTML = html; node.hidden = false; }});
}}

/* While a subscription filter is active, each action shows how many of its affected resources
   fall inside the selected subscriptions (its title/total always describe the whole tenant). */
function updateScopeBadges() {{
    document.querySelectorAll('.action-item[data-subscriptions]').forEach(function(card) {{
        var badge = card.querySelector('.action-scope');
        if (!activeSubscriptions) {{ if (badge) badge.remove(); return; }}
        var counts;
        try {{ counts = JSON.parse(card.getAttribute('data-subscriptions')); }} catch (e) {{ return; }}
        if (!counts || Array.isArray(counts)) return;
        var total = 0, selected = 0;
        Object.keys(counts).forEach(function(name) {{
            total += counts[name];
            if (activeSubscriptions[name.toLowerCase()]) selected += counts[name];
        }});
        if (!badge) {{
            badge = document.createElement('div');
            badge.className = 'action-scope';
            var anchor = card.querySelector('.action-desc');
            if (anchor) anchor.insertAdjacentElement('afterend', badge); else card.appendChild(badge);
        }}
        var figures = selected + ' / ' + total;
        badge.innerHTML = '<span class="i18n-en">In selected subscriptions: ' + figures + '</span>' +
            '<span class="i18n-es">En las suscripciones seleccionadas: ' + figures + '</span>' +
            '<span class="i18n-pt">Nas assinaturas selecionadas: ' + figures + '</span>';
    }});
}}

/* Overview charts count the action items that pass the subscription filter. */
function refreshActionCharts() {{
    var severityCounts = {{}}, pillarCounts = {{}};
    document.querySelectorAll('#actionsListAll .action-item').forEach(function(card) {{
        if (card.classList.contains('is-sub-hidden')) return;
        var severity = card.getAttribute('data-severity'), pillar = card.getAttribute('data-pillar');
        severityCounts[severity] = (severityCounts[severity] || 0) + 1;
        pillarCounts[pillar] = (pillarCounts[pillar] || 0) + 1;
    }});
    [['actionSeverity', severityCounts], ['actionByPillar', pillarCounts]].forEach(function(pair) {{
        var cfg = window.WAF_CHART_DATA && window.WAF_CHART_DATA[pair[0]];
        if (!cfg || !cfg.keys) return;
        cfg.values = cfg.keys.map(function(key) {{ return pair[1][key] || 0; }});
        var entry = window.__wafCharts && window.__wafCharts[cfg.canvasId];
        if (entry && entry.chart) {{
            entry.chart.data.datasets[0].data = cfg.values;
            entry.chart.update();
        }}
    }});
}}

/* Remediation cards: hide the ones outside the selected subscriptions, then re-apply the
   "show only N, expand for more" rule over the cards that remain. */
function refreshRecommendationCards() {{
    var grid = document.querySelector('.resource-rec-grid');
    if (!grid) return;
    var visibleLimit = Number(grid.getAttribute('data-visible')) || 4;
    var btn = grid.nextElementSibling;
    if (btn && !btn.classList.contains('table-more-note')) btn = null;
    var expanded = !!btn && btn.classList.contains('expanded');
    var shown = 0, extra = 0;
    grid.querySelectorAll('.resource-rec-card').forEach(function(card) {{
        var subHidden = cardSubscriptionHidden(card);
        card.classList.toggle('is-sub-hidden', subHidden);
        if (subHidden) {{ card.classList.remove('extra-row'); card.hidden = false; return; }}
        var isExtra = shown >= visibleLimit;
        card.classList.toggle('extra-row', isExtra);
        card.hidden = isExtra && !expanded;
        shown++;
        if (isExtra) extra++;
    }});
    if (btn) {{
        btn.style.display = extra > 0 ? '' : 'none';
        btn.querySelectorAll('.rec-more-count').forEach(function(el) {{ el.textContent = extra; }});
    }}
}}

function applyDismissedState() {{
    var dismissed = loadJsonList(ACTION_DISMISSED_KEY);
    document.querySelectorAll('.action-item[data-item-id]').forEach(function(card) {{
        var isCustom = card.getAttribute('data-custom') === 'true';
        var id = card.getAttribute('data-item-id');
        var isDismissed = !isCustom && dismissed.indexOf(id) !== -1;
        if (!isCustom) {{
            var checkbox = card.querySelector('.action-check');
            if (checkbox) checkbox.checked = isDismissed;
            card.classList.toggle('is-dismissed', isDismissed);
        }}
        var severity = card.getAttribute('data-severity');
        var severityHidden = !!activeSeverities && activeSeverities.indexOf(severity) === -1;
        card.classList.toggle('is-severity-hidden', severityHidden);
        var subHidden = cardSubscriptionHidden(card);
        card.classList.toggle('is-sub-hidden', subHidden);
        card.style.display = ((isDismissed && !showDismissedActions) || severityHidden || subHidden) ? 'none' : '';
    }});
    updateActionCounters();
    updateScopeBadges();
    refreshActionCharts();
    refreshRecommendationCards();
}}

function applySeverityFilter(severity) {{
    if (severity === 'all') {{
        activeSeverities = null;
    }} else {{
        if (!activeSeverities) activeSeverities = [];
        var idx = activeSeverities.indexOf(severity);
        if (idx === -1) {{ activeSeverities.push(severity); }} else {{ activeSeverities.splice(idx, 1); }}
        if (activeSeverities.length === 0) activeSeverities = null;
    }}
    document.querySelectorAll('.severity-chip').forEach(function(chip) {{
        var chipSeverity = chip.getAttribute('data-severity');
        var isActive = chipSeverity === 'all' ? !activeSeverities : (!!activeSeverities && activeSeverities.indexOf(chipSeverity) !== -1);
        chip.classList.toggle('is-active', isActive);
    }});
    applyDismissedState();
}}

document.querySelectorAll('.severity-chip').forEach(function(chip) {{
    chip.addEventListener('click', function() {{ applySeverityFilter(chip.getAttribute('data-severity')); }});
}});

function toggleDismiss(id, checked) {{
    var dismissed = loadJsonList(ACTION_DISMISSED_KEY);
    var idx = dismissed.indexOf(id);
    if (checked && idx === -1) dismissed.push(id);
    if (!checked && idx !== -1) dismissed.splice(idx, 1);
    saveJsonList(ACTION_DISMISSED_KEY, dismissed);
    applyDismissedState();
}}

function deleteCustomItem(id) {{
    saveJsonList(ACTION_CUSTOM_KEY, loadJsonList(ACTION_CUSTOM_KEY).filter(function(item) {{ return item.id !== id; }}));
    document.querySelectorAll('.action-item[data-item-id="' + id + '"]').forEach(function(card) {{ card.remove(); }});
    updateActionCounters();
}}

document.addEventListener('change', function(e) {{
    if (!e.target.classList || !e.target.classList.contains('action-check')) return;
    var card = e.target.closest('.action-item');
    var id = e.target.getAttribute('data-item-id');
    if (card && card.getAttribute('data-custom') === 'true') {{
        deleteCustomItem(id);
    }} else {{
        toggleDismiss(id, e.target.checked);
    }}
}});

function buildActionCardHtml(item, showPillar) {{
    var color = SEVERITY_COLORS_JS[item.severity] || 'var(--cp-text-muted)';
    var pillarEs = PILLAR_NAMES_ES[item.pillar] || item.pillar;
    var pillarPt = PILLAR_NAMES_PT[item.pillar] || item.pillar;
    var pillarHtml = showPillar
        ? ('<span class="pillar-tag"><span class="i18n-en">' + item.pillar + '</span><span class="i18n-es">' + pillarEs + '</span><span class="i18n-pt">' + pillarPt + '</span></span>')
        : '';
    var resourcesHtml = '';
    if (item.resources && item.resources.length) {{
        resourcesHtml = '<ul class="resource-list">' + item.resources.map(function(r) {{ return '<li>' + r + '</li>'; }}).join('') + '</ul>';
    }}
    return '<div class="action-item" data-item-id="' + item.id + '" data-pillar="' + item.pillar + '" data-severity="' + item.severity + '" data-custom="true" style="border-left: 4px solid ' + color + ';">' +
        '<div class="action-header">' +
            '<label class="action-check-wrap" title="Remove"><input type="checkbox" class="action-check" data-item-id="' + item.id + '"></label>' +
            '<span class="action-num">+</span>' +
            '<span class="severity-badge" style="background:' + color + ';">' + item.severity.toUpperCase() + '</span>' +
            pillarHtml +
            '<span class="action-title">' + item.title + '</span>' +
        '</div>' +
        '<p class="action-desc">' + item.description + '</p>' +
        resourcesHtml +
        '</div>';
}}

function renderCustomItems() {{
    loadJsonList(ACTION_CUSTOM_KEY).forEach(function(item) {{
        var allList = document.getElementById('actionsListAll');
        if (allList && !allList.querySelector('[data-item-id="' + item.id + '"]')) {{
            allList.insertAdjacentHTML('afterbegin', buildActionCardHtml(item, true));
        }}
        var pillarList = document.querySelector('.action-list[data-pillar="' + item.pillar + '"]');
        if (pillarList) {{
            var emptyState = pillarList.querySelector('.empty-state');
            if (emptyState) emptyState.remove();
            if (!pillarList.querySelector('[data-item-id="' + item.id + '"]')) {{
                pillarList.insertAdjacentHTML('afterbegin', buildActionCardHtml(item, false));
            }}
        }}
    }});
    applyDismissedState();
}}

var addActionBtn = document.getElementById('addActionBtn');
var addActionForm = document.getElementById('addActionForm');
var cancelActionBtn = document.getElementById('cancelActionBtn');
var saveActionBtn = document.getElementById('saveActionBtn');
var toggleDismissedBtn = document.getElementById('toggleDismissedBtn');

if (addActionBtn && addActionForm) {{
    addActionBtn.addEventListener('click', function() {{ addActionForm.classList.toggle('is-open'); }});
}}
if (cancelActionBtn && addActionForm) {{
    cancelActionBtn.addEventListener('click', function() {{
        addActionForm.classList.remove('is-open');
        addActionForm.querySelectorAll('input, textarea').forEach(function(el) {{ el.value = ''; }});
    }});
}}
if (saveActionBtn && addActionForm) {{
    saveActionBtn.addEventListener('click', function() {{
        var titleEl = document.getElementById('newActionTitle');
        var descEl = document.getElementById('newActionDesc');
        var pillarEl = document.getElementById('newActionPillar');
        var severityEl = document.getElementById('newActionSeverity');
        var resourcesEl = document.getElementById('newActionResources');
        var title = (titleEl.value || '').trim();
        if (!title) {{ titleEl.focus(); return; }}
        var item = {{
            id: 'custom-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7),
            title: title,
            description: (descEl.value || '').trim() || 'Manually added action item.',
            pillar: pillarEl.value,
            severity: severityEl.value,
            resources: (resourcesEl.value || '').split(',').map(function(s) {{ return s.trim(); }}).filter(Boolean)
        }};
        var custom = loadJsonList(ACTION_CUSTOM_KEY);
        custom.push(item);
        saveJsonList(ACTION_CUSTOM_KEY, custom);
        renderCustomItems();
        addActionForm.classList.remove('is-open');
        titleEl.value = ''; descEl.value = ''; resourcesEl.value = '';
    }});
}}
function refreshDismissedToggleLabel() {{
    if (!toggleDismissedBtn) return;
    var enSpan = toggleDismissedBtn.querySelector('.i18n-en.toggle-label');
    var esSpan = toggleDismissedBtn.querySelector('.i18n-es.toggle-label');
    var ptSpan = toggleDismissedBtn.querySelector('.i18n-pt.toggle-label');
    if (enSpan) enSpan.textContent = toggleDismissedBtn.getAttribute(showDismissedActions ? 'data-hide-en' : 'data-show-en');
    if (esSpan) esSpan.textContent = toggleDismissedBtn.getAttribute(showDismissedActions ? 'data-hide-es' : 'data-show-es');
    if (ptSpan) ptSpan.textContent = toggleDismissedBtn.getAttribute(showDismissedActions ? 'data-hide-pt' : 'data-show-pt');
}}
if (toggleDismissedBtn) {{
    toggleDismissedBtn.addEventListener('click', function() {{
        showDismissedActions = !showDismissedActions;
        refreshDismissedToggleLabel();
        applyDismissedState();
    }});
}}

(function() {{
    setTheme(document.documentElement.getAttribute('data-theme') || 'light', false);
    var saved = 'en';
    try {{ saved = localStorage.getItem('wafDashboardLang') || 'en'; }} catch (e) {{}}
    setLang(saved);
    var savedView = 'overview';
    try {{ savedView = localStorage.getItem('wafDashboardView') || 'overview'; }} catch (e) {{}}
    setDashboardView(savedView, false);
    renderCustomItems();
    applyDismissedState();
    refreshDismissedToggleLabel();
    initTableFilters();
    initCostPivotBlocks();
    initGlobalSubscriptionFilter();
    refreshScores();
    initAllCharts();
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

    # Compare against the last recorded run for this environment (score deltas + resolved items)
    history = ScoreHistoryTracker(base_dir)
    progress = history.compare_and_record(
        engine.get_overall_score(), engine.scores, engine.action_items, DashboardGenerator._action_item_id
    )
    if progress["previous"]:
        print(f"\n📈 Since last scan ({str(progress['previous'].get('timestamp', ''))[:16]}):")
        print(f"   Overall score delta: {progress['overall_delta']:+d}")
        print(f"   Resolved action items: {len(progress['resolved'])}")
        print(f"   New action items: {progress['new_count']}")
    else:
        print("\n📈 First recorded scan for this environment — no history to compare against yet.")

    # Generate dashboard
    print("\n🎨 Generating dashboard HTML...")
    generator = DashboardGenerator(engine, data, base_dir, progress=progress)
    html = generator.generate()
    
    output_path = os.path.join(base_dir, "07_Dashboard", "WAF_Dashboard.html")
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
