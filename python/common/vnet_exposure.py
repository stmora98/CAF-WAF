"""Which VNets host internet-reachable public IPs.

Azure Resource Graph can't express this in a single query (no `let`; at most three
union/join/mv-expand operators per query), so each exposure path is queried separately
and merged here. Used to prioritize DDoS Protection findings in the dashboard.
"""
from __future__ import annotations

from typing import Any, Callable, Dict, Iterable, List

NIC_QUERY = """
resources
| where type =~ "microsoft.network/networkinterfaces"
| mv-expand ipc = properties.ipConfigurations
| extend subnetId = tostring(ipc.properties.subnet.id), publicIpId = tostring(ipc.properties.publicIPAddress.id)
| where isnotempty(subnetId) and isnotempty(publicIpId)
| project subnetId, publicIpId"""

APP_GATEWAY_QUERY = """
resources
| where type =~ "microsoft.network/applicationgateways"
| extend subnetId = tostring(properties.gatewayIPConfigurations[0].properties.subnet.id)
| mv-expand fe = properties.frontendIPConfigurations
| extend publicIpId = tostring(fe.properties.publicIPAddress.id)
| where isnotempty(subnetId) and isnotempty(publicIpId)
| project subnetId, publicIpId"""

EDGE_SERVICES_QUERY = """
resources
| where type in~ ("microsoft.network/azurefirewalls", "microsoft.network/bastionhosts", "microsoft.network/virtualnetworkgateways")
| mv-expand ipc = properties.ipConfigurations
| extend subnetId = tostring(ipc.properties.subnet.id), publicIpId = tostring(ipc.properties.publicIPAddress.id),
         via = case(type =~ "microsoft.network/azurefirewalls", "Azure Firewall", type =~ "microsoft.network/bastionhosts", "Azure Bastion", "Virtual network gateway")
| where isnotempty(subnetId) and isnotempty(publicIpId)
| project subnetId, publicIpId, via"""

LB_FRONTEND_QUERY = """
resources
| where type =~ "microsoft.network/loadbalancers"
| mv-expand fe = properties.frontendIPConfigurations
| extend publicIpId = tostring(fe.properties.publicIPAddress.id)
| where isnotempty(publicIpId)
| project lbId = tolower(id), publicIpId"""

LB_NIC_MEMBERS_QUERY = """
resources
| where type =~ "microsoft.network/networkinterfaces"
| mv-expand ipc = properties.ipConfigurations
| mv-expand pool = ipc.properties.loadBalancerBackendAddressPools
| extend subnetId = tostring(ipc.properties.subnet.id), lbId = tolower(tostring(split(tostring(pool.id), "/backendAddressPools/")[0]))
| where isnotempty(subnetId) and isnotempty(lbId)
| project subnetId, lbId"""

LB_VMSS_MEMBERS_QUERY = """
resources
| where type =~ "microsoft.compute/virtualmachinescalesets"
| mv-expand nic = properties.virtualMachineProfile.networkProfile.networkInterfaceConfigurations
| mv-expand ipc = nic.properties.ipConfigurations
| mv-expand pool = ipc.properties.loadBalancerBackendAddressPools
| extend subnetId = tostring(ipc.properties.subnet.id), lbId = tolower(tostring(split(tostring(pool.id), "/backendAddressPools/")[0]))
| where isnotempty(subnetId) and isnotempty(lbId)
| project subnetId, lbId"""


def _vnet_id(subnet_id: Any) -> str:
    return str(subnet_id).split("/subnets/")[0].lower()


def build_exposure_rows(
    nic_rows: Iterable[Dict[str, Any]],
    app_gateway_rows: Iterable[Dict[str, Any]],
    edge_rows: Iterable[Dict[str, Any]],
    lb_frontend_rows: Iterable[Dict[str, Any]],
    lb_member_rows: Iterable[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Merges every exposure path into one row per VNet (public IP count + exposure paths)."""
    links = []

    def add(rows: Iterable[Dict[str, Any]], default_via: str) -> None:
        for row in rows:
            links.append((_vnet_id(row["subnetId"]), str(row["publicIpId"]).lower(), row.get("via") or default_via))

    add(nic_rows, "Network interface")
    add(app_gateway_rows, "Application Gateway")
    add(edge_rows, "Edge service")

    # Public load balancers expose the VNets that host their backend pool members.
    frontends: Dict[str, List[str]] = {}
    for row in lb_frontend_rows:
        frontends.setdefault(str(row["lbId"]).lower(), []).append(str(row["publicIpId"]).lower())
    for member in lb_member_rows:
        for public_ip_id in frontends.get(str(member["lbId"]).lower(), []):
            links.append((_vnet_id(member["subnetId"]), public_ip_id, "Public load balancer"))

    by_vnet: Dict[str, Dict[str, set]] = {}
    for vnet_id, public_ip_id, via in links:
        entry = by_vnet.setdefault(vnet_id, {"ips": set(), "via": set()})
        entry["ips"].add(public_ip_id)
        entry["via"].add(via)

    rows = [
        {
            "VNetId": vnet_id,
            "VNetName": vnet_id.rsplit("/", 1)[-1],
            "PublicIPCount": len(entry["ips"]),
            "ExposedVia": ", ".join(sorted(entry["via"])),
        }
        for vnet_id, entry in by_vnet.items()
    ]
    return sorted(rows, key=lambda row: row["VNetName"])


def collect_exposure_rows(run: Callable[[str], List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    """Runs every exposure query through `run` (query -> rows) and merges the results."""
    return build_exposure_rows(
        run(NIC_QUERY),
        run(APP_GATEWAY_QUERY),
        run(EDGE_SERVICES_QUERY),
        run(LB_FRONTEND_QUERY),
        run(LB_NIC_MEMBERS_QUERY) + run(LB_VMSS_MEMBERS_QUERY),
    )
