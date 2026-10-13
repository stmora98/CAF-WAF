import importlib.util
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "python"))

SPEC = importlib.util.spec_from_file_location("generate_dashboard", ROOT / "generate-dashboard.py")
DASHBOARD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DASHBOARD)

from common.vnet_exposure import build_exposure_rows  # noqa: E402

VNET_PREFIX = "/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Network/virtualNetworks/"


def vnet(name, ddos=False):
    return {"id": VNET_PREFIX + name, "name": name, "ddosProtection": ddos}


def exposure(name, ips=1, via="Network interface"):
    return {"VNetId": (VNET_PREFIX + name).lower(), "VNetName": name.lower(), "PublicIPCount": ips, "ExposedVia": via}


def ddos_items(discovery):
    engine = DASHBOARD.InsightEngine({"discovery": discovery})
    engine._analyze_reliability()
    return [item for item in engine.action_items if "DDoS" in item["title"]]


class DdosSeverityTests(unittest.TestCase):
    def test_exposed_vnets_are_high_and_unexposed_are_low(self):
        items = ddos_items({
            "VNets": [vnet("exposed"), vnet("internal"), vnet("protected", ddos=True)],
            "VNetPublicExposure": [exposure("Exposed", ips=3, via="Application Gateway")],
        })

        by_severity = {item["severity"]: item for item in items}
        self.assertEqual({"high", "low"}, set(by_severity))
        self.assertTrue(by_severity["high"]["title"].startswith("1 VNets with internet-facing public IPs"))
        self.assertIn("3 public IP(s)", by_severity["high"]["description"])
        self.assertIn("Application Gateway", by_severity["high"]["description"])
        self.assertEqual("exposed", by_severity["high"]["resources"][0]["name"])
        self.assertEqual("internal", by_severity["low"]["resources"][0]["name"])

    def test_missing_exposure_sheet_keeps_default_severity_and_explains_why(self):
        items = ddos_items({"VNets": [vnet("a"), vnet("b")]})

        self.assertEqual(1, len(items))
        self.assertEqual("medium", items[0]["severity"])
        self.assertIn("re-run the Discovery phase", items[0]["description"])

    def test_empty_exposure_sheet_means_no_exposure_found(self):
        items = ddos_items({
            "VNets": [vnet("a")],
            "VNetPublicExposure": [{"Result": "No resources found"}],
        })

        self.assertEqual(["low"], [item["severity"] for item in items])

    def test_exposed_vnets_are_ordered_by_public_ip_count(self):
        items = ddos_items({
            "VNets": [vnet("small"), vnet("big")],
            "VNetPublicExposure": [exposure("small", ips=1), exposure("big", ips=9)],
        })

        high = next(item for item in items if item["severity"] == "high")
        self.assertEqual(["big", "small"], [r["name"] for r in high["resources"]])

    def test_vnet_id_matching_is_case_insensitive(self):
        items = ddos_items({
            "VNets": [{"ResourceId": VNET_PREFIX + "MixedCase", "name": "MixedCase", "ddosProtection": False}],
            "VNetPublicExposure": [exposure("MIXEDCASE")],
        })

        self.assertEqual(["high"], [item["severity"] for item in items])


class ExposureMergeTests(unittest.TestCase):
    def test_merges_paths_and_deduplicates_public_ips_case_insensitively(self):
        rows = build_exposure_rows(
            nic_rows=[
                {"subnetId": VNET_PREFIX + "VnetA/subnets/s1", "publicIpId": "/x/publicIPAddresses/PIP-1"},
                {"subnetId": VNET_PREFIX + "VnetA/subnets/s2", "publicIpId": "/x/publicipaddresses/pip-1"},
            ],
            app_gateway_rows=[],
            edge_rows=[{"subnetId": VNET_PREFIX + "VnetA/subnets/fw", "publicIpId": "/x/publicIPAddresses/PIP-FW", "via": "Azure Firewall"}],
            lb_frontend_rows=[],
            lb_member_rows=[],
        )

        self.assertEqual(1, len(rows))
        self.assertEqual(2, rows[0]["PublicIPCount"])
        self.assertEqual("Azure Firewall, Network interface", rows[0]["ExposedVia"])
        self.assertEqual((VNET_PREFIX + "vneta").lower(), rows[0]["VNetId"])
        self.assertEqual("vneta", rows[0]["VNetName"])

    def test_public_load_balancer_exposes_the_vnet_of_its_backend_members(self):
        lb = "/subscriptions/s1/resourcegroups/rg/providers/microsoft.network/loadbalancers/lb1"

        rows = build_exposure_rows(
            nic_rows=[],
            app_gateway_rows=[],
            edge_rows=[],
            lb_frontend_rows=[{"lbId": lb, "publicIpId": "/x/publicIPAddresses/PIP-LB"}],
            lb_member_rows=[
                {"subnetId": VNET_PREFIX + "VnetB/subnets/web", "lbId": lb.upper()},
                {"subnetId": VNET_PREFIX + "VnetC/subnets/aks", "lbId": lb},
                {"subnetId": VNET_PREFIX + "Other/subnets/x", "lbId": "/some/internal/lb"},
            ],
        )

        self.assertEqual(["vnetb", "vnetc"], [row["VNetName"] for row in rows])
        self.assertTrue(all(row["ExposedVia"] == "Public load balancer" for row in rows))

    def test_no_links_returns_no_rows(self):
        self.assertEqual([], build_exposure_rows([], [], [], [], []))


if __name__ == "__main__":
    unittest.main()
