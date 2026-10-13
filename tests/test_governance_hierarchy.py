import importlib.util
import unittest
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("generate_dashboard", ROOT / "generate-dashboard.py")
DASHBOARD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DASHBOARD)


class GovernanceHierarchyTests(unittest.TestCase):
    def test_infers_management_group_nodes_from_subscription_parents(self):
        data = {
            "governance": {
                "MgmtGroups": [{"Result": "No data found"}],
                "Subscriptions": [
                    {"subscriptionId": "sub-1", "name": "Production", "mgParent": "Tenant Root Group"},
                    {"subscriptionId": "sub-2", "name": "Development", "mgParent": "Tenant Root Group"},
                ],
            },
            "discovery": {},
        }
        generator = DASHBOARD.DashboardGenerator(SimpleNamespace(), data)

        html = generator._render_governance_view()

        self.assertIn(">1</div><div class=\"label\"", html)
        self.assertIn("Tenant Root Group", html)
        self.assertIn("Production", html)
        self.assertIn("Development", html)
        self.assertIn("se infirieron del valor mgParent", html)

    def test_keeps_subscription_only_fallback_without_parent_values(self):
        data = {
            "governance": {
                "MgmtGroups": [{"Result": "No data found"}],
                "Subscriptions": [{"subscriptionId": "sub-1", "name": "Production", "mgParent": ""}],
            },
            "discovery": {},
        }
        generator = DASHBOARD.DashboardGenerator(SimpleNamespace(), data)

        html = generator._render_governance_view()

        self.assertIn("Production", html)
        self.assertNotIn("se infirieron del valor mgParent", html)


if __name__ == "__main__":
    unittest.main()
