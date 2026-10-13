import importlib.util
import json
import unittest
from html import unescape
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("generate_dashboard", ROOT / "generate-dashboard.py")
DASHBOARD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DASHBOARD)

PROD_ID = "07bfb481-56a0-4a78-9b49-a86373a0d0e9"
DEV_ID = "765f91e0-c098-44ab-b3fe-a90aee6d4eb6"
SUBSCRIPTIONS = [
    {"subscriptionId": PROD_ID, "name": "Produccion"},
    {"subscriptionId": DEV_ID, "name": "Dev"},
]


def sample_data(**sheets):
    return {"discovery": {"Subscriptions": SUBSCRIPTIONS, **sheets}}


class SubscriptionNormalizationTests(unittest.TestCase):
    def test_guid_columns_are_replaced_by_the_subscription_name(self):
        data = sample_data(VMs=[{"name": "a", "subscriptionId": PROD_ID}])
        data["governance"] = {"OrphanedResources": [{"Name": "disk", "Subscription": DEV_ID}]}

        DASHBOARD.attach_subscription_names(data)

        self.assertEqual("Produccion", data["discovery"]["VMs"][0]["Subscription"])
        self.assertEqual("Dev", data["governance"]["OrphanedResources"][0]["Subscription"])

    def test_existing_display_names_are_kept(self):
        data = sample_data(VMs=[{"Subscription": "Already Named", "subscriptionId": PROD_ID}])

        DASHBOARD.attach_subscription_names(data)

        self.assertEqual("Already Named", data["discovery"]["VMs"][0]["Subscription"])

    def test_subscription_is_taken_from_the_arm_resource_id_when_no_column_exists(self):
        data = sample_data(Findings=[{"ResourceId": f"/subscriptions/{DEV_ID.upper()}/resourceGroups/rg/providers/x/y/z"}])

        DASHBOARD.attach_subscription_names(data)

        self.assertEqual("Dev", data["discovery"]["Findings"][0]["Subscription"])

    def test_rows_without_any_subscription_information_are_left_untouched(self):
        data = sample_data(Misc=[{"name": "tenant-level"}])

        DASHBOARD.attach_subscription_names(data)

        self.assertNotIn("Subscription", data["discovery"]["Misc"][0])

    def test_without_a_subscriptions_sheet_nothing_changes(self):
        data = {"discovery": {"VMs": [{"subscriptionId": PROD_ID}]}}

        DASHBOARD.attach_subscription_names(data)

        self.assertNotIn("Subscription", data["discovery"]["VMs"][0])


class ActionSubscriptionTests(unittest.TestCase):
    def engine(self, data):
        return DASHBOARD.InsightEngine(DASHBOARD.attach_subscription_names(data))

    def test_counts_rows_per_subscription(self):
        engine = self.engine(sample_data())
        rows = [{"Subscription": "Dev"}, {"Subscription": "Dev"}, {"Subscription": "Produccion"}]

        self.assertEqual({"Dev": 2, "Produccion": 1}, engine._subscription_counts(rows))

    def test_unknown_subscription_on_any_row_keeps_the_item_visible_everywhere(self):
        engine = self.engine(sample_data())

        self.assertEqual({}, engine._subscription_counts([{"Subscription": "Dev"}, {"name": "no info"}]))

    def test_checklist_findings_use_the_subscription_of_every_resource(self):
        data = sample_data()
        data["checklists"] = {"Findings": [
            {"Guid": "g1", "Text": "t", "Service": "kv", "WafPillar": "Security", "Severity": "High",
             "ResourceId": f"/subscriptions/{PROD_ID}/resourceGroups/a/providers/x/y/1"},
            {"Guid": "g1", "Text": "t", "Service": "kv", "WafPillar": "Security", "Severity": "High",
             "ResourceId": f"/subscriptions/{DEV_ID}/resourceGroups/a/providers/x/y/2"},
        ]}
        engine = self.engine(data)

        engine._analyze_checklists()

        self.assertEqual({"Dev": 1, "Produccion": 1}, engine.action_items[0]["subscriptions"])

    def test_action_cards_expose_their_subscriptions_for_the_client_filter(self):
        action = {
            "title": "t", "description": "d", "pillar": "Security", "severity": "high",
            "subscriptions": {"Dev/Test": 3, "Produccion": 1},
        }
        generator = DASHBOARD.DashboardGenerator(SimpleNamespace(action_items=[action]), {})

        html = generator._render_action_card(action, 1)

        attribute = html.split('data-subscriptions="')[1].split('"')[0]
        self.assertEqual({"Dev/Test": 3, "Produccion": 1}, json.loads(unescape(attribute)))

    def test_cards_without_subscription_data_have_no_attribute(self):
        action = {"title": "t", "description": "d", "pillar": "Security", "severity": "high"}
        generator = DASHBOARD.DashboardGenerator(SimpleNamespace(action_items=[action]), {})

        self.assertNotIn("data-subscriptions", generator._render_action_card(action, 1))


if __name__ == "__main__":
    unittest.main()
