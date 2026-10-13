import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("generate_dashboard", ROOT / "generate-dashboard.py")
DASHBOARD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DASHBOARD)

ScoreModel = DASHBOARD.ScoreModel
SEP = DASHBOARD.KEY_SEPARATOR


def advisor_spec(pool, impacted, forced=()):
    return {"kind": "advisor", "weights": {"A": 60, "B": 40}, "forcedZero": list(forced), "pool": pool, "impacted": impacted}


class AdvisorPillarTests(unittest.TestCase):
    def test_all_subscriptions_matches_the_official_weighted_formula(self):
        spec = advisor_spec({"s1": 100, "s2": 100}, {"A": {"s1": 20, "s2": 0}, "B": {"s1": 10}})

        result = ScoreModel.pillar(spec)

        # A: 90% healthy x60, B: 95% healthy x40 -> 92
        self.assertEqual(92, result["score"])
        self.assertEqual({"A": 90.0, "B": 95.0}, {row["subcategory"]: row["pct"] for row in result["rows"]})
        self.assertEqual(200, result["pool"])

    def test_selection_only_counts_the_selected_subscriptions(self):
        spec = advisor_spec({"s1": 100, "s2": 100}, {"A": {"s1": 20}, "B": {"s1": 10}})

        self.assertEqual(84, ScoreModel.pillar(spec, {"s1"})["score"])  # 80% x 60 + 90% x 40
        self.assertEqual(100, ScoreModel.pillar(spec, {"s2"})["score"])

    def test_unassigned_rows_belong_to_every_selection(self):
        spec = advisor_spec({"s1": 100, "s2": 100}, {"A": {"": 10}})

        self.assertEqual(ScoreModel.pillar(spec, {"s1"})["rows"][0]["healthy"], 90)
        self.assertEqual(ScoreModel.pillar(spec, {"s2"})["rows"][0]["healthy"], 90)

    def test_forced_zero_subcategory_is_fully_impacted_for_any_selection(self):
        spec = advisor_spec({"s1": 100}, {}, forced=["A"])

        self.assertEqual({"A": 0.0, "B": 100.0}, {r["subcategory"]: r["pct"] for r in ScoreModel.pillar(spec, {"s1"})["rows"]})

    def test_selection_without_resources_is_unavailable(self):
        spec = advisor_spec({"s1": 100, "s2": 0}, {})

        result = ScoreModel.pillar(spec, {"s2"})

        self.assertFalse(result["available"])
        self.assertIsNone(result["score"])

    def test_impacted_resources_never_exceed_the_pool(self):
        spec = advisor_spec({"s1": 10}, {"A": {"s1": 50}})

        self.assertEqual(0.0, ScoreModel.pillar(spec)["rows"][0]["pct"])

    def test_half_scores_round_up_like_the_page_script(self):
        spec = {"kind": "ratio", "pool": {"s": 200}, "impacted": {"s": 3}, "sources": []}

        self.assertEqual(99, ScoreModel.pillar(spec)["score"])  # 98.5 -> 99, not banker's 98


class CostAndSecurityPillarTests(unittest.TestCase):
    def test_cost_ratio_and_per_source_figures_follow_the_selection(self):
        spec = {
            "kind": "ratio", "pool": {"s1": 100, "s2": 100}, "impacted": {"s1": 30, "s2": 10},
            "sources": [{"name": "Idle", "unit": "resources", "count": {"s1": 5, "s2": 2}, "affected": {"s1": 5, "s2": 2}}],
        }

        everything, only_s1 = ScoreModel.pillar(spec), ScoreModel.pillar(spec, {"s1"})

        self.assertEqual((80, 160, 7), (everything["score"], everything["healthy"], everything["sources"][0]["count"]))
        self.assertEqual((70, 70, 5), (only_s1["score"], only_s1["healthy"], only_s1["sources"][0]["count"]))

    def test_secure_score_averages_only_selected_subscriptions_that_have_one(self):
        spec = {"kind": "secure", "scores": [["a", 30.0], ["b", 50.0], ["c", 70.0]]}

        self.assertEqual((50, 3), (ScoreModel.pillar(spec)["score"], ScoreModel.pillar(spec)["scored"]))
        self.assertEqual(40, ScoreModel.pillar(spec, {"a", "b", "zzz"})["score"])
        self.assertFalse(ScoreModel.pillar(spec, {"zzz"})["available"])


class EngineDecompositionTests(unittest.TestCase):
    SUBS = [
        {"subscriptionId": "11111111-1111-1111-1111-111111111111", "name": "Prod"},
        {"subscriptionId": "22222222-2222-2222-2222-222222222222", "name": "Dev"},
    ]

    def engine(self):
        data = {
            "discovery": {
                "Subscriptions": self.SUBS,
                "VMs": [
                    {"name": "vm1", "subscriptionId": self.SUBS[0]["subscriptionId"]},
                    {"name": "vm2", "subscriptionId": self.SUBS[1]["subscriptionId"]},
                ],
            },
            "governance": {"ResourceSummary": [
                {"type": "vm", "location": "x", "subscriptionId": self.SUBS[0]["subscriptionId"], "Count": 10},
                {"type": "vm", "location": "x", "subscriptionId": self.SUBS[1]["subscriptionId"], "Count": 30},
            ]},
        }
        return DASHBOARD.InsightEngine(DASHBOARD.attach_subscription_names(data))

    def test_resource_pool_is_split_per_subscription(self):
        self.assertEqual({"Prod": 10, "Dev": 30}, self.engine()._pool_by_subscription())

    def test_without_per_subscription_inventory_the_total_belongs_to_every_selection(self):
        engine = DASHBOARD.InsightEngine({"discovery": {"ResourceSummary": [{"Count": 7}]}})

        self.assertEqual({"": 7}, engine._pool_by_subscription())

    def test_vms_without_zone_are_charged_to_their_own_subscription(self):
        engine = self.engine()
        engine._analyze_reliability()

        spec = engine.score_specs["Reliability"]
        impacted = spec["impacted"]["Zone Resiliency"]
        self.assertEqual({"Prod": 1, "Dev": 1}, impacted)
        self.assertLess(ScoreModel.pillar(spec, {"Prod"})["rows"][0]["pct"], ScoreModel.pillar(spec, {"Dev"})["rows"][0]["pct"])

    def test_selecting_every_subscription_equals_the_unfiltered_score(self):
        engine = self.engine()
        engine.analyze()

        for pillar, spec in engine.score_specs.items():
            everything = ScoreModel.pillar(spec)
            selected = ScoreModel.pillar(spec, {"Prod", "Dev"})
            self.assertEqual(everything["score"], selected["score"], pillar)

    def test_payload_lists_subscriptions_and_pillar_specs(self):
        engine = self.engine()
        engine.analyze()

        payload = engine.score_model_payload()

        self.assertEqual(["Dev", "Prod"], payload["subscriptions"])
        self.assertEqual(set(engine.scores), set(payload["pillars"]))

    def test_qualified_keys_split_back_into_per_subscription_counts(self):
        keys = {f"Prod{SEP}a", f"Prod{SEP}b", f"Dev{SEP}a", f"{SEP}orphan"}

        self.assertEqual({"Prod": 2, "Dev": 1, "": 1}, ScoreModel.group_keys(keys))

    def test_subscription_without_a_secure_score_stays_out_of_the_security_average(self):
        data = {
            "discovery": {"Subscriptions": self.SUBS, "Storage": [
                {"name": "st1", "publicBlob": "true", "subscriptionId": self.SUBS[0]["subscriptionId"]},
            ]},
            "security": {"SecureScores": [{"subscriptionId": self.SUBS[1]["subscriptionId"], "percentageScore": 0.4}]},
        }
        engine = DASHBOARD.InsightEngine(DASHBOARD.attach_subscription_names(data))

        engine._analyze_security()

        # Prod has findings but no Secure Score: it is neither counted as 0 nor as 100.
        self.assertEqual([["Dev", 40.0]], engine.score_specs["Security"]["scores"])
        self.assertEqual(40, engine.scores["Security"])


if __name__ == "__main__":
    unittest.main()
