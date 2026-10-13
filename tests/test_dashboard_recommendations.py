import importlib.util
import unittest
from pathlib import Path
from types import SimpleNamespace


MODULE_PATH = Path(__file__).resolve().parents[1] / "generate-dashboard.py"
SPEC = importlib.util.spec_from_file_location("generate_dashboard", MODULE_PATH)
DASHBOARD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DASHBOARD)


def action(title, pillar="Security", severity="high", description="", resources=None, source="Test source"):
    return {
        "title": title,
        "title_es": title,
        "title_pt": title,
        "pillar": pillar,
        "severity": severity,
        "description": description,
        "description_es": description,
        "description_pt": description,
        "resources": resources or [],
        "source": source,
    }


class RecommendationCardTests(unittest.TestCase):
    def generator(self, actions):
        return DASHBOARD.DashboardGenerator(SimpleNamespace(action_items=actions), {})

    def test_remediation_capability_takes_precedence_over_affected_service(self):
        item = action("27 resource(s) affected: Diagnostic logs in Key Vault should be enabled")

        guide = self.generator([item])._implementation_guide(item)

        self.assertEqual("Azure Monitor", guide["service"])
        self.assertEqual("direct", guide["confidence"])

    def test_cost_remediation_takes_precedence_over_workload_service(self):
        item = action("48 resource(s) affected: Consider App Service reserved instance")

        guide = self.generator([item])._implementation_guide(item)

        self.assertEqual("Azure Cost Management", guide["service"])

    def test_parenthesized_vm_signal_maps_to_virtual_machines(self):
        item = action("Enable Trusted Launch for Existing Generation 2 VM(s)", pillar="Operational Excellence")

        guide = self.generator([item])._implementation_guide(item)

        self.assertEqual("Azure Virtual Machines", guide["service"])

    def test_generic_tls_finding_is_not_forced_to_key_vault_or_defender(self):
        item = action("127 Storage accounts not enforcing TLS 1.2")

        guide = self.generator([item])._implementation_guide(item)

        self.assertEqual("Azure Blob Storage", guide["service"])
        self.assertNotIn(guide["service"], {"Azure Key Vault", "Microsoft Defender for Cloud"})

    def test_unknown_finding_remains_as_architecture_review(self):
        item = action("Validate recovery objectives with workload owners", pillar="Reliability")

        recommendation = self.generator([item])._recommend_azure_resources()[0]

        self.assertIs(item, recommendation["action"])
        self.assertEqual("Reliability architecture review", recommendation["guide"]["service"])
        self.assertEqual("review", recommendation["guide"]["confidence"])

    def test_cards_are_action_driven_and_include_traceability(self):
        item = action(
            "6 resource(s) affected: Enable Backups on your Virtual Machines",
            pillar="Reliability",
            description="Protect the affected workloads.",
            resources=["vm-a"],
            source="Azure Advisor",
        )

        html = self.generator([item])._render_azure_resource_recommendations()

        self.assertIn("Enable Backups on your Virtual Machines", html)
        self.assertIn("Azure Backup", html)
        self.assertIn("Azure Advisor", html)
        self.assertIn("vm-a", html)

    def test_recommendations_are_balanced_across_pillars(self):
        actions = [
            action(f"{100 - index} resource(s) affected: Security issue {index}")
            for index in range(5)
        ]
        actions.append(action("50 VMs without Availability Zones", pillar="Reliability"))

        recommendations = self.generator(actions)._recommend_azure_resources(limit=12)

        by_volume = [rec for rec in recommendations if not rec["reserved"]]
        self.assertEqual(3, sum(rec["action"]["pillar"] == "Security" for rec in by_volume))
        self.assertTrue(any(rec["action"]["pillar"] == "Reliability" for rec in by_volume))

    def test_reserved_slot_surfaces_exposed_high_severity_finding_with_small_scope(self):
        big = [
            action(f"{300 - index} resource(s) affected: Reliability gap {index}", pillar="Reliability", source="Azure Advisor")
            for index in range(4)
        ]
        exposed = action("3 VNets with internet-facing public IPs and no DDoS Protection", pillar="Reliability", source="Resource inventory")
        exposed["internet_exposed"] = True

        recommendations = self.generator(big + [exposed])._recommend_azure_resources()

        reserved = [rec for rec in recommendations if rec["reserved"]]
        self.assertEqual([exposed], [rec["action"] for rec in reserved])
        self.assertEqual(3, sum(not rec["reserved"] and rec["action"]["pillar"] == "Reliability" for rec in recommendations))
        self.assertIs(exposed, recommendations[-1]["action"])

    def test_reserved_slot_prefers_toolkit_measured_findings_over_generic_advisor_ones(self):
        advisor = action("90 resource(s) affected: Advisor finding", pillar="Security", source="Azure Advisor")
        measured = action("5 Key Vaults without Purge Protection", pillar="Security", source="Security posture assessment")
        filler = [action(f"{500 - i} resource(s) affected: Volume {i}", pillar="Security") for i in range(3)]

        reserved = [rec for rec in self.generator(filler + [advisor, measured])._recommend_azure_resources() if rec["reserved"]]

        self.assertEqual([measured], [rec["action"] for rec in reserved])

    def test_reserved_slot_only_uses_high_or_critical_findings_and_one_per_pillar(self):
        filler = [action(f"{500 - i} resource(s) affected: Volume {i}", pillar="Cost Optimization") for i in range(3)]
        medium = action("2 resource(s) affected: Minor cost issue", pillar="Cost Optimization", severity="medium")
        high_a = action("4 resource(s) affected: Small cost gap A", pillar="Cost Optimization")
        high_b = action("3 resource(s) affected: Small cost gap B", pillar="Cost Optimization")

        recommendations = self.generator(filler + [medium, high_a, high_b])._recommend_azure_resources()

        reserved = [rec["action"] for rec in recommendations if rec["reserved"]]
        self.assertEqual([high_a], reserved)

    def test_reserved_card_is_labelled_in_html(self):
        big = [action(f"{300 - i} resource(s) affected: Gap {i}", pillar="Reliability") for i in range(3)]
        small = action("2 resource(s) affected: Small but serious", pillar="Reliability")

        html = self.generator(big + [small])._render_azure_resource_recommendations()

        self.assertEqual(1, html.count("Reserved slot: high severity, smaller scope"))

    def test_reserved_card_is_labelled_and_never_duplicates_a_main_card(self):
        item = action("2 resource(s) affected: Only finding", pillar="Performance Efficiency")

        recommendations = self.generator([item])._recommend_azure_resources()
        html = self.generator([item])._render_azure_resource_recommendations()

        self.assertEqual(1, len(recommendations))
        self.assertFalse(recommendations[0]["reserved"])
        self.assertNotIn("Reserved slot", html)


    def test_only_four_cards_are_visible_and_the_rest_can_be_expanded(self):
        pillars = ["Security", "Reliability", "Cost Optimization", "Operational Excellence", "Performance Efficiency"]
        actions = [action(f"{10 - index} affected: Finding {index}", pillar=pillars[index % 5]) for index in range(6)]

        html = self.generator(actions)._render_azure_resource_recommendations()

        self.assertEqual(6, html.count("<article"))
        self.assertEqual(2, html.count(" hidden>"))
        self.assertEqual(2, html.count("resource-rec-card resource-rec-high extra-row"))
        self.assertIn("toggleTableRows(this)", html)
        self.assertIn('Show <span class="rec-more-count">2</span> more', html)

    def test_no_expand_button_when_four_or_fewer_cards(self):
        actions = [action(f"Finding {index}", pillar=pillar) for index, pillar in enumerate(["Security", "Reliability", "Cost Optimization", "Operational Excellence"])]

        html = self.generator(actions)._render_azure_resource_recommendations()

        self.assertEqual(4, html.count("<article"))
        self.assertNotIn("toggleTableRows", html)
        self.assertNotIn(" hidden>", html)


if __name__ == "__main__":
    unittest.main()
