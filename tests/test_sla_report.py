
import unittest

from src.sla_report import classify_sla


class TestSlaClassification(unittest.TestCase):

    def test_resolved_ticket_meets_sla(self):
        self.assertEqual(
            classify_sla("Resolved", 4, 2),
            "Met",
        )

    def test_resolved_ticket_breaches_sla(self):
        self.assertEqual(
            classify_sla("Resolved", 4, 8),
            "Breached",
        )

    def test_open_ticket_within_sla(self):
        self.assertEqual(
            classify_sla("Open", 4, 2),
            "Within SLA",
        )

    def test_open_ticket_breaches_sla(self):
        self.assertEqual(
            classify_sla("Open", 4, 7),
            "Breached",
        )

    def test_ticket_at_exact_sla_limit(self):
        self.assertEqual(
            classify_sla("Resolved", 4, 4),
            "Met",
        )

    def test_missing_duration_is_unknown(self):
        self.assertEqual(
            classify_sla("Resolved", 4, None),
            "Unknown",
        )

    def test_invalid_target_is_unknown(self):
        self.assertEqual(
            classify_sla("Resolved", 0, 2),
            "Unknown",
        )

    def test_unsupported_status_is_unknown(self):
        self.assertEqual(
            classify_sla("Cancelled", 4, 2),
            "Unknown",
        )


if __name__ == "__main__":
    unittest.main()
