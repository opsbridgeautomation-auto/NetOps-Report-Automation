
import unittest

from src.sla_report import (
    classify_sla,
    validate_ticket_metadata,
)


class TestSlaClassification(unittest.TestCase):

    def test_resolved_ticket_meets_sla(self):
        self.assertEqual(classify_sla("Resolved", 4, 2), "Met")

    def test_resolved_ticket_breaches_sla(self):
        self.assertEqual(
            classify_sla("Resolved", 4, 8), "Breached"
        )

    def test_open_ticket_within_sla(self):
        self.assertEqual(
            classify_sla("Open", 4, 2), "Within SLA"
        )

    def test_open_ticket_breaches_sla(self):
        self.assertEqual(
            classify_sla("Open", 4, 7), "Breached"
        )

    def test_ticket_at_exact_sla_limit(self):
        self.assertEqual(classify_sla("Resolved", 4, 4), "Met")

    def test_missing_duration_is_unknown(self):
        self.assertEqual(
            classify_sla("Resolved", 4, None), "Unknown"
        )

    def test_invalid_target_is_unknown(self):
        self.assertEqual(
            classify_sla("Resolved", 0, 2), "Unknown"
        )

    def test_unsupported_status_is_unknown(self):
        self.assertEqual(
            classify_sla("Cancelled", 4, 2), "Unknown"
        )


class TestTicketMetadataValidation(unittest.TestCase):

    def test_valid_ticket_metadata(self):
        seen = set()

        issues, skip = validate_ticket_metadata(
            {"Ticket_ID": "INC1001", "Priority": "High"},
            2,
            seen,
        )

        self.assertEqual(issues, [])
        self.assertFalse(skip)

    def test_missing_ticket_id_is_skipped(self):
        issues, skip = validate_ticket_metadata(
            {"Ticket_ID": "", "Priority": "High"},
            2,
            set(),
        )

        self.assertTrue(skip)
        self.assertTrue(
            any("Ticket_ID is missing" in issue for issue in issues)
        )

    def test_duplicate_ticket_id_is_skipped(self):
        seen = set()
        row = {"Ticket_ID": "INC1001", "Priority": "High"}

        _, first_skip = validate_ticket_metadata(row, 2, seen)
        issues, second_skip = validate_ticket_metadata(row, 3, seen)

        self.assertFalse(first_skip)
        self.assertTrue(second_skip)
        self.assertTrue(
            any("Duplicate Ticket_ID" in issue for issue in issues)
        )

    def test_unsupported_priority_is_reported(self):
        issues, skip = validate_ticket_metadata(
            {"Ticket_ID": "INC2001", "Priority": "Urgent"},
            2,
            set(),
        )

        self.assertFalse(skip)
        self.assertTrue(
            any("Unsupported priority" in issue for issue in issues)
        )


if __name__ == "__main__":
    unittest.main()
