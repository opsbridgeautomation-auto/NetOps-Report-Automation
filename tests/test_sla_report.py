
import unittest
import tempfile
import zipfile
from pathlib import Path
from src.sla_report import (
    classify_sla,
    validate_ticket_metadata,
    validate_ticket_dates,
    generate_excel_report,
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


class TestTicketDateValidation(unittest.TestCase):

    def test_valid_resolved_ticket_dates(self):
        row = {
            "Created_Date": "2026-10-01 08:00",
            "Resolved_Date": "2026-10-01 10:00",
            "Status": "Resolved",
        }

        self.assertEqual(validate_ticket_dates(row, 2), [])

    def test_missing_created_date(self):
        row = {
            "Created_Date": "",
            "Resolved_Date": "2026-10-01 10:00",
            "Status": "Resolved",
        }

        issues = validate_ticket_dates(row, 2)

        self.assertTrue(
            any("Created_Date is missing" in issue for issue in issues)
        )

    def test_invalid_created_date(self):
        row = {
            "Created_Date": "yesterday",
            "Resolved_Date": "2026-10-01 10:00",
            "Status": "Resolved",
        }

        issues = validate_ticket_dates(row, 2)

        self.assertTrue(
            any("Created_Date must use" in issue for issue in issues)
        )

    def test_resolved_ticket_missing_resolved_date(self):
        row = {
            "Created_Date": "2026-10-01 08:00",
            "Resolved_Date": "",
            "Status": "Resolved",
        }

        issues = validate_ticket_dates(row, 2)

        self.assertTrue(
            any("Resolved_Date is missing" in issue for issue in issues)
        )

    def test_resolution_before_creation_is_reported(self):
        row = {
            "Created_Date": "2026-10-01 10:00",
            "Resolved_Date": "2026-10-01 08:00",
            "Status": "Resolved",
        }

        issues = validate_ticket_dates(row, 2)

        self.assertTrue(
            any("earlier than Created_Date" in issue for issue in issues)
        )

    def test_open_ticket_with_resolved_date_is_reported(self):
        row = {
            "Created_Date": "2026-10-01 08:00",
            "Resolved_Date": "2026-10-01 10:00",
            "Status": "Open",
        }

        issues = validate_ticket_dates(row, 2)

        self.assertTrue(
            any(
                "populated for an open ticket" in issue
                for issue in issues
            )
        )

class TestExcelReportGeneration(unittest.TestCase):

    def test_workbook_contains_expected_sheets(self):
        sample_rows = [
            {
                "Ticket_ID": "TEST001",
                "Created_Date": "2026-10-01 08:00",
                "Resolved_Date": "2026-10-01 10:00",
                "Priority": "High",
                "Status": "Resolved",
                "Category": "Network",
                "Assignment_Group": "Team-A",
                "Site": "Test-Site",
                "SLA_Target_Hours": "4",
                "Resolution_Hours": "2",
                "Elapsed_Hours": "",
                "SLA_Status": "Met",
            },
            {
                "Ticket_ID": "TEST002",
                "Created_Date": "2026-10-01 09:00",
                "Resolved_Date": "",
                "Priority": "Critical",
                "Status": "Open",
                "Category": "Connectivity",
                "Assignment_Group": "Team-B",
                "Site": "Test-Site",
                "SLA_Target_Hours": "4",
                "Resolution_Hours": "",
                "Elapsed_Hours": "8",
                "SLA_Status": "Breached",
            },
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            output_file = generate_excel_report(
                sample_rows, [], temp_dir
            )

            self.assertTrue(Path(output_file).is_file())

            with zipfile.ZipFile(output_file) as workbook:
                workbook_xml = workbook.read(
                    "xl/workbook.xml"
                ).decode("utf-8")

            expected_sheets = [
                "Summary",
                "Tickets",
                "Breaches",
                "Data Quality",
            ]

            for sheet_name in expected_sheets:
                with self.subTest(sheet=sheet_name):
                    self.assertIn(
                        f'name="{sheet_name}"',
                        workbook_xml,
                    )

if __name__ == "__main__":
    unittest.main()
