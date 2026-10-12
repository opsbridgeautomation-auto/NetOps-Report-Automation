from datetime import datetime
import csv
from collections import Counter
from pathlib import Path
import xlsxwriter
import argparse
# Find the project directory regardless of where the script is launched.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "sample_tickets.csv"
REPORT_DIR = PROJECT_ROOT / "reports"

REQUIRED_COLUMNS = [
    "Ticket_ID",
    "Created_Date",
    "Resolved_Date",
    "Priority",
    "Status",
    "Category",
    "Assignment_Group",
    "Site",
    "SLA_Target_Hours",
    "Resolution_Hours",
    "Elapsed_Hours",
]

OPEN_STATUSES = {"open", "in progress", "pending", "new", "on hold"}


VALID_PRIORITIES = {"critical", "high", "medium", "low"}


def validate_ticket_metadata(row, row_number, seen_ticket_ids):
    """Validate ticket IDs and priorities.

    Returns:
        (issues, skip_row)
    """
    issues = []
    ticket_id = (row.get("Ticket_ID") or "").strip()

    if not ticket_id:
        issues.append(
            f"Row {row_number}: Ticket_ID is missing; record skipped."
        )
        return issues, True

    normalized_id = ticket_id.casefold()

    if normalized_id in seen_ticket_ids:
        issues.append(
            f"Row {row_number}: Duplicate Ticket_ID "
            f"{ticket_id!r}; duplicate record skipped."
        )
        return issues, True

    seen_ticket_ids.add(normalized_id)

    priority = (row.get("Priority") or "").strip().casefold()

    if not priority:
        issues.append(f"Row {row_number}: Priority is missing.")
    elif priority not in VALID_PRIORITIES:
        issues.append(
            f"Row {row_number}: Unsupported priority "
            f"{row['Priority']!r}."
        )

    return issues, False


def validate_ticket_dates(row, row_number):
    """Check ticket timestamps and basic date consistency."""
    issues = []

    created_text = (row.get("Created_Date") or "").strip()
    resolved_text = (row.get("Resolved_Date") or "").strip()
    status = (row.get("Status") or "").strip().casefold()

    created_at = None
    resolved_at = None

    date_format = "%Y-%m-%d %H:%M"

    # Every ticket should have a valid creation timestamp.
    if not created_text:
        issues.append(
            f"Row {row_number}: Created_Date is missing."
        )
    else:
        try:
            created_at = datetime.strptime(
                created_text, date_format
            )
        except ValueError:
            issues.append(
                f"Row {row_number}: Created_Date must use "
                "YYYY-MM-DD HH:MM format."
            )

    # Resolved tickets must have a valid resolution timestamp.
    if status == "resolved":
        if not resolved_text:
            issues.append(
                f"Row {row_number}: Resolved_Date is missing "
                "for a resolved ticket."
            )
        else:
            try:
                resolved_at = datetime.strptime(
                    resolved_text, date_format
                )
            except ValueError:
                issues.append(
                    f"Row {row_number}: Resolved_Date must use "
                    "YYYY-MM-DD HH:MM format."
                )

            if (
                created_at is not None
                and resolved_at is not None
                and resolved_at < created_at
            ):
                issues.append(
                    f"Row {row_number}: Resolved_Date is earlier "
                    "than Created_Date."
                )

    # An open ticket should not have a resolution timestamp.
    elif status in OPEN_STATUSES and resolved_text:
        issues.append(
            f"Row {row_number}: Resolved_Date is populated "
            "for an open ticket."
        )

    return issues

def classify_sla(status, target_hours, measured_hours):
    """Classify ticket SLA status using its target and measured duration."""

    if target_hours is None or measured_hours is None:
        return "Unknown"

    if target_hours <= 0 or measured_hours < 0:
        return "Unknown"

    status = (status or "").strip().lower()

    if status == "resolved":
        return "Met" if measured_hours <= target_hours else "Breached"

    if status in OPEN_STATUSES:
        return "Within SLA" if measured_hours <= target_hours else "Breached"

    return "Unknown"


def parse_number(value, field_name, row_number, warnings, required=False):
    """Convert a field to a number and report invalid or missing values."""
    raw_value = (value or "").strip()

    if not raw_value:
        if required:
            warnings.append(
                f"Row {row_number}: {field_name} is missing."
            )
        return None

    try:
        return float(raw_value)
    except ValueError:
        warnings.append(
            f"Row {row_number}: {field_name} "
            f"has an invalid number: {raw_value!r}."
        )
        return None

def generate_excel_report(analyzed_rows, warnings, report_dir):
    """Create a formatted Excel workbook for SLA reporting."""

    report_dir = Path(report_dir)
    report_dir.mkdir(parents=True, exist_ok=True)

    output_file = report_dir / "sla_report.xlsx"
    workbook = xlsxwriter.Workbook(str(output_file))

    # Workbook styles
    title_fmt = workbook.add_format({
        "bold": True,
        "font_size": 18,
        "font_color": "white",
        "bg_color": "#17365D",
        "align": "left",
        "valign": "vcenter",
    })

    subtitle_fmt = workbook.add_format({
        "italic": True,
        "font_color": "#666666",
    })

    header_fmt = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#4472C4",
        "border": 1,
        "text_wrap": True,
    })

    label_fmt = workbook.add_format({
        "bold": True,
        "border": 1,
        "bg_color": "#D9EAF7",
    })

    value_fmt = workbook.add_format({"border": 1})
    number_fmt = workbook.add_format({
        "border": 1,
        "num_format": "0",
    })
    percent_fmt = workbook.add_format({
        "border": 1,
        "num_format": "0.0%",
    })

    breach_fmt = workbook.add_format({
        "font_color": "#9C0006",
        "bg_color": "#FFC7CE",
    })

    success_fmt = workbook.add_format({
        "font_color": "#006100",
        "bg_color": "#C6EFCE",
    })

    warning_fmt = workbook.add_format({
        "font_color": "#9C6500",
        "bg_color": "#FFEB9C",
    })

    # Calculate the summary directly from analyzed tickets.
    resolved_rows = [
        row for row in analyzed_rows
        if (row.get("Status") or "").strip().lower() == "resolved"
    ]

    open_rows = [
        row for row in analyzed_rows
        if (row.get("Status") or "").strip().lower() in OPEN_STATUSES
    ]

    resolved_known = [
        row for row in resolved_rows
        if row.get("SLA_Status") in {"Met", "Breached"}
    ]

    resolved_met = sum(
        row.get("SLA_Status") == "Met"
        for row in resolved_known
    )

    resolved_breached = sum(
        row.get("SLA_Status") == "Breached"
        for row in resolved_known
    )

    open_breached = sum(
        row.get("SLA_Status") == "Breached"
        for row in open_rows
    )

    unknown_sla = sum(
        row.get("SLA_Status") == "Unknown"
        for row in analyzed_rows
    )

    compliance = (
        resolved_met / len(resolved_known)
        if resolved_known
        else None
    )

    priority_counts = Counter(
        (row.get("Priority") or "").strip() or "Unspecified"
        for row in analyzed_rows
    )

    # --------------------------------------------------
    # Sheet 1: Management Summary
    # --------------------------------------------------
    summary = workbook.add_worksheet("Summary")
    summary.hide_gridlines(2)
    summary.merge_range("A1:E1", "OPSBRIDGE SLA REPORT", title_fmt)
    summary.set_row(0, 30)
    summary.write("A2", "IT Operations Reporting", subtitle_fmt)

    summary.write("A3", "Metric", header_fmt)
    summary.write("B3", "Value", header_fmt)

    metrics = [
        ("Total tickets", len(analyzed_rows)),
        ("Resolved tickets", len(resolved_rows)),
        ("Open/in-progress tickets", len(open_rows)),
        ("Resolved SLA met", resolved_met),
        ("Resolved SLA breaches", resolved_breached),
        ("Open tickets breaching SLA", open_breached),
        (
            "Resolved SLA compliance",
            compliance if compliance is not None else "N/A",
        ),
        ("Tickets with unknown SLA status", unknown_sla),
    ]

    for row_index, (label, value) in enumerate(metrics, start=3):
        summary.write_string(row_index, 0, label, label_fmt)

        if isinstance(value, (int, float)):
            fmt = (
                percent_fmt
                if label == "Resolved SLA compliance"
                else number_fmt
            )
            summary.write_number(row_index, 1, value, fmt)
        else:
            summary.write_string(row_index, 1, str(value), value_fmt)

    summary.set_column("A:A", 34)
    summary.set_column("B:B", 18)

    # Priority breakdown
    summary.write("D3", "Priority", header_fmt)
    summary.write("E3", "Ticket Count", header_fmt)

    priorities = sorted(priority_counts.items())

    for row_index, (priority, count) in enumerate(priorities, start=3):
        summary.write_string(row_index, 3, str(priority), value_fmt)
        summary.write_number(row_index, 4, count, number_fmt)

    summary.set_column("D:D", 20)
    summary.set_column("E:E", 15)

    if priorities:
        chart = workbook.add_chart({"type": "column"})

        chart.add_series({
            "name": "Ticket Count",
            "categories": [
                "Summary", 3, 3, 2 + len(priorities), 3
            ],
            "values": [
                "Summary", 3, 4, 2 + len(priorities), 4
            ],
            "data_labels": {"value": True},
        })

        chart.set_title({"name": "Tickets by Priority"})
        chart.set_x_axis({"name": "Priority"})
        chart.set_y_axis({
            "name": "Tickets",
            "min": 0,
            "major_gridlines": {"visible": False},
        })
        chart.set_legend({"none": True})
        chart.set_size({"width": 520, "height": 300})

        summary.insert_chart("D10", chart)

    # --------------------------------------------------
    # Sheet 2: All Analyzed Tickets
    # --------------------------------------------------
    ticket_columns = (
        list(analyzed_rows[0].keys())
        if analyzed_rows
        else REQUIRED_COLUMNS + ["SLA_Status"]
    )

    # Keep column names unique in the output.
    ticket_columns = list(dict.fromkeys(ticket_columns))

    tickets_sheet = workbook.add_worksheet("Tickets")
    tickets_sheet.freeze_panes(1, 0)

    tickets_sheet.write_row(0, 0, ticket_columns, header_fmt)

    for row_index, row in enumerate(analyzed_rows, start=1):
        for col_index, column in enumerate(ticket_columns):
            value = row.get(column, "")
            tickets_sheet.write_string(
                row_index,
                col_index,
                "" if value is None else str(value),
            )

    tickets_sheet.autofilter(
        0, 0, len(analyzed_rows), len(ticket_columns) - 1
    )

    for col_index, column in enumerate(ticket_columns):
        width = max(14, min(len(column) + 3, 24))

        if column in {"Created_Date", "Resolved_Date"}:
            width = 21
        elif column == "SLA_Status":
            width = 16

        tickets_sheet.set_column(col_index, col_index, width)

    if analyzed_rows and "SLA_Status" in ticket_columns:
        sla_col = ticket_columns.index("SLA_Status")

        tickets_sheet.conditional_format(
            1, sla_col, len(analyzed_rows), sla_col,
            {
                "type": "text",
                "criteria": "containing",
                "value": "Breached",
                "format": breach_fmt,
            },
        )

        for good_status in ("Met", "Within SLA"):
            tickets_sheet.conditional_format(
                1, sla_col, len(analyzed_rows), sla_col,
                {
                    "type": "text",
                    "criteria": "equal to",
                    "value": good_status,
                    "format": success_fmt,
                },
            )

        tickets_sheet.conditional_format(
            1, sla_col, len(analyzed_rows), sla_col,
            {
                "type": "text",
                "criteria": "equal to",
                "value": "Unknown",
                "format": warning_fmt,
            },
        )

    # --------------------------------------------------
    # Sheet 3: SLA Breaches
    # --------------------------------------------------
    breach_rows = [
        row for row in analyzed_rows
        if row.get("SLA_Status") == "Breached"
    ]

    breaches_sheet = workbook.add_worksheet("Breaches")
    breaches_sheet.freeze_panes(1, 0)
    breaches_sheet.write_row(0, 0, ticket_columns, header_fmt)

    for row_index, row in enumerate(breach_rows, start=1):
        for col_index, column in enumerate(ticket_columns):
            value = row.get(column, "")
            breaches_sheet.write_string(
                row_index,
                col_index,
                "" if value is None else str(value),
            )

    breaches_sheet.autofilter(
        0, 0, len(breach_rows), len(ticket_columns) - 1
    )
    breaches_sheet.set_column(
        0, len(ticket_columns) - 1, 20
    )

    # --------------------------------------------------
    # Sheet 4: Data Quality
    # --------------------------------------------------
    quality_sheet = workbook.add_worksheet("Data Quality")
    quality_sheet.write("A1", "Validation Issue", header_fmt)
    quality_sheet.set_column("A:A", 95)

    if warnings:
        for row_index, warning in enumerate(warnings, start=1):
            quality_sheet.write_string(
                row_index, 0, str(warning)
            )
    else:
        quality_sheet.write_string(
            1, 0, "No validation warnings."
        )

    workbook.close()
    return output_file


def main(input_file=INPUT_FILE, report_dir=REPORT_DIR):
    input_file = Path(input_file)
    report_dir = Path(report_dir)
    if not input_file.exists():
        print(f"ERROR: Input file not found: {input_file}")
        print("Check that the specified input file exists.")
        return

    warnings = []
    analyzed_rows = []
    seen_ticket_ids = set()

    # Read and validate the input CSV.
    
    warnings = []
    analyzed_rows = []
    seen_ticket_ids = set()

    # Read and validate the input CSV.
    with input_file.open(
        "r", newline="", encoding="utf-8-sig"
    ) as csv_file:
        reader = csv.DictReader(
            line for line in csv_file if line.strip()
        )

        if not reader.fieldnames:
            print("ERROR: The input CSV is empty or has no header.")
            return

        missing_columns = [
            column
            for column in REQUIRED_COLUMNS
            if column not in reader.fieldnames
        ]

        if missing_columns:
            print("ERROR: Required columns are missing:")
            for column in missing_columns:
                print(f"  - {column}")
            return

        input_columns = list(reader.fieldnames)

        # IMPORTANT: Process and append every ticket inside this loop.
        for row_number, source_row in enumerate(reader, start=2):
            row = {
                column: (source_row.get(column) or "").strip()
                for column in input_columns
            }

            metadata_issues, skip_row = validate_ticket_metadata(
                row, row_number, seen_ticket_ids
            )
            warnings.extend(metadata_issues)

            if skip_row:
                continue
            warnings.extend(validate_ticket_dates(row, row_number))
            status = row["Status"].strip().lower()

            target_hours = parse_number(
                row["SLA_Target_Hours"],
                "SLA_Target_Hours",
                row_number,
                warnings,
                required=True,
            )

            if target_hours is not None and target_hours <= 0:
                warnings.append(
                    f"Row {row_number}: "
                    "SLA target must be greater than zero."
                )
                target_hours = None

            if status == "resolved":
                measured_hours = parse_number(
                    row["Resolution_Hours"],
                    "Resolution_Hours",
                    row_number,
                    warnings,
                    required=True,
                )
            elif status in OPEN_STATUSES:
                measured_hours = parse_number(
                    row["Elapsed_Hours"],
                    "Elapsed_Hours",
                    row_number,
                    warnings,
                    required=True,
                )
            else:
                measured_hours = None
                warnings.append(
                    f"Row {row_number}: unsupported status "
                    f"{row['Status']!r}."
                )

            if measured_hours is not None and measured_hours < 0:
                warnings.append(
                    f"Row {row_number}: "
                    "measured hours cannot be negative."
                )
                measured_hours = None

            sla_status = classify_sla(
                status, target_hours, measured_hours
            )

            # This line must remain INSIDE the for loop.
            analyzed_rows.append({
                **row,
                "SLA_Status": sla_status,
            })

    # Prepare summary statistics.
    resolved_rows = [
        row for row in analyzed_rows
        if row["Status"].lower() == "resolved"
    ]

    open_rows = [
        row for row in analyzed_rows
        if row["Status"].lower() in OPEN_STATUSES
    ]

    resolved_known = [
        row for row in resolved_rows
        if row["SLA_Status"] in {"Met", "Breached"}
    ]

    resolved_met = sum(
        row["SLA_Status"] == "Met" for row in resolved_known
    )

    resolved_breached = sum(
        row["SLA_Status"] == "Breached" for row in resolved_known
    )

    open_breached = sum(
        row["SLA_Status"] == "Breached" for row in open_rows
    )

    unknown_sla = sum(
        row["SLA_Status"] == "Unknown" for row in analyzed_rows
    )

    compliance = (
        resolved_met / len(resolved_known) * 100
        if resolved_known
        else None
    )

    priority_counts = Counter(
        row["Priority"] or "Unspecified"
        for row in analyzed_rows
    )

    # Create the output directory only when needed.
    report_dir.mkdir(parents=True, exist_ok=True)

    # Export ticket-level analysis.
    detail_file = report_dir / "ticket_analysis.csv"
    output_columns = input_columns + (
        ["SLA_Status"] if "SLA_Status" not in input_columns else []
    )

    with detail_file.open(
        "w", newline="", encoding="utf-8"
    ) as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=output_columns,
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(analyzed_rows)

    # Build the human-readable summary.
    summary_lines = [
        "OPSB​​RIDGE SLA REPORT".replace("\u200b", ""),
        "=" * 32,
        f"Input file: {input_file.name}",
        f"Total tickets: {len(analyzed_rows)}",
        f"Resolved tickets: {len(resolved_rows)}",
        f"Open/in-progress tickets: {len(open_rows)}",
        "",
        "RESOLVED TICKET SLA RESULTS",
        "-" * 32,
        f"Resolved tickets with usable SLA data: {len(resolved_known)}",
        f"SLA met: {resolved_met}",
        f"SLA breached: {resolved_breached}",
        (
            f"Resolved SLA compliance: {compliance:.1f}%"
            if compliance is not None
            else "Resolved SLA compliance: N/A"
        ),
        "",
        "OPEN TICKET SLA RESULTS",
        "-" * 32,
        f"Open tickets currently breaching SLA: {open_breached}",
        f"Tickets with unknown SLA status: {unknown_sla}",
        "",
        "TICKETS BY PRIORITY",
        "-" * 32,
    ]

    for priority, count in sorted(priority_counts.items()):
        summary_lines.append(f"{priority}: {count}")

    summary_lines.extend([
        "",
        "DATA VALIDATION",
        "-" * 32,
    ])

    if warnings:
        summary_lines.extend(f"- {warning}" for warning in warnings)
    else:
        summary_lines.append("No validation warnings.")

    summary_file = report_dir / "sla_summary.txt"
    summary_file.write_text(
        "\n".join(summary_lines) + "\n",
        encoding="utf-8",
    )
    excel_file = generate_excel_report(
    analyzed_rows,
    warnings,
    report_dir,
)
    print("\n".join(summary_lines))
    print("\nReports generated successfully:")
    print(f"  Summary: {summary_file}")
    print(f"  Details: {detail_file}")
    print(f"  Excel:   {excel_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate SLA reports from a ticket CSV."
    )

    parser.add_argument(
        "--input",
        type=Path,
        default=INPUT_FILE,
        help="Path to the input ticket CSV.",
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=REPORT_DIR,
        help="Directory where reports will be saved.",
    )

    args = parser.parse_args()

    main(
        input_file=args.input,
        report_dir=args.output,
    )
