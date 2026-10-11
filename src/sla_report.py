
import csv
from collections import Counter
from pathlib import Path


# Find the project directory regardless of where the script is launched.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "sample_tickets.csv"
REPORT_DIR = PROJECT_ROOT / "reports"

REQUIRED_COLUMNS = [
    "Ticket_ID",
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


def main():
    if not INPUT_FILE.exists():
        print(f"ERROR: Input file not found: {INPUT_FILE}")
        print("Check that data/sample_tickets.csv exists.")
        return

    warnings = []
    analyzed_rows = []

    # Read and validate the input CSV.
    with INPUT_FILE.open(
        "r", newline="", encoding="utf-8-sig"
    ) as csv_file:
        reader = csv.DictReader(
            line for line in csv_file if line.strip())

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

        for row_number, source_row in enumerate(reader, start=2):
            row = {
                column: (source_row.get(column) or "").strip()
                for column in input_columns
            }

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
                    f"Row {row_number}: SLA target must be greater than zero."
                )
                target_hours = None

            # Use resolution time for resolved tickets and elapsed
            # time for open tickets.
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
                    f"Row {row_number}: measured hours cannot be negative."
                )
                measured_hours = None

            sla_status = classify_sla(status, target_hours, measured_hours)

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
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    # Export ticket-level analysis.
    detail_file = REPORT_DIR / "ticket_analysis.csv"
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
        f"Input file: {INPUT_FILE.name}",
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

    summary_file = REPORT_DIR / "sla_summary.txt"
    summary_file.write_text(
        "\n".join(summary_lines) + "\n",
        encoding="utf-8",
    )

    print("\n".join(summary_lines))
    print("\nReports generated successfully:")
    print(f"  Summary: {summary_file}")
    print(f"  Details: {detail_file}")


if __name__ == "__main__":
    main()
