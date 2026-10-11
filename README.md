
# NetOps Report Automation

**A Python-based IT operations reporting tool developed under OpsBridge.**

NetOps Report Automation transforms structured incident-ticket CSV exports into readable SLA summaries, ticket-level analysis, and formatted Excel reports.

## Features

- Analyze resolved and open tickets against supplied SLA targets.
- Calculate resolved-ticket SLA compliance.
- Identify resolved and open tickets that have breached their SLA.
- Summarize tickets by priority.
- Validate ticket IDs, duplicate records, priorities, and timestamp consistency.
- Flag missing or invalid numeric values.
- Export ticket-level analysis to CSV.
- Generate a formatted Excel workbook with charts and multiple worksheets.
- Run automated tests using Python's built-in unittest framework.

## Excel Workbook

The generated `reports/sla_report.xlsx` contains:

- **Summary:** SLA metrics and a ticket-priority chart.
- **Tickets:** ticket-level analysis with filters and SLA status highlighting.
- **Breaches:** records with a calculated breached SLA status.
- **Data Quality:** validation warnings.

## Technology Stack

- Python
- CSV processing
- XlsxWriter
- Excel-compatible `.xlsx` reporting
- `unittest` for automated testing
- Git and GitHub

## Project Structure

```text
NetOps-Report-Automation/
├── data/
│   └── sample_tickets.csv
├── src/
│   └── sla_report.py
├── tests/
│   └── test_sla_report.py
├── reports/
│   ├── sla_summary.txt
│   ├── ticket_analysis.csv
│   └── sla_report.xlsx
├── requirements.txt
├── README.md
└── .gitignore
```

## Requirements

- Python 3.10 or later
- pip
- XlsxWriter

## Installation

Clone the repository and navigate to its directory.

Install the dependencies:

```powershell
py -m pip install -r requirements.txt
```

## Run the Application

From the project root, execute:

```powershell
py .\src\sla_report.py
```

The application reads `data/sample_tickets.csv` and generates:

- `reports/sla_summary.txt`
- `reports/ticket_analysis.csv`
- `reports/sla_report.xlsx`

## Run Automated Tests

```powershell
py -m unittest discover -s tests -v
```

The test suite covers SLA classification, ticket metadata validation, date validation, and Excel workbook generation.

## Demonstration Dataset

The repository includes synthetic ticket data for demonstration and testing.

For the current sample dataset, the expected results are:

- Total tickets: 12
- Resolved tickets: 8
- Open/in-progress tickets: 4
- Resolved SLA compliance: 62.5%
- Resolved SLA breaches: 3
- Open tickets breaching SLA: 2

## Current Limitations

This is a portfolio prototype, not a certified enterprise reporting system.

- SLA decisions use supplied duration values and targets.
- Business-hours calendars, paused SLA clocks, time zones, and organization-specific SLA policies are not implemented.
- Timestamp validation checks basic consistency but does not independently recalculate SLA durations.
- Results depend on the completeness and correctness of the input data.

Validate the output against the source system before using the reports for operational decisions.

## Security and Data Handling

Use synthetic data for public demonstrations.

Do not commit credentials, customer records, proprietary scripts, or confidential company information. Processing real organizational data requires appropriate authorization and data-handling safeguards.

## Project Status

**Status: Working prototype — v0.5**

Developed as a practical IT operations automation portfolio project under OpsBridge.
