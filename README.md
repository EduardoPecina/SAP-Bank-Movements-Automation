# SAP to Excel Bank Movement Automation

Automates the daily capture of bank transactions from SAP into a
shared Excel cash-control workbook, replacing a manual process of
filtering, exporting, and copy-pasting data every business day.

> **Note:** this is a sanitized portfolio copy. All company-specific
> identifiers (paths, account numbers, company codes, SharePoint
> URLs, employee emails) have been replaced with environment
> variables or generic placeholders — see `.env.example`. No real
> business data is included anywhere in this repository.

## Problem

A finance analyst spent 5–10 minutes every business day manually:

1. Logging into SAP to pull the previous business day's bank
   transactions matching a specific movement class.
2. Exporting that list to a text file.
3. Copying the data by hand into a shared Excel workbook used for
   daily cash control.

This script automates the entire flow end to end.

## What this does

- Calculates the correct business date to query: the last business
  day before today, accounting for weekends, public holidays, **and**
  a daily cutoff time before which the upstream source system hasn't
  published the day's data yet.
- Filters and exports the matching transactions directly from SAP.
- Parses the exported (semi-structured, tab-delimited) file.
- Appends the new rows to a live Excel Table on a cloud-hosted
  (SharePoint) workbook, letting Excel's own formula columns
  recalculate automatically.

## Tech notes

- **SAP GUI Scripting** via COM automation (`pywin32`) — navigating
  screens, filling dynamic-position fields, handling ALV list
  scrolling, and driving a native SAP export dialog using SAP's own
  scriptable fields (no Windows-dialog automation needed for this
  particular flow).
- **Business-day date logic** respecting national holidays and a
  daily data-availability cutoff (via the `holidays` library).
- **Positional parsing** of a legacy tab-delimited export format,
  locating the start of each record by pattern-matching a date field
  rather than assuming a fixed column count.
- **Excel COM automation** (`win32com`) appending rows to a live
  `ListObject` (Excel Table) instead of writing to raw cell ranges,
  so formula-driven columns extend automatically.
- **Defensive design**: guaranteed Excel process cleanup via
  `try/finally` (no orphaned `EXCEL.EXE` processes on failure), and a
  hard stop when the day's source data isn't available yet instead of
  running against stale or partial data.

## Requirements

```
pip install pywin32 holidays python-dotenv
```

Requires an already-open, logged-in SAP GUI session with GUI
Scripting enabled (Options > Accessibility & Scripting > Scripting).

## Configuration

Copy `.env.example` to `.env` and fill in your own values (SAP
company/account codes, export folder, SharePoint workbook URL
template, etc.), then load them before running — e.g. with
`python-dotenv`, or by exporting them in your shell.

## Usage

```bash
python kyriba_movements_automation_portfolio.py
```

The script exits early with no changes made if run before the
configured data-cutoff hour.
