"""
SAP TO EXCEL BANK MOVEMENT AUTOMATION
======================================

PORTFOLIO / DEMO VERSION -- all company-specific identifiers (paths,
account numbers, company codes, SharePoint URLs, employee emails) have
been replaced with environment variables or generic placeholders. This
is a sanitized copy of a production script; no real business data is
included.

PROBLEM
-------
A finance analyst spent 5-10 minutes every business day manually:
  1. Logging into SAP to pull the previous business day's bank
     transactions matching a specific movement class.
  2. Exporting that list to a text file.
  3. Copying the data by hand into a shared Excel workbook used for
     daily cash control.

This script automates the entire flow end to end.

WHAT THIS DEMONSTRATES
-----------------------
  - SAP GUI Scripting via COM automation (win32com) -- navigating
    screens, filling dynamic-position fields, handling ALV list
    scrolling, and driving a native SAP export dialog.
  - Business-day date logic respecting national holidays and a daily
    data-availability cutoff time (via the `holidays` library).
  - Parsing a semi-structured, tab-delimited legacy export format with
    positional field detection.
  - Driving Excel via COM (win32com) to append rows to a live
    Excel Table (ListObject) on a cloud-hosted (SharePoint) workbook,
    letting Excel's own formula columns recalculate automatically.
  - Defensive design: idempotent file handling, guaranteed cleanup of
    the Excel process via try/finally, and a hard stop when the
    day's data isn't available yet instead of running against stale
    or partial data.

CONFIGURATION
-------------
All environment-specific values are read from environment variables
(see the CONFIGURATION section below). Copy `.env.example` to `.env`,
fill in your own values, and load them before running (e.g. with
`python-dotenv`), or export them in your shell.

REQUIREMENTS
------------
  pip install pywin32 holidays python-dotenv --break-system-packages

Requires an already-open, logged-in SAP GUI session with GUI Scripting
enabled (Options > Accessibility & Scripting > Scripting).
"""

import os
import re
import time
from datetime import date, timedelta, datetime

import holidays
import win32com.client as win32


# =============================================================================
# CONFIGURATION -- populate via environment variables, never hardcode
# real company data here.
# =============================================================================

COMPANY_CODE = os.environ.get("SAP_COMPANY_CODE", "XXXX")          # e.g. SAP "Sociedad"
BANK_ACCOUNT = os.environ.get("SAP_BANK_ACCOUNT", "00000000")       # G/L account number
MOVEMENT_CLASS = os.environ.get("SAP_MOVEMENT_CLASS", "XX")         # e.g. movement/document class filter
DATA_CUTOFF_HOUR = int(os.environ.get("DATA_CUTOFF_HOUR", "10"))    # hour before which source data isn't ready
HOLIDAY_COUNTRY = os.environ.get("HOLIDAY_COUNTRY", "US")           # ISO country code for the `holidays` lib

EXPORT_FOLDER = os.environ.get("EXPORT_FOLDER", r"C:\temp\sap_exports")

SAP_FAVORITE_NODE = os.environ.get("SAP_FAVORITE_NODE", "F00001")   # favorites-tree node id
SAP_LAYOUT_VARIANT = os.environ.get("SAP_LAYOUT_VARIANT", "/DEFAULT")

EXCEL_WORKBOOK_URL_TEMPLATE = os.environ.get(
    "EXCEL_WORKBOOK_URL_TEMPLATE",
    "https://YOUR_TENANT.sharepoint.com/personal/YOUR_USER/Documents/"
    "Reports/{year}/{month_folder}/Workbook_{month_name}_{year}.xlsx",
)
EXCEL_TABLE_NAME = os.environ.get("EXCEL_TABLE_NAME", "BankMovements")

# Column names expected in the exported file, in the exact order they
# appear (the parser aligns data to this list positionally).
SOURCE_FIELDS = [
    "Date", "GL_Company", "Doc_Number", "Account", "Currency",
    "Amount_Doc_Currency", "Amount_Local_Currency", "Amount_Local_Currency_2",
    "Text", "Movement_Class", "Company", "Customer", "Reference",
]
AMOUNT_FIELDS = {"Amount_Doc_Currency", "Amount_Local_Currency", "Amount_Local_Currency_2"}


# =============================================================================
# STEP 1: Determine which business date to query
# =============================================================================

def is_business_day(day: date, holiday_calendar: holidays.HolidayBase) -> bool:
    """A day is a business day if it's not a weekend and not a public holiday."""
    return day.weekday() < 5 and day not in holiday_calendar


def previous_business_day(reference_date: date) -> date:
    """Returns the closest business day strictly before `reference_date`."""
    calendar = holidays.country_holidays(HOLIDAY_COUNTRY, years=[reference_date.year - 1, reference_date.year])
    candidate = reference_date - timedelta(days=1)
    while not is_business_day(candidate, calendar):
        candidate -= timedelta(days=1)
    return candidate


def date_to_query(now: datetime = None) -> date | None:
    """
    Determines the date that should be queried in SAP.

    Returns None if it's still before DATA_CUTOFF_HOUR, since the
    upstream source system hasn't published the day's data yet -- in
    that case the process should not run.
    """
    if now is None:
        now = datetime.now()
    if now.hour < DATA_CUTOFF_HOUR:
        return None
    return previous_business_day(now.date())


# =============================================================================
# STEP 2: SAP navigation and filtering
# =============================================================================

def connect_to_sap():
    """Attaches to the first already-open SAP GUI session on this machine."""
    sap_gui_auto = win32.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def find_label_by_text(session, target_text: str, container: str = "wnd[0]/usr") -> str | None:
    """
    Searches the children of `container` for a label whose text matches
    `target_text` exactly, and returns its full technical ID.

    Used to locate column headers dynamically in classic SAP list
    screens, where on-screen position shifts depending on horizontal
    scroll -- more robust than hardcoding fixed row/column coordinates.
    """
    area = session.findById(container)
    for child in area.Children:
        try:
            text = child.Text.strip()
        except Exception:
            continue
        if text == target_text:
            return child.Id
    return None


def navigate_to_results(session, query_date: str) -> None:
    """Navigates from the favorites screen to the filtered movement list."""
    session.findById("wnd[0]").maximize()
    session.findById(
        "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"
    ).selectedNode = SAP_FAVORITE_NODE
    session.findById(
        "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell"
    ).doubleClickNode(SAP_FAVORITE_NODE)
    time.sleep(1)

    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = BANK_ACCOUNT
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = COMPANY_CODE
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = query_date
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = query_date
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = SAP_LAYOUT_VARIANT
    session.findById("wnd[0]").sendVKey(8)  # Execute (F8)
    time.sleep(1)

    # Horizontal scroll is required for the "movement class" column to
    # be rendered -- find_label_by_text can't find a control that isn't
    # currently visible on screen.
    session.findById("wnd[0]/usr").horizontalScrollbar.position = 41
    time.sleep(0.5)

    class_column_id = find_label_by_text(session, "Clase")  # column header, adjust to your SAP language
    if class_column_id is None:
        raise RuntimeError("Could not find the movement-class column on screen.")

    session.findById(class_column_id).setFocus()
    session.findById("wnd[0]").sendVKey(2)             # activate column filter
    session.findById("wnd[0]/tbar[1]/btn[38]").press()  # open filter editor
    time.sleep(0.5)
    session.findById(
        "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW"
    ).text = MOVEMENT_CLASS
    session.findById("wnd[1]").sendVKey(0)  # confirm filter
    time.sleep(1)


def start_export(session) -> None:
    """
    Opens System > List > Save > Local File, and selects the
    "unconverted text" format in the format dialog.

    After this step, window wnd[1] changes content to show the
    destination fields (Directory / Filename / Encoding), handled by
    save_export_file().
    """
    session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[2]").select()
    time.sleep(1)
    session.findById(
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
    ).select()
    session.findById(
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
    ).setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1.5)


# =============================================================================
# STEP 3: Export to local file
# =============================================================================

def save_export_file(session, destination_folder: str, filename: str) -> None:
    """
    Writes the destination folder and filename directly into SAP's own
    fields (ctxtDY_PATH / ctxtDY_FILENAME) and confirms the save.

    These are native SAP GUI Scripting fields -- no Windows-dialog
    automation (pywinauto/win32gui) is required. Writing both fields
    explicitly, rather than relying on whatever SAP remembers from a
    previous session, prevents the file from being saved to the wrong
    folder if another process used SAP more recently.
    """
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = destination_folder
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1.5)

    # The first time a new folder is accessed, SAP GUI may show a
    # security prompt asking to confirm file access. If the user
    # already checked "remember my decision" during a manual run, this
    # shouldn't reappear; the try/except covers the case where it does,
    # without breaking the flow if it doesn't.
    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()  # "Allow"
        time.sleep(1)
    except Exception:
        pass


# =============================================================================
# STEP 4: Parse the exported file
# =============================================================================

DATE_PATTERN = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


def clean_amount(text: str) -> float | str | None:
    """Converts '2,966.00' to 2966.0 (float). Non-numeric input is returned as-is."""
    text = text.strip()
    if text == "":
        return None
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return text


def parse_line(line: str, fields: list[str]) -> dict | None:
    """
    Extracts values from one line of the exported file, aligning them
    with `fields` in the order they appear.

    The file is tab-delimited, with blank cells represented by
    consecutive tabs (so splitting on tab preserves field position even
    when a value is empty). The start of the data is located by finding
    the first value shaped like a date (dd.mm.yyyy); lines with no date
    at all (headers, totals, etc.) are discarded by returning None.
    """
    parts = line.split("\t")

    date_index = next(
        (i for i, value in enumerate(parts) if DATE_PATTERN.match(value.strip())),
        None,
    )
    if date_index is None:
        return None

    row_values = parts[date_index:date_index + len(fields)]
    if len(row_values) < len(fields):
        return None

    row = {}
    for field_name, value in zip(fields, row_values):
        value = value.strip()
        row[field_name] = clean_amount(value) if field_name in AMOUNT_FIELDS else value
    return row


def parse_export_file(file_path: str, fields: list[str], encoding: str = "latin-1") -> list[dict]:
    """
    Reads the exported file line by line and returns the list of valid
    movement records.

    encoding="latin-1" because classic SAP exports aren't UTF-8, and
    contain special characters that would fail under the default codec.
    """
    movements = []
    with open(file_path, encoding=encoding) as f:
        for line in f:
            row = parse_line(line, fields)
            if row is not None:
                movements.append(row)
    return movements


# =============================================================================
# STEP 5: Write to the shared Excel workbook
# =============================================================================

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def build_workbook_url(reference_date: date) -> str:
    """Builds the workbook URL for the month/year of `reference_date`."""
    month_name = MONTH_NAMES[reference_date.month - 1]
    return EXCEL_WORKBOOK_URL_TEMPLATE.format(
        year=reference_date.year,
        month_folder=f"{reference_date.month:02d}",
        month_name=month_name,
    )


def append_to_excel_table(workbook_url: str, movements: list[dict], table_name: str) -> None:
    """
    Opens the workbook via COM, locates `table_name` and appends one row
    per movement, writing values by column name (not position, to be
    resilient to column reordering in the table).

    Formula-driven columns extend automatically when adding rows, since
    this is a real Excel Table (ListObject) and not a plain range.
    """
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(workbook_url)

        table = None
        for sheet in workbook.Sheets:
            for list_object in sheet.ListObjects:
                if list_object.Name == table_name:
                    table = list_object

        if table is None:
            raise RuntimeError(f"Table '{table_name}' not found in the workbook.")

        column_index = {col.Name: col.Index for col in table.ListColumns}

        for movement in movements:
            new_row = table.ListRows.Add()
            row_range = new_row.Range
            for field_name, value in movement.items():
                if field_name in column_index:
                    row_range.Cells(1, column_index[field_name]).Value = value

        workbook.Save()

    finally:
        # Guaranteed cleanup even if something fails mid-way, to avoid
        # leaving an orphaned EXCEL.EXE process running.
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


# =============================================================================
# MAIN ORCHESTRATION
# =============================================================================

def main() -> None:
    print(f"Current time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    query_date_obj = date_to_query()
    if query_date_obj is None:
        print(f"Before {DATA_CUTOFF_HOUR}:00 -- source data not published yet. Aborting.")
        return

    query_date_str = query_date_obj.strftime("%d.%m.%Y")
    print(f"Querying SAP for: {query_date_str}")

    session = connect_to_sap()
    navigate_to_results(session, query_date_str)
    start_export(session)

    export_filename = f"export_{query_date_obj.strftime('%d-%m-%Y')}.txt"
    export_path = os.path.join(EXPORT_FOLDER, export_filename)

    # Remove any leftover file from a same-day rerun, to avoid SAP's
    # overwrite-confirmation dialog.
    try:
        os.remove(export_path)
    except FileNotFoundError:
        pass

    save_export_file(session, EXPORT_FOLDER, export_filename)

    movements = parse_export_file(export_path, fields=SOURCE_FIELDS)
    print(f"Movements found: {len(movements)}")

    if not movements:
        print("Nothing to add.")
        return

    workbook_url = build_workbook_url(query_date_obj)
    append_to_excel_table(workbook_url, movements, table_name=EXCEL_TABLE_NAME)

    print(f"Done: {len(movements)} movements appended to table '{EXCEL_TABLE_NAME}'.")


if __name__ == "__main__":
    main()
