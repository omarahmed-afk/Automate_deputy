import json
import re
import time
import requests
import pandas as pd
import gspread

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials

from dotenv import load_dotenv
import os
from pathlib import Path


# ================= CONFIG =================

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

TOKEN = os.getenv("DEPUTY_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
TARGET_DATE = os.getenv("TARGET_DATE", "").strip().strip('"').strip("'") or None

if not TOKEN:
    raise ValueError("DEPUTY_TOKEN is missing. Check your .env file.")

if not SPREADSHEET_ID:
    raise ValueError("SPREADSHEET_ID is missing. Check your .env file.")

TOKEN = TOKEN.strip().strip('"').strip("'")
SPREADSHEET_ID = SPREADSHEET_ID.strip().strip('"').strip("'")

NY = ZoneInfo("America/New_York")

INSTALL_URL = "https://ptofthecity.na.deputy.com"
API_BASE = f"{INSTALL_URL}/api/v1"

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

GOOGLE_CREDENTIALS_FILE = os.getenv(
    "GOOGLE_CREDENTIALS_FILE",
    str(BASE_DIR / "service_account.json")
)

SHEET_NAME = "Sheet1"

# A = Date
DATE_COLUMN_NUMBER = 1

# D = Location / Clinic names
CLINIC_COLUMN_NUMBER = 4

# Daily block rows
FIRST_DATA_ROW = 2
DAILY_BLOCK_ROWS = 32
LAST_DATA_ROW = FIRST_DATA_ROW + DAILY_BLOCK_ROWS - 1

# automate_dpt_2 output columns
# V = PT
# W = Assistant
# X = PCC
OUTPUT_START_COL = "V"
OUTPUT_END_COL = "X"

# Copy template from A:X
# A=0, X=23, end index is exclusive = 24
TEMPLATE_END_COLUMN_INDEX = 24


# ================= DATE FUNCTIONS =================

def get_target_date():
    """
    If TARGET_DATE exists, use it.
    Otherwise use yesterday based on New York time.
    """

    if TARGET_DATE:
        datetime.strptime(TARGET_DATE, "%Y-%m-%d")
        return TARGET_DATE

    yesterday = datetime.now(NY).date() - timedelta(days=1)
    return str(yesterday)


def get_day_range_unix(target_date):
    start_dt = datetime.strptime(target_date, "%Y-%m-%d").replace(tzinfo=NY)
    end_dt = start_dt + timedelta(days=1)

    return int(start_dt.timestamp()), int(end_dt.timestamp())


def format_sheet_date(target_date):
    """
    Convert 2026-06-29 to 6/29/2026.
    """

    date_obj = datetime.strptime(target_date, "%Y-%m-%d").date()
    return f"{date_obj.month}/{date_obj.day}/{date_obj.year}"


def normalize_sheet_date(value):
    """
    Return normalized m/d/yyyy only if value is a real date.
    Otherwise return empty string.

    Examples:
    2026-06-29 -> 6/29/2026
    6/29/2026  -> 6/29/2026
    """

    if value is None:
        return ""

    s = str(value).strip()

    if not s:
        return ""

    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            d = datetime.strptime(s, fmt).date()
            return f"{d.month}/{d.day}/{d.year}"
        except ValueError:
            pass

    return ""


# ================= DEPUTY API =================

def deputy_query(resource_name, payload):
    url = f"{API_BASE}/resource/{resource_name}/QUERY"

    response = requests.post(
        url,
        headers=HEADERS,
        json=payload,
        timeout=60
    )

    print(resource_name, "status:", response.status_code)

    if response.status_code != 200:
        print(response.text[:1000])
        response.raise_for_status()

    data = response.json()
    return pd.json_normalize(data)


# ================= CLEANING FUNCTIONS =================

def clean_text(value):
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)

    return text


def map_role(role):
    role_clean = clean_text(role)

    if role_clean == "":
        return "Other"

    if "physical therapist" in role_clean or "physical threapist" in role_clean:
        return "PT"

    if (
        "physical therapy assistant" in role_clean
        or role_clean == "pta"
        or "aide" in role_clean
    ):
        return "Assistant"

    if (
        "patient care coordinator" in role_clean
        or "patients care coordinator" in role_clean
        or role_clean == "pcc"
        or "(pcc)" in role_clean
    ):
        return "PCC"

    return "Other"


def clean_time(value):
    if value is None:
        return ""

    if value != value:
        return ""

    s = str(value).strip()

    if s == "" or s.lower() == "nan":
        return ""

    if "T" in s:
        time_part = s.split("T")[1][:5]
    elif " " in s:
        parts = s.split()
        time_part = parts[1][:5]
    else:
        time_part = s[:5]

    hour, minute = time_part.split(":")
    hour = int(hour)
    minute = int(minute)

    am_pm = "am" if hour < 12 else "pm"

    hour_12 = hour % 12
    if hour_12 == 0:
        hour_12 = 12

    if minute == 0:
        return f"{hour_12}{am_pm}"

    return f"{hour_12}:{minute:02d}{am_pm}"


# ================= BUILD DATAFRAME =================

def build_schedule_df(roster_df):
    needed_api_columns = [
        "Id",
        "Date",
        "StartTime",
        "EndTime",
        "StartTimeLocalized",
        "EndTimeLocalized",
        "TotalTime",
        "_DPMetaData.EmployeeInfo.DisplayName",
        "_DPMetaData.OperationalUnitInfo.OperationalUnitName",
        "_DPMetaData.OperationalUnitInfo.CompanyName",
        "MatchedByTimesheet"
    ]

    missing_cols = [col for col in needed_api_columns if col not in roster_df.columns]

    if missing_cols:
        raise ValueError(f"Missing columns from Deputy response: {missing_cols}")

    schedule_df = roster_df[needed_api_columns].copy()

    schedule_df = schedule_df.rename(columns={
        "Id": "Roster_Id",
        "StartTime": "Start_Unix",
        "EndTime": "End_Unix",
        "StartTimeLocalized": "Scheduled_Start",
        "EndTimeLocalized": "Scheduled_End",
        "TotalTime": "Deputy_TotalTime",
        "_DPMetaData.EmployeeInfo.DisplayName": "Employee_Name",
        "_DPMetaData.OperationalUnitInfo.OperationalUnitName": "Role_Name",
        "_DPMetaData.OperationalUnitInfo.CompanyName": "Clinic_Name",
        "MatchedByTimesheet": "Timesheet_Id"
    })

    schedule_df["Employee_Name"] = (
        schedule_df["Employee_Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    schedule_df = schedule_df[
        (schedule_df["Employee_Name"] != "") &
        (schedule_df["Employee_Name"].str.upper() != "EMPTY") &
        (schedule_df["Employee_Name"].str.lower() != "nan")
    ].copy()

    schedule_df["Role_Mapped"] = schedule_df["Role_Name"].apply(map_role)

    schedule_df["Start_Clean"] = schedule_df["Scheduled_Start"].apply(clean_time)
    schedule_df["End_Clean"] = schedule_df["Scheduled_End"].apply(clean_time)

    schedule_df["Shift_Time"] = (
        schedule_df["Start_Clean"] + " – " + schedule_df["End_Clean"]
    )

    schedule_df["Start_DT_NY"] = (
        pd.to_datetime(schedule_df["Start_Unix"], unit="s", utc=True)
        .dt.tz_convert(NY)
    )

    schedule_df["End_DT_NY"] = (
        pd.to_datetime(schedule_df["End_Unix"], unit="s", utc=True)
        .dt.tz_convert(NY)
    )

    schedule_df["Date_Clean"] = schedule_df["Start_DT_NY"].dt.strftime("%Y-%m-%d")

    schedule_df["Total_Hours"] = (
        pd.to_numeric(schedule_df["End_Unix"], errors="coerce") -
        pd.to_numeric(schedule_df["Start_Unix"], errors="coerce")
    ) / 3600

    schedule_df["Total_Hours"] = (
        schedule_df["Total_Hours"]
        .fillna(0)
        .round(2)
    )

    final_schedule = schedule_df[
        [
            "Date_Clean",
            "Clinic_Name",
            "Role_Name",
            "Role_Mapped",
            "Employee_Name",
            "Shift_Time",
            "Total_Hours",
            "Roster_Id",
            "Timesheet_Id"
        ]
    ].copy()

    final_schedule = final_schedule.rename(columns={
        "Date_Clean": "Date"
    })

    return final_schedule


# ================= GOOGLE SHEETS =================

def open_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    google_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

    if google_json:
        service_account_info = json.loads(google_json)
        creds = Credentials.from_service_account_info(
            service_account_info,
            scopes=scopes
        )
    else:
        with open(GOOGLE_CREDENTIALS_FILE, "r") as f:
            service_account_info = json.load(f)

        creds = Credentials.from_service_account_file(
            GOOGLE_CREDENTIALS_FILE,
            scopes=scopes
        )

    print("Service account email:")
    print(service_account_info["client_email"])

    gc = gspread.authorize(creds)

    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    print("\nSpreadsheet title:", spreadsheet.title)
    print("Available worksheets:")
    for ws in spreadsheet.worksheets():
        print("-", ws.title)

    return spreadsheet


def build_clinic_role_totals(final_schedule):
    df = final_schedule.copy()

    df["Total_Hours"] = pd.to_numeric(
        df["Total_Hours"],
        errors="coerce"
    ).fillna(0)

    clinic_totals = (
        df.groupby(["Clinic_Name", "Role_Mapped"], as_index=False)["Total_Hours"]
        .sum()
    )

    pivot = (
        clinic_totals
        .pivot_table(
            index="Clinic_Name",
            columns="Role_Mapped",
            values="Total_Hours",
            fill_value=0,
            aggfunc="sum"
        )
        .reset_index()
    )

    pivot.columns.name = None

    for col in ["PT", "Assistant", "PCC"]:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot["PT"] = pd.to_numeric(
        pivot["PT"],
        errors="coerce"
    ).fillna(0).round(2)

    pivot["Assistant"] = pd.to_numeric(
        pivot["Assistant"],
        errors="coerce"
    ).fillna(0).round(2)

    pivot["PCC"] = pd.to_numeric(
        pivot["PCC"],
        errors="coerce"
    ).fillna(0).round(2)

    pivot["Clinic_Key"] = pivot["Clinic_Name"].apply(clean_text)

    return pivot[["Clinic_Name", "Clinic_Key", "PT", "Assistant", "PCC"]]


# ================= DATE-BASED UPDATE LOGIC =================

def find_existing_date_start_row(ws, sheet_date):
    """
    Search only column A.
    If date exists, return the row number.
    If not, return None.
    """

    target = normalize_sheet_date(sheet_date)
    col_a_values = ws.col_values(DATE_COLUMN_NUMBER)

    for row_number, value in enumerate(col_a_values, start=1):
        if normalize_sheet_date(value) == target:
            return row_number

    return None


def get_next_daily_block_start_row_by_date(ws):
    """
    Find next block based only on dates in column A.

    Example:
    A2  = 6/28/2026
    A34 = 6/29/2026
    next = 66
    """

    col_a_values = ws.col_values(DATE_COLUMN_NUMBER)

    date_rows = []

    for row_number, value in enumerate(col_a_values, start=1):
        if normalize_sheet_date(value):
            date_rows.append(row_number)

    if not date_rows:
        return FIRST_DATA_ROW

    last_date_row = max(date_rows)
    return last_date_row + DAILY_BLOCK_ROWS


def append_or_update_daily_report_by_date(spreadsheet, final_schedule):
    """
    Main logic for automate_dpt_2.

    1. Get target date.
    2. Search for that date in column A.
    3. If date exists:
        - Fill V:X for that same date block.
        - Do NOT update/fill date.
    4. If date does not exist:
        - Copy template block below last date.
        - Write date once in A{start_row}.
        - Fill V:X.
    """

    ws = spreadsheet.worksheet(SHEET_NAME)
    sheet_id = ws.id

    pivot = build_clinic_role_totals(final_schedule)

    target_date = get_target_date()
    sheet_date = format_sheet_date(target_date)

    existing_start_row = find_existing_date_start_row(ws, sheet_date)

    if existing_start_row:
        start_row = existing_start_row
        end_row = start_row + DAILY_BLOCK_ROWS - 1
        should_write_date = False

        print(f"Date {sheet_date} exists in column A at row {start_row}.")
        print(f"Filling existing block rows {start_row}:{end_row}.")
    else:
        start_row = get_next_daily_block_start_row_by_date(ws)
        end_row = start_row + DAILY_BLOCK_ROWS - 1
        should_write_date = True

        print(f"Date {sheet_date} not found in column A.")
        print(f"Appending new block rows {start_row}:{end_row}.")

        source_start_index = FIRST_DATA_ROW - 1
        source_end_index = LAST_DATA_ROW

        destination_start_index = start_row - 1
        destination_end_index = destination_start_index + DAILY_BLOCK_ROWS

        spreadsheet.batch_update({
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": destination_start_index,
                            "endIndex": destination_end_index
                        },
                        "inheritFromBefore": False
                    }
                },
                {
                    "copyPaste": {
                        "source": {
                            "sheetId": sheet_id,
                            "startRowIndex": source_start_index,
                            "endRowIndex": source_end_index,
                            "startColumnIndex": 0,
                            "endColumnIndex": TEMPLATE_END_COLUMN_INDEX
                        },
                        "destination": {
                            "sheetId": sheet_id,
                            "startRowIndex": destination_start_index,
                            "endRowIndex": destination_end_index,
                            "startColumnIndex": 0,
                            "endColumnIndex": TEMPLATE_END_COLUMN_INDEX
                        },
                        "pasteType": "PASTE_NORMAL",
                        "pasteOrientation": "NORMAL"
                    }
                }
            ]
        })

        time.sleep(2)

    # Read locations from the target block
    all_clinic_names = ws.get(f"D{start_row}:D{end_row}")
    all_clinic_names = [row[0] if row else "" for row in all_clinic_names]

    print(f"\nClinics from D{start_row}:D{end_row}:")
    print(all_clinic_names[:15])

    output_values = []
    unmatched = []

    for clinic_name in all_clinic_names:
        clinic_key = clean_text(clinic_name)

        if clinic_key == "":
            output_values.append(["", "", ""])
            continue

        matched = pivot[pivot["Clinic_Key"] == clinic_key]

        if matched.empty:
            output_values.append([0, 0, 0])
            unmatched.append(clinic_name)
        else:
            output_values.append([
                float(matched["PT"].iloc[0]),
                float(matched["Assistant"].iloc[0]),
                float(matched["PCC"].iloc[0])
            ])

    # Write date only if it is a newly appended block
    if should_write_date:
        ws.update(
            range_name=f"A{start_row}",
            values=[[sheet_date]]
        )

    # Fill hours only in this date block
    ws.batch_clear([f"{OUTPUT_START_COL}{start_row}:{OUTPUT_END_COL}{end_row}"])

    ws.update(
        range_name=f"{OUTPUT_START_COL}{start_row}:{OUTPUT_END_COL}{end_row}",
        values=output_values
    )

    print(f"\nReport {sheet_date} filled in rows {start_row}:{end_row}.")
    print(f"Hours filled: {OUTPUT_START_COL}{start_row}:{OUTPUT_END_COL}{end_row}")

    if unmatched:
        print("\nUnmatched clinics:")
        for clinic in unmatched:
            print("-", clinic)

        print("\nDeputy clinic names available:")
        for clinic in pivot["Clinic_Name"].tolist():
            print("-", clinic)


# ================= MAIN REPORT =================

def run_report():
    target_date = get_target_date()

    print("Running report for:", target_date)

    start_unix, end_unix = get_day_range_unix(target_date)

    roster_payload = {
        "search": {
            "s1": {
                "field": "StartTime",
                "data": start_unix,
                "type": "ge"
            },
            "s2": {
                "field": "StartTime",
                "data": end_unix,
                "type": "lt"
            }
        }
    }

    roster_df = deputy_query("Roster", roster_payload)

    if roster_df.empty:
        print("No roster data found.")
        return

    final_schedule = build_schedule_df(roster_df)

    print("\nRole check:")
    print(
        final_schedule
        .groupby(["Role_Name", "Role_Mapped"], as_index=False)["Total_Hours"]
        .sum()
        .sort_values(["Role_Mapped", "Role_Name"])
    )

    spreadsheet = open_google_sheet()

    append_or_update_daily_report_by_date(
        spreadsheet=spreadsheet,
        final_schedule=final_schedule
    )
    print("\nDone. Sheet2 updated by date.")


if __name__ == "__main__":
    run_report()