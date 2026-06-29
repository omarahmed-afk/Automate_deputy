import json
import re
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

# Google Sheets
GOOGLE_CREDENTIALS_FILE = Path(
    os.getenv("GOOGLE_CREDENTIALS_FILE", str(BASE_DIR / "service_account.json"))
)

SHEET_NAME = "Sheet1"

# Daily report layout
DAILY_BLOCK_START_ROW = 2

# Column indexes in Python are zero-based
DATE_COL_INDEX = 0          # A
PT_COL_INDEX = 21           # V
ASSISTANT_COL_INDEX = 23    # X
PCC_COL_INDEX = 24          # Y
LOCATION_COL_INDEX = 29     # AD

TEMPLATE_LAST_COL = "AD"
TEMPLATE_COL_COUNT = 30     # A:AD


# ================= DATE FUNCTIONS =================

def get_target_date():
    """
    Use yesterday based on New York time.
    Example: if today is 2026-06-28, target date = 2026-06-27.
    """
    yesterday = datetime.now(NY).date() - timedelta(days=1)
    return str(yesterday)


def get_day_range_unix(target_date):
    start_dt = datetime.strptime(target_date, "%Y-%m-%d").replace(tzinfo=NY)
    end_dt = start_dt + timedelta(days=1)

    return int(start_dt.timestamp()), int(end_dt.timestamp())


def format_sheet_date(target_date):
    """
    Convert 2026-06-27 to 6/27 for Google Sheet display.
    """
    date_obj = datetime.strptime(target_date, "%Y-%m-%d").date()
    return f"{date_obj.month}/{date_obj.day}"


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


# ================= DAILY INSERT LOGIC =================

def get_current_top_block_row_count(ws):
    """
    Count rows in the current top daily block dynamically.

    Expected layout:
    A  = Date
    AD = Location

    It stops when:
    - the date in column A changes, or
    - it reaches an empty area after data.
    """

    start_row = DAILY_BLOCK_START_ROW
    rows = ws.get(f"A{start_row}:{TEMPLATE_LAST_COL}1000")

    if not rows:
        raise ValueError("No rows found in the report sheet.")

    first_date = None
    count = 0

    for row in rows:
        row = row + [""] * (TEMPLATE_COL_COUNT - len(row))

        date_value = str(row[DATE_COL_INDEX]).strip()
        location_value = str(row[LOCATION_COL_INDEX]).strip()

        # Skip empty rows before the first data row
        if not date_value and not location_value and count == 0:
            continue

        # Detect the first date of the top block
        if first_date is None and date_value:
            first_date = date_value

        # Stop when the next date block starts
        if first_date and date_value and date_value != first_date:
            break

        # Stop when empty area starts after data
        if not date_value and not location_value and count > 0:
            break

        count += 1

    if count == 0:
        raise ValueError(
            "Could not detect daily block rows. "
            "Check Date in column A and Location in column AD."
        )

    return count


def insert_daily_report_on_top(spreadsheet, final_schedule):
    """
    Insert a new daily report block at the top.
    Old daily reports move down automatically.

    Expected layout:
    A  = Date
    AD = Location
    V  = PT Hours
    X  = Assistant Hours
    Y  = PCC Hours
    """

    ws = spreadsheet.worksheet(SHEET_NAME)

    target_date = get_target_date()
    sheet_date_text = format_sheet_date(target_date)

    pivot = build_clinic_role_totals(final_schedule)

    print("\nClinic totals from Deputy:")
    print(pivot[["Clinic_Name", "PT", "Assistant", "PCC"]])

    start_row = DAILY_BLOCK_START_ROW
    daily_block_rows = get_current_top_block_row_count(ws)
    end_row = start_row + daily_block_rows - 1

    print(f"\nDetected daily block rows: {daily_block_rows}")

    template_rows = ws.get(f"A{start_row}:{TEMPLATE_LAST_COL}{end_row}")

    new_rows = []
    unmatched = []

    for row in template_rows:
        row = row + [""] * (TEMPLATE_COL_COUNT - len(row))

        location_name = row[LOCATION_COL_INDEX]
        location_key = clean_text(location_name)

        if location_key:
            row[DATE_COL_INDEX] = sheet_date_text
        else:
            row[DATE_COL_INDEX] = ""

        matched = pivot[pivot["Clinic_Key"] == location_key]

        if location_key == "":
            row[PT_COL_INDEX] = ""
            row[ASSISTANT_COL_INDEX] = ""
            row[PCC_COL_INDEX] = ""
        elif matched.empty:
            row[PT_COL_INDEX] = 0
            row[ASSISTANT_COL_INDEX] = 0
            row[PCC_COL_INDEX] = 0
            unmatched.append(location_name)
        else:
            row[PT_COL_INDEX] = float(matched["PT"].iloc[0])
            row[ASSISTANT_COL_INDEX] = float(matched["Assistant"].iloc[0])
            row[PCC_COL_INDEX] = float(matched["PCC"].iloc[0])

        new_rows.append(row[:TEMPLATE_COL_COUNT])

    ws.insert_rows(
        new_rows,
        row=start_row,
        value_input_option="USER_ENTERED"
    )

    print(f"\nInserted new daily report for {sheet_date_text} at row {start_row}.")

    if unmatched:
        print("\nUnmatched locations from Sheet:")
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

    insert_daily_report_on_top(
        spreadsheet=spreadsheet,
        final_schedule=final_schedule
    )

    print("\nDone. New daily report inserted on top.")


if __name__ == "__main__":
    run_report()