import json
import re
import requests
import pandas as pd
import gspread
import argparse
import logging
import sys

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

GOOGLE_CREDENTIALS_FILE = BASE_DIR / "service_account.json"

# Google Sheet tab name
SHEET_NAME = "Sheet1"

# Daily report structure
DAILY_BLOCK_START_ROW = 2
DAILY_BLOCK_ROWS = 32

# Column layout:
# A = Date
# D = Location
# V = PT Hours
# W = Assistant Hours
# X = PCC Hours

# Python list indexes are zero-based
DATE_COL_INDEX = 0       # A
LOCATION_COL_INDEX = 3   # D
PT_COL_INDEX = 21        # V
ASSISTANT_COL_INDEX = 22 # W
PCC_COL_INDEX = 23       # X

# Template range A:X = 24 columns
TEMPLATE_LAST_COL = "X"
TEMPLATE_COL_COUNT = 24


# ================= LOGGING =================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ],
)

log = logging.getLogger(__name__)


# ================= DATE FUNCTIONS =================

def get_target_date(override: str | None = None) -> str:
    """
    Return target date as YYYY-MM-DD.
    Default = yesterday based on New York time.
    """
    if override:
        datetime.strptime(override, "%Y-%m-%d")
        return override

    yesterday = datetime.now(NY).date() - timedelta(days=1)
    return str(yesterday)


def get_day_range_unix(target_date: str):
    start_dt = datetime.strptime(target_date, "%Y-%m-%d").replace(tzinfo=NY)
    end_dt = start_dt + timedelta(days=1)

    return int(start_dt.timestamp()), int(end_dt.timestamp())


def format_sheet_date(target_date: str) -> str:
    """
    Convert 2026-06-27 to 6/27 for Google Sheet display.
    """
    date_obj = datetime.strptime(target_date, "%Y-%m-%d").date()
    return f"{date_obj.month}/{date_obj.day}"


# ================= DEPUTY API =================

def deputy_query(resource_name: str, payload: dict) -> pd.DataFrame:
    url = f"{API_BASE}/resource/{resource_name}/QUERY"

    response = requests.post(
        url,
        headers=HEADERS,
        json=payload,
        timeout=60
    )

    log.info(f"{resource_name} status: {response.status_code}")

    if response.status_code != 200:
        log.error(response.text[:1000])
        response.raise_for_status()

    return pd.json_normalize(response.json())


# ================= CLEANING FUNCTIONS =================

def clean_text(value) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)

    return text


def map_role(role: str) -> str:
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


def clean_time(value) -> str:
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

def build_schedule_df(roster_df: pd.DataFrame) -> pd.DataFrame:
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
        "MatchedByTimesheet",
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
        "MatchedByTimesheet": "Timesheet_Id",
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
            "Timesheet_Id",
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
        "https://www.googleapis.com/auth/drive",
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

    log.info(f"Service account: {service_account_info['client_email']}")

    gc = gspread.authorize(creds)

    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    log.info(f"Spreadsheet: {spreadsheet.title}")

    return spreadsheet


# ================= PIVOT HELPERS =================

def build_clinic_role_totals(final_schedule: pd.DataFrame) -> pd.DataFrame:
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


def match_hours(pivot: pd.DataFrame, clinic_key: str):
    matched = pivot[pivot["Clinic_Key"] == clinic_key]

    if matched.empty:
        return 0.0, 0.0, 0.0

    row = matched.iloc[0]

    return (
        float(row["PT"]),
        float(row["Assistant"]),
        float(row["PCC"]),
    )


# ================= DAILY REPORT INSERT =================

def insert_daily_report_on_top(
    spreadsheet,
    final_schedule: pd.DataFrame,
    target_date: str
):
    """
    Insert a new 32-row daily report block on top.
    Old reports move down automatically.

    This preserves the design by:
    1. inserting blank rows,
    2. copying the previous top block formatting/content,
    3. updating only Date and PT/Assistant/PCC hours.

    Layout:
    A = Date
    D = Location
    V = PT Hours
    W = Assistant Hours
    X = PCC Hours
    """

    ws = spreadsheet.worksheet(SHEET_NAME)
    sheet_id = ws.id

    pivot = build_clinic_role_totals(final_schedule)

    log.info("\nClinic totals from Deputy:")
    log.info(pivot[["Clinic_Name", "PT", "Assistant", "PCC"]].to_string(index=False))

    start_row = DAILY_BLOCK_START_ROW
    block_rows = DAILY_BLOCK_ROWS
    end_row = start_row + block_rows - 1

    log.info(f"Using fixed daily block: rows {start_row}-{end_row} ({block_rows} clinics)")

    # Read the current top block before inserting rows.
    template_rows = ws.get(f"A{start_row}:{TEMPLATE_LAST_COL}{end_row}")

    start_index = start_row - 1
    end_index = start_index + block_rows

    # Insert new rows and copy the old top block into them to preserve design.
    spreadsheet.batch_update({
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": end_index,
                    },
                    "inheritFromBefore": False,
                }
            },
            {
                "copyPaste": {
                    "source": {
                        "sheetId": sheet_id,
                        "startRowIndex": end_index,
                        "endRowIndex": end_index + block_rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": TEMPLATE_COL_COUNT,
                    },
                    "destination": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_index,
                        "endRowIndex": end_index,
                        "startColumnIndex": 0,
                        "endColumnIndex": TEMPLATE_COL_COUNT,
                    },
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL",
                }
            },
        ]
    })

    sheet_date_text = format_sheet_date(target_date)

    date_values = []
    output_values = []
    unmatched = []

    for row in template_rows:
        row = list(row) + [""] * (TEMPLATE_COL_COUNT - len(row))

        location_name = str(row[LOCATION_COL_INDEX]).strip()
        location_key = clean_text(location_name)

        if location_key:
            date_values.append([sheet_date_text])
        else:
            date_values.append([""])

        pt, assistant, pcc = match_hours(pivot, location_key)

        if location_key and pt == 0 and assistant == 0 and pcc == 0:
            unmatched.append(location_name)

        if location_key:
            output_values.append([pt, assistant, pcc])
        else:
            output_values.append(["", "", ""])

    new_end_row = start_row + block_rows - 1

    # Update only the new block.
    ws.update(
        range_name=f"A{start_row}:A{new_end_row}",
        values=date_values
    )

    ws.update(
        range_name=f"V{start_row}:X{new_end_row}",
        values=output_values
    )

    log.info(f"Inserted new daily report for {sheet_date_text} at row {start_row}.")

    if unmatched:
        log.warning("\nClinics in Sheet1 with no Deputy data today:")
        for clinic in unmatched:
            log.warning(f"- {clinic}")

        log.info("\nDeputy clinic names available:")
        for clinic in pivot["Clinic_Name"].tolist():
            log.info(f"- {clinic}")


# ================= MAIN REPORT =================

def run_report(date_override: str | None = None):
    target_date = get_target_date(date_override)

    log.info(f"Running report for: {target_date}")

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
        log.warning("No roster data found.")
        return

    final_schedule = build_schedule_df(roster_df)

    log.info("\nRole summary:")
    log.info(
        final_schedule
        .groupby(["Role_Name", "Role_Mapped"], as_index=False)["Total_Hours"]
        .sum()
        .sort_values(["Role_Mapped", "Role_Name"])
        .to_string(index=False)
    )

    spreadsheet = open_google_sheet()

    insert_daily_report_on_top(
        spreadsheet=spreadsheet,
        final_schedule=final_schedule,
        target_date=target_date
    )

    log.info("Done. New daily report inserted on top.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Pull Deputy roster and insert daily report into Google Sheets."
    )

    parser.add_argument(
        "--date",
        metavar="YYYY-MM-DD",
        default=None,
        help="Optional target date. Default = yesterday in NY time."
    )

    args = parser.parse_args()

    run_report(date_override=args.date)