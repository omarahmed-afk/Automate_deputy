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


# ================= CONFIG =================
from dotenv import load_dotenv
import os
from pathlib import Path

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

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

NY = ZoneInfo("America/New_York")

INSTALL_URL = "https://ptofthecity.na.deputy.com"
API_BASE = f"{INSTALL_URL}/api/v1"

GOOGLE_CREDENTIALS_FILE = BASE_DIR / "service_account.json"

SHEET_NAME = "Sheet1"

# Column D = index 4 (1-based for gspread)
CLINIC_COLUMN_NUMBER = 4

# Data starts at row 2 (row 1 = headers)
FIRST_DATA_ROW = 2
DAILY_BLOCK_START_ROW = 2

# Output columns (1-based letter reference):
# V = PT Hours      → col index 22 → row list index 21
# W = Assistant     → col index 23 → row list index 22
# X = PCC           → col index 24 → row list index 23
COL_PT        = 21   # V (0-based row index)
COL_ASSISTANT = 22   # W
COL_PCC       = 23   # X
RANGE_OUTPUT  = "V"  # Start column letter for batch write

# ================= LOGGING =================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(BASE_DIR / "deputy_report.log"),
    ],
)
log = logging.getLogger(__name__)


# ================= DATE FUNCTIONS =================

def get_target_date(override: str | None = None) -> str:
    """Return ISO date string: override if provided, else yesterday NY time."""
    if override:
        # Validate format
        datetime.strptime(override, "%Y-%m-%d")
        return override
    yesterday = datetime.now(NY).date() - timedelta(days=1)
    return str(yesterday)


def get_day_range_unix(target_date: str):
    start_dt = datetime.strptime(target_date, "%Y-%m-%d").replace(tzinfo=NY)
    end_dt = start_dt + timedelta(days=1)
    return int(start_dt.timestamp()), int(end_dt.timestamp())


# ================= DEPUTY API =================

def deputy_query(resource_name: str, payload: dict) -> pd.DataFrame:
    url = f"{API_BASE}/resource/{resource_name}/QUERY"
    response = requests.post(url, headers=HEADERS, json=payload, timeout=60)
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
    if value is None or value != value:
        return ""
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if "T" in s:
        time_part = s.split("T")[1][:5]
    elif " " in s:
        time_part = s.split()[1][:5]
    else:
        time_part = s[:5]
    hour, minute = time_part.split(":")
    hour, minute = int(hour), int(minute)
    am_pm = "am" if hour < 12 else "pm"
    hour_12 = hour % 12 or 12
    return f"{hour_12}{am_pm}" if minute == 0 else f"{hour_12}:{minute:02d}{am_pm}"


# ================= BUILD DATAFRAME =================

def build_schedule_df(roster_df: pd.DataFrame) -> pd.DataFrame:
    needed = [
        "Id", "Date", "StartTime", "EndTime",
        "StartTimeLocalized", "EndTimeLocalized", "TotalTime",
        "_DPMetaData.EmployeeInfo.DisplayName",
        "_DPMetaData.OperationalUnitInfo.OperationalUnitName",
        "_DPMetaData.OperationalUnitInfo.CompanyName",
        "MatchedByTimesheet",
    ]
    missing = [c for c in needed if c not in roster_df.columns]
    if missing:
        raise ValueError(f"Missing columns from Deputy response: {missing}")

    df = roster_df[needed].copy().rename(columns={
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

    df["Employee_Name"] = df["Employee_Name"].fillna("").astype(str).str.strip()
    df = df[
        (df["Employee_Name"] != "") &
        (df["Employee_Name"].str.upper() != "EMPTY") &
        (df["Employee_Name"].str.lower() != "nan")
    ].copy()

    df["Role_Mapped"] = df["Role_Name"].apply(map_role)
    df["Start_Clean"] = df["Scheduled_Start"].apply(clean_time)
    df["End_Clean"] = df["Scheduled_End"].apply(clean_time)
    df["Shift_Time"] = df["Start_Clean"] + " – " + df["End_Clean"]

    df["Start_DT_NY"] = (
        pd.to_datetime(df["Start_Unix"], unit="s", utc=True).dt.tz_convert(NY)
    )
    df["End_DT_NY"] = (
        pd.to_datetime(df["End_Unix"], unit="s", utc=True).dt.tz_convert(NY)
    )
    df["Date_Clean"] = df["Start_DT_NY"].dt.strftime("%Y-%m-%d")

    df["Total_Hours"] = (
        (pd.to_numeric(df["End_Unix"], errors="coerce") -
         pd.to_numeric(df["Start_Unix"], errors="coerce")) / 3600
    ).fillna(0).round(2)

    return df[[
        "Date_Clean", "Clinic_Name", "Role_Name", "Role_Mapped",
        "Employee_Name", "Shift_Time", "Total_Hours",
        "Roster_Id", "Timesheet_Id",
    ]].rename(columns={"Date_Clean": "Date"})


# ================= GOOGLE SHEETS =================

def open_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    google_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if google_json:
        info = json.loads(google_json)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        with open(GOOGLE_CREDENTIALS_FILE) as f:
            info = json.load(f)
        creds = Credentials.from_service_account_file(
            str(GOOGLE_CREDENTIALS_FILE), scopes=scopes
        )

    log.info(f"Service account: {info['client_email']}")
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    log.info(f"Spreadsheet: {spreadsheet.title}")
    return spreadsheet


# ================= PIVOT HELPERS =================

def build_clinic_role_totals(final_schedule: pd.DataFrame) -> pd.DataFrame:
    df = final_schedule.copy()
    df["Total_Hours"] = pd.to_numeric(df["Total_Hours"], errors="coerce").fillna(0)

    pivot = (
        df.groupby(["Clinic_Name", "Role_Mapped"])["Total_Hours"]
        .sum()
        .reset_index()
        .pivot_table(
            index="Clinic_Name",
            columns="Role_Mapped",
            values="Total_Hours",
            fill_value=0,
            aggfunc="sum",
        )
        .reset_index()
    )
    pivot.columns.name = None

    for col in ["PT", "Assistant", "PCC"]:
        if col not in pivot.columns:
            pivot[col] = 0.0

    pivot["PT"]        = pd.to_numeric(pivot["PT"],        errors="coerce").fillna(0).round(2)
    pivot["Assistant"] = pd.to_numeric(pivot["Assistant"], errors="coerce").fillna(0).round(2)
    pivot["PCC"]       = pd.to_numeric(pivot["PCC"],       errors="coerce").fillna(0).round(2)
    pivot["Clinic_Key"] = pivot["Clinic_Name"].apply(clean_text)

    return pivot[["Clinic_Name", "Clinic_Key", "PT", "Assistant", "PCC"]]


def _match_hours(pivot: pd.DataFrame, clinic_key: str):
    """Return (PT, Assistant, PCC) floats for a clinic key, or (0,0,0) if unmatched."""
    matched = pivot[pivot["Clinic_Key"] == clinic_key]
    if matched.empty:
        return 0.0, 0.0, 0.0
    row = matched.iloc[0]
    return float(row["PT"]), float(row["Assistant"]), float(row["PCC"])


# ================= DAILY BLOCK DETECTION =================

def get_current_top_block_row_count(ws) -> int:
    """
    Count how many rows the current top daily block occupies.
    Stops when the date in column A changes or a blank A+D row is hit
    after at least one real row.
    """
    start_row = DAILY_BLOCK_START_ROW
    rows = ws.get(f"A{start_row}:D1000")
    if not rows:
        raise ValueError("No rows found in the report sheet.")

    first_date = None
    count = 0

    for row in rows:
        row = row + [""] * (4 - len(row))
        date_val = str(row[0]).strip()   # A = Date
        loc_val  = str(row[3]).strip()   # D = Location

        # Skip leading blank rows
        if not date_val and not loc_val and count == 0:
            continue

        # Capture the first date we see
        if first_date is None and date_val:
            first_date = date_val

        # New date → new block begins, stop here
        if first_date and date_val and date_val != first_date:
            break

        # Blank row after content → block ended
        if not date_val and not loc_val and count > 0:
            break

        count += 1

    if count == 0:
        raise ValueError(
            "Could not detect daily block rows. "
            "Ensure Date is in column A and Location in column D."
        )
    return count


# ================= MAIN WRITE FUNCTION =================

def insert_daily_report_on_top(spreadsheet, final_schedule: pd.DataFrame, target_date: str):
    """
    Insert new daily block on top while preserving sheet design/format.
    
    Layout:
    A = Date
    D = Location
    V = PT Hours
    W = Assistant Hours
    X = PCC Hours
    """

    ws = spreadsheet.worksheet(SHEET_NAME)
    pivot = build_clinic_role_totals(final_schedule)

    start_row = DAILY_BLOCK_START_ROW
    block_rows = get_current_top_block_row_count(ws)
    end_row = start_row + block_rows - 1

    log.info(f"Detected top block: rows {start_row}-{end_row} ({block_rows} rows)")

    # Read current top block as template
    template = ws.get(f"A{start_row}:X{end_row}")

    sheet_id = ws.id

    start_index = start_row - 1
    end_index = start_index + block_rows

    # 1) Insert blank rows above current top block
    # 2) Copy old top block formatting/design into new rows
    spreadsheet.batch_update({
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": end_index
                    },
                    "inheritFromBefore": False
                }
            },
            {
                "copyPaste": {
                    "source": {
                        "sheetId": sheet_id,
                        "startRowIndex": end_index,
                        "endRowIndex": end_index + block_rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": 24
                    },
                    "destination": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_index,
                        "endRowIndex": end_index,
                        "startColumnIndex": 0,
                        "endColumnIndex": 24
                    },
                    "pasteType": "PASTE_NORMAL",
                    "pasteOrientation": "NORMAL"
                }
            }
        ]
    })

    date_values = []
    output_values = []
    unmatched = []

    for row in template:
        row = list(row) + [""] * (24 - len(row))

        location_name = str(row[3]).strip()  # D = Location
        location_key = clean_text(location_name)

        if location_key:
            date_values.append([target_date])
        else:
            date_values.append([""])

        pt, assistant, pcc = _match_hours(pivot, location_key)

        if location_key and pt == 0 and assistant == 0 and pcc == 0:
            unmatched.append(location_name)

        if location_key:
            output_values.append([pt, assistant, pcc])
        else:
            output_values.append(["", "", ""])

    new_end_row = start_row + block_rows - 1

    # Update only Date and Hours in the newly inserted block
    ws.update(
        range_name=f"A{start_row}:A{new_end_row}",
        values=date_values
    )

    ws.update(
        range_name=f"V{start_row}:X{new_end_row}",
        values=output_values
    )

    log.info(f"Inserted new daily report for {target_date} at row {start_row}.")

    if unmatched:
        log.warning("Clinics in Sheet1 with no Deputy data today:")
        for c in unmatched:
            log.warning(f"  - {c}")
            
# ================= MAIN =================

def run_report(date_override: str | None = None):
    target_date = get_target_date(date_override)
    log.info(f"Running report for: {target_date}")

    start_unix, end_unix = get_day_range_unix(target_date)

    roster_df = deputy_query("Roster", {
        "search": {
            "s1": {"field": "StartTime", "data": start_unix, "type": "ge"},
            "s2": {"field": "StartTime", "data": end_unix,   "type": "lt"},
        }
    })

    if roster_df.empty:
        log.warning("No roster data found for this date.")
        return

    final_schedule = build_schedule_df(roster_df)

    log.info("\nRole summary:")
    log.info(
        final_schedule
        .groupby(["Role_Name", "Role_Mapped"])["Total_Hours"]
        .sum()
        .reset_index()
        .sort_values(["Role_Mapped", "Role_Name"])
        .to_string(index=False)
    )

    spreadsheet = open_google_sheet()
    insert_daily_report_on_top(spreadsheet, final_schedule, target_date)
    log.info("Done.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Pull Deputy roster and write to Google Sheets."
    )
    parser.add_argument(
        "--date",
        metavar="YYYY-MM-DD",
        default=None,
        help="Target date (default: yesterday in NY time)",
    )
    args = parser.parse_args()
    run_report(date_override=args.date)