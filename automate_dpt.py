import requests
import pandas as pd
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

# ================= CONFIG =================

TOKEN = "b12c6ed397378d77df567e3835cdaad5"

INSTALL_URL = "https://ptofthecity.na.deputy.com"
API_BASE = f"{INSTALL_URL}/api/v1"

NY = ZoneInfo("America/New_York")

# For daily automation, keep None = yesterday NY date
# For testing today, write for example: "2026-06-23"
TARGET_DATE = "2026-06-23"

OUTPUT_FOLDER = Path(r"D:\depty_overtime")
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = OUTPUT_FOLDER / "role_hours_daily.xlsx"

ROLE_MAP = {
    "physical therapist": "PT",
    "physical threapist": "PT",
    "physical therapy assistant": "PTA",
    "patients care coordinator": "PCC",
    "patients care coordinator (pcc)": "PCC",
    "aide": "PT Aide",
}

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

# ==========================================


def get_target_date():
    if TARGET_DATE is not None:
        return TARGET_DATE

    # daily automation pulls yesterday based on New York time
    yesterday = datetime.now(NY).date() - timedelta(days=1)
    return str(yesterday)


def get_day_range_unix(target_date):
    start_dt = datetime.strptime(target_date, "%Y-%m-%d").replace(tzinfo=NY)
    end_dt = start_dt + timedelta(days=1)

    return int(start_dt.timestamp()), int(end_dt.timestamp())


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


def clean_role(role):
    if pd.isna(role):
        return None

    return str(role).strip().lower()


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
    else:
        return f"{hour_12}:{minute:02d}{am_pm}"


def build_schedule_df(roster_df):
    schedule_df = roster_df[
        [
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
    ].copy()

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

    # delete rows where employee name is empty
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

    # clean role
    schedule_df["Role_Key"] = schedule_df["Role_Name"].apply(clean_role)

    schedule_df["Role_Mapped"] = (
        schedule_df["Role_Key"]
        .map(ROLE_MAP)
        .fillna(schedule_df["Role_Name"])
    )

    # clean readable shift time
    schedule_df["Start_Clean"] = schedule_df["Scheduled_Start"].apply(clean_time)
    schedule_df["End_Clean"] = schedule_df["Scheduled_End"].apply(clean_time)

    schedule_df["Shift_Time"] = (
        schedule_df["Start_Clean"] + " – " + schedule_df["End_Clean"]
    )

    # convert unix times to New York datetime
    schedule_df["Start_DT_NY"] = (
        pd.to_datetime(schedule_df["Start_Unix"], unit="s", utc=True)
        .dt.tz_convert(NY)
    )

    schedule_df["End_DT_NY"] = (
        pd.to_datetime(schedule_df["End_Unix"], unit="s", utc=True)
        .dt.tz_convert(NY)
    )

    # final date based on New York time
    schedule_df["Date_Clean"] = schedule_df["Start_DT_NY"].dt.strftime("%Y-%m-%d")

    # calculate scheduled hours
    schedule_df["Total_Hours"] = (
        schedule_df["End_Unix"] - schedule_df["Start_Unix"]
    ) / 3600

    schedule_df["Total_Hours"] = schedule_df["Total_Hours"].round(2)

    # choose only needed columns for schedule sheet
    needed_columns = [
        "Date_Clean",
        "Clinic_Name",
        "Role_Mapped",
        "Employee_Name",
        "Shift_Time",
        "Total_Hours",
        "Roster_Id",
        "Timesheet_Id"
    ]

    final_schedule = schedule_df[needed_columns].copy()

    final_schedule = final_schedule.rename(columns={
        "Date_Clean": "Date"
    })

    return final_schedule


def total_hours_by_day_role_clinic(final_schedule):
    hours_by_day_role_clinic = (
        final_schedule
        .groupby(
            ["Date", "Role_Mapped", "Clinic_Name"],
            as_index=False
        )["Total_Hours"]
        .sum()
        .sort_values(["Date", "Role_Mapped", "Clinic_Name"])
    )

    hours_by_day_role_clinic["Total_Hours"] = (
        hours_by_day_role_clinic["Total_Hours"].round(2)
    )

    return hours_by_day_role_clinic


def total_hours_pivot(hours_by_day_role_clinic):
    pivot_df = (
        hours_by_day_role_clinic
        .pivot_table(
            index=["Date", "Clinic_Name"],
            columns="Role_Mapped",
            values="Total_Hours",
            fill_value=0
        )
        .reset_index()
    )

    pivot_df.columns.name = None

    return pivot_df


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

    hours_by_day_role_clinic = total_hours_by_day_role_clinic(final_schedule)

    hours_pivot = total_hours_pivot(hours_by_day_role_clinic)

    print("Saving to:", OUTPUT_FILE)

    with pd.ExcelWriter(str(OUTPUT_FILE), engine="openpyxl") as writer:
        final_schedule.to_excel(writer, sheet_name="Schedule", index=False)
        hours_by_day_role_clinic.to_excel(writer, sheet_name="Total Hours", index=False)
        hours_pivot.to_excel(writer, sheet_name="Pivot", index=False)

    print("Done. File saved:")
    print(OUTPUT_FILE)


if __name__ == "__main__":
    run_report()