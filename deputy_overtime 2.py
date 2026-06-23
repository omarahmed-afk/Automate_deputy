import sys
import datetime as dt
from zoneinfo import ZoneInfo

import requests
import gspread
from google.oauth2.service_account import Credentials

# ============================== CONFIG ==============================
TOKEN = ""

INSTALL_URL = "https://ptofthecity.na.deputy.com"
API_BASE    = f"{INSTALL_URL}/api/v1"

# ---- Google Sheets ----
GOOGLE_CREDENTIALS_FILE = r""
SPREADSHEET_ID = ""           


# ---- Rules ----
NY               = ZoneInfo("America/New_York")
OT_THRESHOLD_MIN = 11                                  # keep employees 11+ minutes past shift end
TARGET_DATE      = None                                # None = auto by weekday, or "2026-06-01"
OT_MODE          = "end"
CLINICS_FILTER   = []
# ===================================================================

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
}


def query_resource(resource: str, payload: dict) -> list:
    """Call /api/v1/resource/<Name>/QUERY with pagination (500 max per request)."""
    out, start = [], 0
    while True:
        body = dict(payload, max=500, start=start)
        url = f"{API_BASE}/resource/{resource}/QUERY"
        resp = requests.post(url, json=body, headers=HEADERS, timeout=60)
        if resp.status_code == 401:
            sys.exit("Token rejected (401). It may have expired - paste a fresh one.")
        resp.raise_for_status()
        chunk = resp.json()
        if not isinstance(chunk, list):
            raise RuntimeError(f"Unexpected response from {resource}: {chunk}")
        out.extend(chunk)
        if len(chunk) < 500:
            break
        start += 500
    return out


def validate_token():
    r = requests.get(f"{API_BASE}/me", headers=HEADERS, timeout=30)
    if r.status_code == 401:
        sys.exit("Token rejected (401). The 24h window may be over - paste a fresh one.")
    r.raise_for_status()
    me = r.json()
    print(f"Token OK. Logged in as: {me.get('Name', me.get('DisplayName', '?'))}")


def target_range(target):
    """Monday -> previous Fri/Sat/Sun; Tue-Fri -> previous day; or a fixed date string."""
    today = dt.datetime.now(NY).date()
    if isinstance(target, str):
        start_date = end_date = dt.date.fromisoformat(target)
    elif today.weekday() == 0:
        start_date = today - dt.timedelta(days=3)
        end_date = today - dt.timedelta(days=1)
    else:
        start_date = end_date = today - dt.timedelta(days=1)
    start = dt.datetime.combine(start_date, dt.time.min, NY)
    end = dt.datetime.combine(end_date, dt.time.max, NY)
    return start_date, end_date, int(start.timestamp()), int(end.timestamp())


def to_ny(unix_ts: int) -> dt.datetime:
    return dt.datetime.fromtimestamp(unix_ts, NY)


# Only these roles appear in the report. Anything else -> row is dropped.
ROLE_MAP = {
    "physical therapist": "PT",
    "physical threapist": "PT",                 # misspelling found in the data
    "physical therapy assistant": "PTA",
    "patients care coordinator": "PCC",
    "patients care coordinator (pcc)": "PCC",
    "aide": "PT Aide",
}


def map_role(name: str):
    """Return the short code for an approved role, or None to drop the row."""
    return ROLE_MAP.get((name or "").strip().lower())


def fetch_timesheets(start_unix, end_unix):
    """Fetch the day's timesheets with their scheduled Roster attached."""
    search = {
        "f1": {"field": "StartTime", "data": start_unix, "type": "ge"},
        "f2": {"field": "StartTime", "data": end_unix, "type": "le"},
    }
    try:
        timesheets = query_resource("Timesheet", {"search": search, "join": ["RosterObject"]})
    except RuntimeError:
        timesheets = query_resource("Timesheet", {"search": search})
    print(f"  Found {len(timesheets)} timesheets.")

    if any(t.get("RosterObject") for t in timesheets):
        for t in timesheets:
            t["_roster"] = t.get("RosterObject")
        print("  Roster pulled inline via join.")
        return timesheets

    roster_ids = sorted({t["Roster"] for t in timesheets if t.get("Roster")})
    rosters = {}
    for i in range(0, len(roster_ids), 200):
        rp = {"search": {"r1": {"field": "Id", "data": roster_ids[i:i + 200], "type": "in"}}}
        for r in query_resource("Roster", rp):
            rosters[r["Id"]] = r
    for t in timesheets:
        t["_roster"] = rosters.get(t.get("Roster"))
    print(f"  Found {len(rosters)} scheduled rosters (separate query).")
    return timesheets


def build_rows(start_unix, end_unix):
    timesheets = fetch_timesheets(start_unix, end_unix)

    rows = []
    for t in timesheets:
        actual_end = t.get("EndTime")
        if not actual_end:                            # shift still in progress
            continue
        roster = t.get("_roster")
        if not roster or not roster.get("EndTime"):
            continue                                  # no scheduled shift to compare

        sched_end = roster["EndTime"]
        if OT_MODE == "total":
            ot_seconds = (actual_end - t["StartTime"]) - (roster["EndTime"] - roster["StartTime"])
        else:                                         # "end"
            ot_seconds = actual_end - sched_end

        if ot_seconds < OT_THRESHOLD_MIN * 60:        # keep 11 minutes or more
            continue

        meta = t.get("_DPMetaData", {})
        emp = meta.get("EmployeeInfo", {}).get("DisplayName", "")
        ou = meta.get("OperationalUnitInfo", {})
        facility = ou.get("CompanyName", "")
        if "pediatric" in facility.lower():           # skip any Pediatric facility
            continue
        role = map_role(ou.get("OperationalUnitName", ""))
        if role is None:                              # role not approved -> skip
            continue
        if CLINICS_FILTER and facility not in CLINICS_FILTER:
            continue

        rows.append([
            emp,                                              # C
            facility,                                         # E
            role,                                             # F
            to_ny(t["StartTime"]).strftime("%m/%d/%Y"),       # G
            int(round(ot_seconds / 60)),                      # J (number only, no "MIN")
            to_ny(sched_end).strftime("%I:%M %p"),            # L
            (t.get("SupervisorComment") or t.get("EmployeeComment")
             or roster.get("Comment") or ""),                # M
        ])

    rows.sort(key=lambda r: (r[1], r[0]))             # by facility, then name
    print(f"  {len(rows)} employees stayed {OT_THRESHOLD_MIN}+ minutes past their shift.")
    return rows


# --------------------------- Google Sheets ---------------------------
def open_worksheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)


def write_to_sheet(rows):
    if not rows:
        print("  Nothing to write.")
        return
    ws = open_worksheet()
    last_used = len(ws.col_values(ANCHOR_COL))        # last filled row in column C
    start_row = max(FIRST_DATA_ROW, last_used + 1)
    end_row = start_row + len(rows) - 1

    def col(idx):
        return [[r[idx]] for r in rows]

    ws.batch_update([
        {"range": f"C{start_row}:C{end_row}", "values": col(0)},   # EMPLOYEE NAME
        {"range": f"E{start_row}:E{end_row}", "values": col(1)},   # FACILITY
        {"range": f"F{start_row}:F{end_row}", "values": col(2)},   # ROLE
        {"range": f"G{start_row}:G{end_row}", "values": col(3)},   # DATE
        {"range": f"J{start_row}:J{end_row}", "values": col(4)},   # OVERTIME (number)
        {"range": f"L{start_row}:L{end_row}", "values": col(5)},   # SCHD. SHIFT END
        {"range": f"M{start_row}:M{end_row}", "values": col(6)},   # REASON
    ], value_input_option="USER_ENTERED")
    print(f"  Wrote {len(rows)} rows to '{WORKSHEET_NAME}' (rows {start_row}-{end_row}).")


def main():
    if not TOKEN:
        sys.exit("Please paste your Deputy access token into the TOKEN variable.")
    if SPREADSHEET_ID == "PASTE_SPREADSHEET_ID_HERE":
        sys.exit("Please set SPREADSHEET_ID to your Google Sheet's id.")

    validate_token()
    start_date, end_date, start_unix, end_unix = target_range(TARGET_DATE)
    if start_date == end_date:
        print(f"Date: {start_date.strftime('%m/%d/%Y')} (NY time)")
    else:
        print(f"Dates: {start_date.strftime('%m/%d/%Y')} - "
              f"{end_date.strftime('%m/%d/%Y')} (NY time)")

    rows = build_rows(start_unix, end_unix)
    write_to_sheet(rows)
    print("Done.")


if __name__ == "__main__":
    main()