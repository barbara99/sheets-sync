import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Config ───────────────────────────────────────────────────────────────────

START_ROW    = 30
ID_COL       = 0
MASTER_SS_ID = "YOUR_MASTER_SPREADSHEET_ID"
MASTER_SHEET = "Sheet7"

SOURCES = [
    {"id": "1PaYgXe2fzKkR-y-CXnei0RM2fQbCUlpVKoFAShjpX7w", "sheet": "Sheet1"},
    {"id": "1bIsyZ2cF-uAFI3fa7G8xGL98ZK8o5Ra5WWbV0FRLD88", "sheet": "Sheet1"},
    # add all 600 sources here
]

MAX_WORKERS = 20

# ── Setup ─────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Load credentials from environment variable (set in GitHub Secrets)
creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
client     = gspread.authorize(creds)

# ── Helpers ───────────────────────────────────────────────────────────────────

def fetch_sheet(spreadsheet_id, sheet_name, start_row):
    try:
        ss       = client.open_by_key(spreadsheet_id)
        sheet    = ss.worksheet(sheet_name)
        all_rows = sheet.get_all_values()
        data     = all_rows[start_row - 1:]
        data     = [row for row in data if any(cell.strip() for cell in row)]
        return data
    except Exception as e:
        logging.warning(f"Failed to fetch {spreadsheet_id}/{sheet_name}: {e}")
        return []

def fetch_source(source):
    return fetch_sheet(source["id"], source["sheet"], START_ROW)

# ── Main Sync ─────────────────────────────────────────────────────────────────

def sync():
    logging.info("Reading master sheet...")
    master_data = fetch_sheet(MASTER_SS_ID, MASTER_SHEET, START_ROW)

    existing_ids = set()
    for row in master_data:
        id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
        if id_val:
            existing_ids.add(id_val)

    logging.info(f"Master has {len(existing_ids)} existing IDs")

    logging.info(f"Fetching {len(SOURCES)} sources in parallel...")
    new_rows  = []
    completed = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(fetch_source, source): source for source in SOURCES}
        for future in as_completed(futures):
            source_data = future.result()
            completed  += 1
            if completed % 50 == 0:
                logging.info(f"  Fetched {completed}/{len(SOURCES)} sources...")
            for row in source_data:
                id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                if not id_val:
                    continue
                if id_val not in existing_ids:
                    existing_ids.add(id_val)
                    new_rows.append(row)

    logging.info(f"New rows to append: {len(new_rows)}")

    if not new_rows:
        logging.info("Nothing to append. Master is up to date.")
        return

    max_cols = max(len(row) for row in new_rows)
    new_rows = [row + [""] * (max_cols - len(row)) for row in new_rows]

    logging.info("Writing to master sheet...")
    ss           = client.open_by_key(MASTER_SS_ID)
    master_sheet = ss.worksheet(MASTER_SHEET)
    master_sheet.append_rows(new_rows, value_input_option="RAW")

    logging.info(f"Done. {len(new_rows)} rows appended.")

if __name__ == "__main__":
    sync()
