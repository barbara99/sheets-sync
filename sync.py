
import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Config ───────────────────────────────────────────────────────────────────

START_ROW    = 23
ID_COL       = 0
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"

SOURCES = [
    {"id": "1PaYgXe2fzKkR-y-CXnei0RM2fQbCUlpVKoFAShjpX7w", "sheet": "Sheet1"},
    {"id": "1bIsyZ2cF-uAFI3fa7G8xGL98ZK8o5Ra5WWbV0FRLD88", "sheet": "Sheet1"},
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
    # Step 1: Read master sheet
    logging.info("Reading master sheet...")
    master_ss    = client.open_by_key(MASTER_SS_ID)
    master_sheet = master_ss.worksheet(MASTER_SHEET)
    master_data  = fetch_sheet(MASTER_SS_ID, MASTER_SHEET, START_ROW)

    # Build map: ID → row index in master_data
    id_to_index = {}
    for i, row in enumerate(master_data):
        id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
        if id_val:
            id_to_index[id_val] = i

    logging.info(f"Master has {len(id_to_index)} existing IDs")

    # Step 2: Fetch all sources in parallel
    logging.info(f"Fetching {len(SOURCES)} sources in parallel...")
    all_source_rows = {}  # id → latest row from sources
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
                all_source_rows[id_val] = row  # latest source wins on duplicate

    logging.info(f"Total unique IDs found across all sources: {len(all_source_rows)}")

    # Step 3: Figure out what to add, update, or remove
    new_rows     = []  # rows to append
    updates      = []  # (row_index_in_master, new_row)
    ids_to_remove = [] # row indexes in master to delete

    # Check for updates and deletions
    for id_val, index in id_to_index.items():
        if id_val not in all_source_rows:
            # ID no longer exists in any source — remove it
            ids_to_remove.append(index)
        else:
            # ID exists — check if data changed
            source_row  = all_source_rows[id_val]
            master_row  = master_data[index]
            # Normalize lengths for comparison
            max_len     = max(len(source_row), len(master_row))
            source_norm = source_row + [""] * (max_len - len(source_row))
            master_norm = master_row + [""] * (max_len - len(master_row))
            if source_norm != master_norm:
                updates.append((index, source_row))

    # Check for new IDs
    for id_val, row in all_source_rows.items():
        if id_val not in id_to_index:
            new_rows.append(row)

    logging.info(f"New: {len(new_rows)} | Updates: {len(updates)} | Removals: {len(ids_to_remove)}")

    # Step 4: Apply removals (delete from bottom up to preserve row indexes)
    if ids_to_remove:
        # Convert master_data indexes to actual sheet row numbers
        # Sort descending so we delete from bottom up
        rows_to_delete = sorted(
            [START_ROW + i for i in ids_to_remove],
            reverse=True
        )
        for sheet_row in rows_to_delete:
            master_sheet.delete_rows(sheet_row)
            logging.info(f"Deleted row {sheet_row}")

        # Re-read master after deletions
        master_data = fetch_sheet(MASTER_SS_ID, MASTER_SHEET, START_ROW)
        id_to_index = {}
        for i, row in enumerate(master_data):
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if id_val:
                id_to_index[id_val] = i

    # Step 5: Apply updates in one batch
    if updates:
        # Normalize all rows to same column width
        max_cols = max(len(row) for _, row in updates)
        for index, row in updates:
            normalized = row + [""] * (max_cols - len(row))
            sheet_row  = START_ROW + index
            master_sheet.update(
                f"A{sheet_row}",
                [normalized],
                value_input_option="RAW"
            )
        logging.info(f"Updated {len(updates)} rows")

    # Step 6: Append new rows in one batch
    if new_rows:
        max_cols = max(len(row) for row in new_rows)
        new_rows = [row + [""] * (max_cols - len(row)) for row in new_rows]
        master_sheet.append_rows(new_rows, value_input_option="RAW")
        logging.info(f"Appended {len(new_rows)} new rows")

    logging.info("Sync complete.")

if __name__ == "__main__":
    sync()
