import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
from flask import Flask, request, jsonify
from concurrent.futures import ThreadPoolExecutor, as_completed

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)

START_ROW    = 30
ID_COL       = 0
MASTER_SS_ID = "YOUR_MASTER_SPREADSHEET_ID"
MASTER_SHEET = "Sheet7"

SOURCES = [
    {"id": "1PaYgXe2fzKkR-y-CXnei0RM2fQbCUlpVKoFAShjpX7w", "sheet": "Sheet1"},
    {"id": "1bIsyZ2cF-uAFI3fa7G8xGL98ZK8o5Ra5WWbV0FRLD88", "sheet": "Sheet1"},
]

MAX_WORKERS = 20

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_client():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return gspread.authorize(creds)

def fetch_sheet(client, spreadsheet_id, sheet_name, start_row):
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

@app.route("/sync", methods=["POST"])
def sync():
    try:
        client = get_client()

        # Read master
        logging.info("Reading master sheet...")
        master_ss    = client.open_by_key(MASTER_SS_ID)
        master_sheet = master_ss.worksheet(MASTER_SHEET)
        master_data  = fetch_sheet(client, MASTER_SS_ID, MASTER_SHEET, START_ROW)

        id_to_index = {}
        for i, row in enumerate(master_data):
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if id_val:
                id_to_index[id_val] = i

        logging.info(f"Master has {len(id_to_index)} existing IDs")

        # Fetch all sources in parallel
        def fetch_source(source):
            return fetch_sheet(client, source["id"], source["sheet"], START_ROW)

        all_source_rows = {}
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(fetch_source, source): source for source in SOURCES}
            for future in as_completed(futures):
                source_data = future.result()
                for row in source_data:
                    id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                    if not id_val:
                        continue
                    all_source_rows[id_val] = row

        # Figure out changes
        new_rows      = []
        updates       = []
        ids_to_remove = []

        for id_val, index in id_to_index.items():
            if id_val not in all_source_rows:
                ids_to_remove.append(index)
            else:
                source_row  = all_source_rows[id_val]
                master_row  = master_data[index]
                max_len     = max(len(source_row), len(master_row))
                source_norm = source_row + [""] * (max_len - len(source_row))
                master_norm = master_row + [""] * (max_len - len(master_row))
                if source_norm != master_norm:
                    updates.append((index, source_row))

        for id_val, row in all_source_rows.items():
            if id_val not in id_to_index:
                new_rows.append(row)

        logging.info(f"New: {len(new_rows)} | Updates: {len(updates)} | Removals: {len(ids_to_remove)}")

        # Apply removals
        if ids_to_remove:
            rows_to_delete = sorted(
                [START_ROW + i for i in ids_to_remove],
                reverse=True
            )
            for sheet_row in rows_to_delete:
                master_sheet.delete_rows(sheet_row)

            master_data = fetch_sheet(client, MASTER_SS_ID, MASTER_SHEET, START_ROW)
            id_to_index = {}
            for i, row in enumerate(master_data):
                id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                if id_val:
                    id_to_index[id_val] = i

        # Apply updates
        if updates:
            max_cols = max(len(row) for _, row in updates)
            for index, row in updates:
                normalized = row + [""] * (max_cols - len(row))
                sheet_row  = START_ROW + index
                master_sheet.update(
                    f"A{sheet_row}",
                    [normalized],
                    value_input_option="RAW"
                )

        # Append new rows
        if new_rows:
            max_cols = max(len(row) for row in new_rows)
            new_rows = [row + [""] * (max_cols - len(row)) for row in new_rows]
            master_sheet.append_rows(new_rows, value_input_option="RAW")

        return jsonify({"status": "ok", "new": len(new_rows), "updates": len(updates), "removed": len(ids_to_remove)}), 200

    except Exception as e:
        logging.error(f"Sync error: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/", methods=["GET"])
def health():
    return "OK", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

web: gunicorn main:app --bind 0.0.0.0:$PORT
