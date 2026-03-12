import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
import time
from flask import Flask, request, jsonify
import googleapiclient.discovery
import io

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)

START_ROW    = 13
ID_COL       = 0
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"
BUFFER_FILE  = "/tmp/sync_buffer.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# In-memory map: ID → source spreadsheet ID
id_to_source = {}

def get_client():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return gspread.authorize(creds)

def get_drive_service():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return googleapiclient.discovery.build("drive", "v3", credentials=creds)

def fetch_sheet(client, spreadsheet_id, sheet_name, start_row):
    """Fetch data from a native Google Sheet."""
    for attempt in range(3):
        try:
            ss       = client.open_by_key(spreadsheet_id)
            sheet    = ss.worksheet(sheet_name)
            all_rows = sheet.get_all_values()
            data     = all_rows[start_row - 1:]
            data     = [row for row in data if any(cell.strip() for cell in row)]
            return data
        except Exception as e:
            if "429" in str(e) or "quota" in str(e).lower():
                logging.warning(f"Rate limited, waiting 30s... (attempt {attempt + 1})")
                time.sleep(30)
            else:
                logging.warning(f"Failed to fetch {spreadsheet_id}/{sheet_name}: {e}")
                return []
    return []

def fetch_excel_source(file_id, sheet_name):
    """Download Excel file from Drive and read directly into memory as rows."""
    try:
        import openpyxl

        drive_service = get_drive_service()

        # Download the Excel file as bytes
        request_obj = drive_service.files().get_media(fileId=file_id)
        file_bytes  = io.BytesIO(request_obj.execute())

        # Load with openpyxl
        workbook  = openpyxl.load_workbook(file_bytes, data_only=True)

        # Find the sheet
        if sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
        else:
            logging.warning(f"Sheet '{sheet_name}' not found. Available: {workbook.sheetnames}")
            ws = workbook.active

        # Convert to list of rows starting from START_ROW
        all_rows = []
        for row in ws.iter_rows(min_row=START_ROW, values_only=True):
            row_as_strings = [str(cell) if cell is not None else "" for cell in row]
            if any(cell.strip() for cell in row_as_strings):
                all_rows.append(row_as_strings)

        logging.info(f"Read {len(all_rows)} rows from Excel file {file_id}")
        return all_rows

    except Exception as e:
        logging.error(f"Failed to read Excel file {file_id}: {e}")
        return []

def fetch_master(client):
    """Fetch master sheet data."""
    try:
        ss       = client.open_by_key(MASTER_SS_ID)
        sheet    = ss.worksheet(MASTER_SHEET)
        all_rows = sheet.get_all_values()
        data     = all_rows[START_ROW - 1:]
        data     = [row for row in data if any(cell.strip() for cell in row)]
        return sheet, data
    except Exception as e:
        logging.error(f"Failed to fetch master: {e}")
        return None, []

def save_to_buffer(data, source_id):
    """Save source data to JSON buffer."""
    try:
        buffer = {}
        if os.path.exists(BUFFER_FILE):
            with open(BUFFER_FILE, "r") as f:
                buffer = json.load(f)

        buffer[source_id] = {
            "rows":      data,
            "timestamp": time.time(),
            "count":     len(data)
        }

        with open(BUFFER_FILE, "w") as f:
            json.dump(buffer, f)

        logging.info(f"Saved {len(data)} rows to buffer for {source_id}")
        return True

    except Exception as e:
        logging.error(f"Failed to save to buffer: {e}")
        return False

def load_from_buffer():
    """Load all buffered data."""
    try:
        if not os.path.exists(BUFFER_FILE):
            return {}
        with open(BUFFER_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        logging.error(f"Failed to load buffer: {e}")
        return {}

def clear_buffer():
    """Clear buffer after successful sync."""
    try:
        if os.path.exists(BUFFER_FILE):
            os.remove(BUFFER_FILE)
            logging.info("Buffer cleared")
    except Exception as e:
        logging.warning(f"Failed to clear buffer: {e}")

def chunk_sources(sources, size=50):
    for i in range(0, len(sources), size):
        yield sources[i:i + size]

def get_all_sources():
    return [
        {"id": "1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt", "sheet": "GHIMS Incident Tracker", "type": "excel"},
        # {"id": "SPREADSHEET_ID", "sheet": "Sheet1", "type": "gsheet"},
        # add all sources here
    ]

@app.route("/sync", methods=["POST"])
def sync():
    global id_to_source
    try:
        data           = request.get_json()
        spreadsheet_id = data.get("spreadsheetId") if data else None
        sheet_name     = data.get("sheet") if data else None

        client = get_client()

        # Read master
        logging.info("Reading master sheet...")
        master_sheet_obj, master_data = fetch_master(client)
        if master_sheet_obj is None:
            return jsonify({"status": "error", "message": "Could not read master sheet"}), 500

        # Build master ID map
        id_to_index = {}
        for i, row in enumerate(master_data):
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if id_val:
                id_to_index[id_val] = i

        logging.info(f"Master has {len(id_to_index)} existing IDs")

        # ── STEP 1: Fetch sources → save to JSON buffer ───────────────────
        all_source_rows = {}

        if spreadsheet_id and sheet_name:
            logging.info(f"Targeted sync for {spreadsheet_id}/{sheet_name}")

            source_type = "gsheet"
            for s in get_all_sources():
                if s["id"] == spreadsheet_id:
                    source_type = s.get("type", "gsheet")
                    break

            if source_type == "excel":
                source_data = fetch_excel_source(spreadsheet_id, sheet_name)
            else:
                source_data = fetch_sheet(client, spreadsheet_id, sheet_name, START_ROW)

            save_to_buffer(source_data, spreadsheet_id)

            for row in source_data:
                id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                if not id_val:
                    continue
                all_source_rows[id_val] = row
                id_to_source[id_val]    = spreadsheet_id

        else:
            logging.info("Full sync — fetching all sources in batches...")
            sources   = get_all_sources()
            completed = 0

            for batch in chunk_sources(sources, size=50):
                for source in batch:
                    if source.get("type") == "excel":
                        source_data = fetch_excel_source(source["id"], source["sheet"])
                    else:
                        source_data = fetch_sheet(client, source["id"], source["sheet"], START_ROW)

                    save_to_buffer(source_data, source["id"])

                    for row in source_data:
                        id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                        if not id_val:
                            continue
                        all_source_rows[id_val] = row
                        id_to_source[id_val]    = source["id"]

                    completed += 1
                    logging.info(f"Fetched {completed}/{len(sources)} sources")

                if completed < len(sources):
                    logging.info("Batch done, pausing 15 seconds...")
                    time.sleep(15)

        logging.info(f"Buffer saved. Total source rows: {len(all_source_rows)}")

        # ── STEP 2: Load from buffer and validate ─────────────────────────
        buffer           = load_from_buffer()
        buffer_row_count = sum(v["count"] for v in buffer.values())
        logging.info(f"Buffer contains {len(buffer)} sources, {buffer_row_count} total rows")

        if buffer_row_count == 0:
            return jsonify({"status": "ok", "message": "Nothing to sync", "new": 0, "updates": 0, "removed": 0}), 200

        # ── STEP 3: Compare and write to master ───────────────────────────
        new_rows      = []
        updates       = []
        ids_to_remove = []

        if spreadsheet_id and sheet_name:
            for id_val, index in id_to_index.items():
                if id_to_source.get(id_val) == spreadsheet_id:
                    if id_val not in all_source_rows:
                        ids_to_remove.append(index)
                        id_to_source.pop(id_val, None)
                    else:
                        source_row  = all_source_rows[id_val]
                        master_row  = master_data[index]
                        max_len     = max(len(source_row), len(master_row))
                        source_norm = source_row + [""] * (max_len - len(source_row))
                        master_norm = master_row + [""] * (max_len - len(master_row))
                        if source_norm != master_norm:
                            updates.append((index, source_row))
        else:
            for id_val, index in id_to_index.items():
                if id_val not in all_source_rows:
                    ids_to_remove.append(index)
                    id_to_source.pop(id_val, None)
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
                master_sheet_obj.delete_rows(sheet_row)
                logging.info(f"Deleted row {sheet_row}")

            _, master_data = fetch_master(client)
            id_to_index    = {}
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
                master_sheet_obj.update(
                    f"A{sheet_row}",
                    [normalized],
                    value_input_option="RAW"
                )

        # Append new rows
        if new_rows:
            max_cols = max(len(row) for row in new_rows)
            new_rows = [row + [""] * (max_cols - len(row)) for row in new_rows]
            master_sheet_obj.append_rows(new_rows, value_input_option="RAW")

        # ── STEP 4: Clear buffer after successful sync ────────────────────
        clear_buffer()
        logging.info("Sync complete. Buffer cleared.")

        return jsonify({
            "status":                   "ok",
            "new":                      len(new_rows),
            "updates":                  len(updates),
            "removed":                  len(ids_to_remove),
            "buffer_sources_processed": len(buffer)
        }), 200

    except Exception as e:
        logging.error(f"Sync error: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/buffer", methods=["GET"])
def view_buffer():
    buffer  = load_from_buffer()
    summary = {}
    for source_id, val in buffer.items():
        summary[source_id] = {
            "rows":      val["count"],
            "timestamp": val["timestamp"]
        }
    return jsonify({
        "total_sources": len(buffer),
        "total_rows":    sum(v["count"] for v in buffer.values()),
        "sources":       summary
    })

@app.route("/diagnose", methods=["GET"])
def diagnose():
    try:
        rows = fetch_excel_source("1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt", "GHIMS Incident Tracker")
        return jsonify({
            "total_rows": len(rows),
            "first_row":  rows[0] if rows else "EMPTY",
            "second_row": rows[1] if len(rows) > 1 else "EMPTY"
        })
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/", methods=["GET"])
def health():
    return "OK", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
