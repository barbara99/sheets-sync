import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
import time
import io
from flask import Flask, request, jsonify
import googleapiclient.discovery

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)

ID_COL       = 0
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"
BUFFER_FILE  = "/tmp/sync_buffer.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

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

def fetch_excel_source(file_id, sheet_name, start_row):
    try:
        import openpyxl

        drive_service = get_drive_service()
        request_obj   = drive_service.files().get_media(fileId=file_id)
        file_bytes    = io.BytesIO(request_obj.execute())

        workbook = openpyxl.load_workbook(file_bytes, data_only=True)

        if sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
        else:
            logging.warning(f"Sheet '{sheet_name}' not found. Available: {workbook.sheetnames}")
            ws = workbook.active

        all_rows = []
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            row_as_strings = [str(cell) if cell is not None else "" for cell in row]

            if not any(cell.strip() for cell in row_as_strings):
                continue

            while row_as_strings and row_as_strings[0] == "":
                row_as_strings.pop(0)

            while row_as_strings and row_as_strings[-1] == "":
                row_as_strings.pop()

            if row_as_strings:
                all_rows.append(row_as_strings)

        logging.info(f"Read {len(all_rows)} rows from Excel file {file_id}")
        return all_rows

    except Exception as e:
        logging.error(f"Failed to read Excel file {file_id}: {e}")
        return []

def fetch_master(client):
    try:
        ss       = client.open_by_key(MASTER_SS_ID)
        sheet    = ss.worksheet(MASTER_SHEET)
        all_rows = sheet.get_all_values()
        data     = [row for row in all_rows if any(cell.strip() for cell in row)]
        return sheet, data
    except Exception as e:
        logging.error(f"Failed to fetch master: {e}")
        return None, []

def get_master_row_map(client):
    """Get accurate 1-indexed row positions from master sheet."""
    try:
        ss       = client.open_by_key(MASTER_SS_ID)
        sheet    = ss.worksheet(MASTER_SHEET)
        all_rows = sheet.get_all_values()
        row_map  = {}
        for i, row in enumerate(all_rows):
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if id_val:
                row_map[id_val] = i + 1  # 1-indexed
        return sheet, row_map
    except Exception as e:
        logging.error(f"Failed to get master row map: {e}")
        return None, {}

def save_to_buffer(data, source_id):
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
    try:
        if not os.path.exists(BUFFER_FILE):
            return {}
        with open(BUFFER_FILE, "r") as f:
            return json.load(f)
    except Exception as e:
        logging.error(f"Failed to load buffer: {e}")
        return {}

def clear_buffer():
    try:
        if os.path.exists(BUFFER_FILE):
            os.remove(BUFFER_FILE)
            logging.info("Buffer cleared")
    except Exception as e:
        logging.warning(f"Failed to clear buffer: {e}")

def chunk_sources(sources, size=50):
    for i in range(0, len(sources), size):
        yield sources[i:i + size]

def get_source_data(client, source):
    start_row = source.get("start_row", 13)
    if source.get("type") == "excel":
        return fetch_excel_source(source["id"], source["sheet"], start_row)
    else:
        return fetch_sheet(client, source["id"], source["sheet"], start_row)

def get_all_sources():
    return [
        {
            "id":        "1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt",
            "sheet":     "GHIMS Incident Tracker",
            "type":      "excel",
            "start_row": 13
        },
        # add more sources here:
        # {
        #     "id":        "SPREADSHEET_ID",
        #     "sheet":     "Sheet1",
        #     "type":      "gsheet",  # or "excel"
        #     "start_row": 5
        # },
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

        # Build master ID map (0-indexed into master_data)
        id_to_index = {}
        for i, row in enumerate(master_data):
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if id_val:
                id_to_index[id_val] = i

        logging.info(f"Master has {len(id_to_index)} existing IDs")
        logging.info(f"Master ID sample: {list(id_to_index.keys())[:5]}")

        # ── STEP 1: Fetch sources → save to JSON buffer ───────────────────
        all_source_rows = {}

        if spreadsheet_id and sheet_name:
            logging.info(f"Targeted sync for {spreadsheet_id}/{sheet_name}")

            matched_source = None
            for s in get_all_sources():
                if s["id"] == spreadsheet_id:
                    matched_source = s
                    break

            source_data = get_source_data(client, matched_source) if matched_source else \
                          fetch_sheet(client, spreadsheet_id, sheet_name, 13)

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
                    source_data = get_source_data(client, source)
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

        logging.info(f"Source rows fetched: {len(all_source_rows)}")
        logging.info(f"Source ID sample: {list(all_source_rows.keys())[:5]}")

        # ── STEP 2: Load and validate buffer ─────────────────────────────
        buffer           = load_from_buffer()
        buffer_row_count = sum(v["count"] for v in buffer.values())
        logging.info(f"Buffer: {len(buffer)} sources, {buffer_row_count} total rows")

        if buffer_row_count == 0:
            return jsonify({"status": "ok", "message": "Nothing to sync", "new": 0, "updates": 0, "removed": 0}), 200

        # ── STEP 3: Compare ───────────────────────────────────────────────
        new_rows      = []
        updates       = []
        ids_to_remove = []

        if spreadsheet_id and sheet_name:
            for id_val, index in id_to_index.items():
                if id_to_source.get(id_val) == spreadsheet_id:
                    if id_val not in all_source_rows:
                        ids_to_remove.append(id_val)
                        id_to_source.pop(id_val, None)
                    else:
                        source_row  = all_source_rows[id_val]
                        master_row  = master_data[index]
                        max_len     = max(len(source_row), len(master_row))
                        source_norm = source_row + [""] * (max_len - len(source_row))
                        master_norm = master_row + [""] * (max_len - len(master_row))
                        if source_norm != master_norm:
                            updates.append((id_val, source_row))
        else:
            for id_val, index in id_to_index.items():
                if id_val not in all_source_rows:
                    ids_to_remove.append(id_val)
                    id_to_source.pop(id_val, None)
                else:
                    source_row  = all_source_rows[id_val]
                    master_row  = master_data[index]
                    max_len     = max(len(source_row), len(master_row))
                    source_norm = source_row + [""] * (max_len - len(source_row))
                    master_norm = master_row + [""] * (max_len - len(master_row))
                    if source_norm != master_norm:
                        updates.append((id_val, source_row))

        for id_val, row in all_source_rows.items():
            if id_val not in id_to_index:
                new_rows.append(row)

        logging.info(f"New: {len(new_rows)} | Updates: {len(updates)} | Removals: {len(ids_to_remove)}")

        # Get accurate row positions for writes
        _, row_map = get_master_row_map(client)

        # Apply removals (bottom up to preserve row numbers)
        if ids_to_remove:
            rows_to_delete = sorted(
                [row_map[id_val] for id_val in ids_to_remove if id_val in row_map],
                reverse=True
            )
            for sheet_row in rows_to_delete:
                master_sheet_obj.delete_rows(sheet_row)
                logging.info(f"Deleted row {sheet_row}")

            # Refresh row map after deletions
            _, row_map = get_master_row_map(client)

        # Apply updates using accurate row positions
        if updates:
            max_cols = max(len(row) for _, row in updates)
            for id_val, row in updates:
                sheet_row = row_map.get(id_val)
                if sheet_row:
                    normalized = row + [""] * (max_cols - len(row))
                    master_sheet_obj.update(
                        f"A{sheet_row}",
                        [normalized],
                        value_input_option="RAW"
                    )
                    logging.info(f"Updated row {sheet_row} for ID {id_val}")

        # Append new rows
        if new_rows:
            max_cols = max(len(row) for row in new_rows)
            new_rows = [row + [""] * (max_cols - len(row)) for row in new_rows]
            master_sheet_obj.append_rows(new_rows, value_input_option="RAW")
            logging.info(f"Appended {len(new_rows)} new rows")

        # ── STEP 4: Clear buffer ──────────────────────────────────────────
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
        client  = get_client()
        source  = get_all_sources()[0]
        rows    = get_source_data(client, source)
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
    
@app.route("/dedupe", methods=["POST"])
def dedupe():
    try:
        client = get_client()

        # Read full master sheet including row positions
        ss       = client.open_by_key(MASTER_SS_ID)
        sheet    = ss.worksheet(MASTER_SHEET)
        all_rows = sheet.get_all_values()

        seen_ids      = {}
        rows_to_delete = []

        for i, row in enumerate(all_rows):
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if not id_val:
                continue

            if id_val in seen_ids:
                # Duplicate — mark for deletion
                rows_to_delete.append(i + 1)  # 1-indexed sheet row
                logging.info(f"Duplicate found: {id_val} at row {i + 1}")
            else:
                seen_ids[id_val] = i + 1  # keep first occurrence

        logging.info(f"Found {len(rows_to_delete)} duplicate rows to delete")

        # Delete from bottom up to preserve row numbers
        for sheet_row in sorted(rows_to_delete, reverse=True):
            sheet.delete_rows(sheet_row)
            logging.info(f"Deleted duplicate at row {sheet_row}")

        return jsonify({
            "status":            "ok",
            "duplicates_removed": len(rows_to_delete),
            "unique_ids_kept":   len(seen_ids)
        }), 200

    except Exception as e:
        logging.error(f"Dedupe error: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
