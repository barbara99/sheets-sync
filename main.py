import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
import time
import io
import re
from flask import Flask, request, jsonify
import googleapiclient.discovery

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)

ID_COL       = 0
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"
MASTER_START = 13
BUFFER_FILE  = "/tmp/sync_buffer.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

id_to_source = {}

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_client():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return gspread.authorize(creds)

def get_drive_service():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return googleapiclient.discovery.build("drive", "v3", credentials=creds)

# ── Parsers ───────────────────────────────────────────────────────────────────

def parse_excel(file_bytes, sheet_name, start_row):
    import openpyxl
    workbook = openpyxl.load_workbook(file_bytes, data_only=True)

    if sheet_name and sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
    else:
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
    return all_rows

def parse_csv(content, start_row):
    import csv
    try:
        text = content.decode("utf-8")
    except:
        text = content.decode("latin-1")
    reader   = csv.reader(text.splitlines())
    all_rows = list(reader)
    data     = all_rows[start_row - 1:]
    data     = [row for row in data if any(cell.strip() for cell in row)]
    return data

# ── Fetchers ──────────────────────────────────────────────────────────────────

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
        drive_service = get_drive_service()
        request_obj   = drive_service.files().get_media(fileId=file_id)
        file_bytes    = io.BytesIO(request_obj.execute())
        return parse_excel(file_bytes, sheet_name, start_row)
    except Exception as e:
        logging.error(f"Failed to read Excel file {file_id}: {e}")
        return []

def detect_and_fetch_url(url, sheet_name, start_row):
    """Download file from any URL and auto-detect type."""

    # Google Sheets URL → use Sheets API
    if "docs.google.com/spreadsheets/d/" in url:
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        if match:
            spreadsheet_id = match.group(1)
            logging.info(f"Google Sheets URL detected, using API: {spreadsheet_id}")
            client = get_client()
            return fetch_sheet(client, spreadsheet_id, sheet_name or "Sheet1", start_row)
        else:
            logging.error(f"Could not extract ID from Google Sheets URL: {url}")
            return []

    # Google Drive file URL → use Drive API
    if "drive.google.com" in url:
        match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        if match:
            file_id = match.group(1)
            logging.info(f"Google Drive URL detected, using Drive API: {file_id}")
            return fetch_excel_source(file_id, sheet_name, start_row)
        else:
            logging.error(f"Could not extract ID from Drive URL: {url}")
            return []

    # Everything else — download and auto-detect
    try:
        import requests

        logging.info(f"Downloading from URL: {url}")
        response = requests.get(url, timeout=30, allow_redirects=True)

        if response.status_code != 200:
            logging.error(f"Failed to download {url}: HTTP {response.status_code}")
            return []

        content_type = response.headers.get("Content-Type", "").lower()
        url_lower    = url.lower()
        logging.info(f"Content-Type: {content_type}")

        # Detect Excel
        if any(x in content_type for x in ["excel", "spreadsheetml", "openxmlformats"]) or \
           any(url_lower.endswith(x) for x in [".xlsx", ".xls"]):
            logging.info("Detected: Excel")
            return parse_excel(io.BytesIO(response.content), sheet_name, start_row)

        # Detect CSV
        elif "csv" in content_type or url_lower.endswith(".csv"):
            logging.info("Detected: CSV")
            return parse_csv(response.content, start_row)

        # Try Excel first, then CSV
        else:
            logging.info("Unknown type — trying Excel first, then CSV")
            try:
                result = parse_excel(io.BytesIO(response.content), sheet_name, start_row)
                if result:
                    return result
            except:
                pass
            try:
                result = parse_csv(response.content, start_row)
                if result:
                    return result
            except:
                pass
            logging.error(f"Could not parse file from {url}")
            return []

    except Exception as e:
        logging.error(f"Failed to fetch URL source {url}: {e}")
        return []

def fetch_master(client):
    try:
        ss       = client.open_by_key(MASTER_SS_ID)
        sheet    = ss.worksheet(MASTER_SHEET)
        all_rows = sheet.get_all_values()

        id_to_row  = {}
        id_to_data = {}

        for i, row in enumerate(all_rows):
            sheet_row = i + 1
            if sheet_row < MASTER_START:
                continue
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if id_val:
                id_to_row[id_val]  = sheet_row
                id_to_data[id_val] = row

        last_row = MASTER_START - 1
        for i in range(len(all_rows) - 1, -1, -1):
            if any(cell.strip() for cell in all_rows[i]):
                last_row = i + 1
                break
        last_row = max(last_row, MASTER_START - 1)

        return sheet, id_to_row, id_to_data, last_row

    except Exception as e:
        logging.error(f"Failed to fetch master: {e}")
        return None, {}, {}, MASTER_START - 1

# ── Source router ─────────────────────────────────────────────────────────────

def get_source_data(client, source):
    start_row = source.get("start_row", 13)

    if "url" in source:
        # Auto-detect from URL — handles Google Sheets, Drive, Excel, CSV, etc.
        return detect_and_fetch_url(
            source["url"],
            source.get("sheet"),
            start_row
        )
    elif source.get("type") == "excel":
        return fetch_excel_source(source["id"], source["sheet"], start_row)
    else:
        return fetch_sheet(client, source["id"], source["sheet"], start_row)

# ── Source list ───────────────────────────────────────────────────────────────

def get_all_sources():
    return [
        # Google Drive Excel file
        {
            "id":        "1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt",
            "sheet":     "GHIMS Incident Tracker",
            "type":      "excel",
            "start_row": 13
        },

        # ── To add more sources, just paste the link: ──────────────────────
        # Google Sheets link
        # {
        #     "url":       "https://docs.google.com/spreadsheets/d/SHEET_ID/edit",
        #     "sheet":     "Sheet1",
        #     "start_row": 13
        # },
        # Google Drive Excel/CSV link
        # {
        #     "url":       "https://drive.google.com/file/d/FILE_ID/view",
        #     "sheet":     "Sheet1",
        #     "start_row": 13
        # },
        # SharePoint / OneDrive / any direct URL
        # {
        #     "url":       "https://company.sharepoint.com/file.xlsx",
        #     "start_row": 13
        # },
    ]

# ── Buffer ────────────────────────────────────────────────────────────────────

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

# ── Routes ────────────────────────────────────────────────────────────────────

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
        master_sheet_obj, id_to_row, id_to_data, last_row = fetch_master(client)
        if master_sheet_obj is None:
            return jsonify({"status": "error", "message": "Could not read master sheet"}), 500

        logging.info(f"Master has {len(id_to_row)} IDs, last row: {last_row}")

        # ── STEP 1: Fetch sources → buffer ────────────────────────────────
        all_source_rows = {}

        if spreadsheet_id and sheet_name:
            logging.info(f"Targeted sync for {spreadsheet_id}/{sheet_name}")
            matched     = next((s for s in get_all_sources() if s.get("id") == spreadsheet_id), None)
            source_data = get_source_data(client, matched) if matched else \
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
                    source_key  = source.get("id") or source.get("url", f"source_{completed}")
                    save_to_buffer(source_data, source_key)
                    for row in source_data:
                        id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                        if not id_val:
                            continue
                        all_source_rows[id_val] = row
                        id_to_source[id_val]    = source_key
                    completed += 1
                    logging.info(f"Fetched {completed}/{len(sources)} sources")
                if completed < len(sources):
                    logging.info("Batch done, pausing 15 seconds...")
                    time.sleep(15)

        logging.info(f"Total source rows: {len(all_source_rows)}")

        # ── STEP 2: Validate buffer ───────────────────────────────────────
        buffer           = load_from_buffer()
        buffer_row_count = sum(v["count"] for v in buffer.values())
        logging.info(f"Buffer: {len(buffer)} sources, {buffer_row_count} rows")

        if buffer_row_count == 0:
            return jsonify({"status": "ok", "message": "Nothing to sync", "new": 0, "updates": 0, "removed": 0}), 200

        # ── STEP 3: Compare ───────────────────────────────────────────────
        new_rows      = []
        updates       = []
        ids_to_remove = []

        if spreadsheet_id and sheet_name:
            for id_val in list(id_to_row.keys()):
                if id_to_source.get(id_val) == spreadsheet_id:
                    if id_val not in all_source_rows:
                        ids_to_remove.append(id_val)
                        id_to_source.pop(id_val, None)
                    else:
                        source_row  = all_source_rows[id_val]
                        master_row  = id_to_data[id_val]
                        max_len     = max(len(source_row), len(master_row))
                        source_norm = source_row + [""] * (max_len - len(source_row))
                        master_norm = master_row + [""] * (max_len - len(master_row))
                        if source_norm != master_norm:
                            updates.append((id_val, source_row))
        else:
            for id_val in list(id_to_row.keys()):
                if id_val not in all_source_rows:
                    ids_to_remove.append(id_val)
                    id_to_source.pop(id_val, None)
                else:
                    source_row  = all_source_rows[id_val]
                    master_row  = id_to_data[id_val]
                    max_len     = max(len(source_row), len(master_row))
                    source_norm = source_row + [""] * (max_len - len(source_row))
                    master_norm = master_row + [""] * (max_len - len(master_row))
                    if source_norm != master_norm:
                        updates.append((id_val, source_row))

        for id_val, row in all_source_rows.items():
            if id_val not in id_to_row:
                new_rows.append(row)

        logging.info(f"New: {len(new_rows)} | Updates: {len(updates)} | Removals: {len(ids_to_remove)}")

        # ── STEP 4: Apply removals ────────────────────────────────────────
        if ids_to_remove:
            rows_to_delete = sorted(
                [id_to_row[id_val] for id_val in ids_to_remove if id_val in id_to_row],
                reverse=True
            )
            for sheet_row in rows_to_delete:
                master_sheet_obj.delete_rows(sheet_row)
                logging.info(f"Deleted row {sheet_row}")
            master_sheet_obj, id_to_row, id_to_data, last_row = fetch_master(client)

        # ── STEP 5: Apply updates ─────────────────────────────────────────
        if updates:
            max_cols = max(len(row) for _, row in updates)
            for id_val, row in updates:
                sheet_row = id_to_row.get(id_val)
                if sheet_row:
                    normalized = row + [""] * (max_cols - len(row))
                    master_sheet_obj.update(
                        f"A{sheet_row}",
                        [normalized],
                        value_input_option="RAW"
                    )
                    logging.info(f"Updated row {sheet_row} for ID {id_val}")

        # ── STEP 6: Append new rows ───────────────────────────────────────
        if new_rows:
            max_cols    = max(len(row) for row in new_rows)
            new_rows    = [row + [""] * (max_cols - len(row)) for row in new_rows]
            write_start = last_row + 1
            master_sheet_obj.update(
                f"A{write_start}",
                new_rows,
                value_input_option="RAW"
            )
            logging.info(f"Appended {len(new_rows)} rows starting at row {write_start}")

        # ── STEP 7: Clear buffer ──────────────────────────────────────────
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

@app.route("/dedupe", methods=["POST"])
def dedupe():
    try:
        client   = get_client()
        ss       = client.open_by_key(MASTER_SS_ID)
        sheet    = ss.worksheet(MASTER_SHEET)
        all_rows = sheet.get_all_values()

        seen_ids       = {}
        rows_to_delete = []

        for i, row in enumerate(all_rows):
            sheet_row = i + 1
            if sheet_row < MASTER_START:
                continue
            id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
            if not id_val:
                continue
            if id_val in seen_ids:
                rows_to_delete.append(sheet_row)
                logging.info(f"Duplicate: {id_val} at row {sheet_row}")
            else:
                seen_ids[id_val] = sheet_row

        for sheet_row in sorted(rows_to_delete, reverse=True):
            sheet.delete_rows(sheet_row)
            logging.info(f"Deleted duplicate at row {sheet_row}")

        return jsonify({
            "status":             "ok",
            "duplicates_removed": len(rows_to_delete),
            "unique_ids_kept":    len(seen_ids)
        }), 200

    except Exception as e:
        logging.error(f"Dedupe error: {e}")
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
        client = get_client()
        source = get_all_sources()[0]
        rows   = get_source_data(client, source)
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
