import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
import time
from flask import Flask, request, jsonify
import googleapiclient.discovery

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)

START_ROW    = 13
ID_COL       = 0
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"

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

def fetch_master(client):
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

def sync_excel_source(client, excel_file_id, sheet_name):
    """Convert Excel file to temp Google Sheet, read data, delete temp sheet."""
    temp_id = None
    try:
        drive_service = get_drive_service()

        # Copy and convert Excel to Google Sheets format
        copied_file = drive_service.files().copy(
            fileId=excel_file_id,
            body={
                "name": "TEMP_SYNC_COPY",
                "mimeType": "application/vnd.google-apps.spreadsheet"
            }
        ).execute()

        temp_id = copied_file["id"]
        logging.info(f"Created temp Google Sheet: {temp_id}")

        # Read data from temp sheet
        data = fetch_sheet(client, temp_id, sheet_name, START_ROW)
        logging.info(f"Read {len(data)} rows from temp sheet")

        return data

    except Exception as e:
        logging.error(f"Failed to sync Excel source {excel_file_id}: {e}")
        return []

    finally:
        # Always delete temp sheet even if something fails
        if temp_id:
            try:
                drive_service = get_drive_service()
                drive_service.files().delete(fileId=temp_id).execute()
                logging.info(f"Deleted temp sheet: {temp_id}")
            except Exception as e:
                logging.warning(f"Failed to delete temp sheet {temp_id}: {e}")

def chunk_sources(sources, size=50):
    for i in range(0, len(sources), size):
        yield sources[i:i + size]

def get_all_sources():
    return [
        {"id": "1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt", "sheet": "GHIMS Incident Tracker", "type": "excel"},
        # {"id": "1bIsyZ2cF-uAFI3fa7G8xGL98ZK8o5Ra5WWbV0FRLD88", "sheet": "Sheet1", "type": "gsheet"},
        # add all 600 sources here, specify type as "excel" or "gsheet"
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

        # Fetch sources
        all_source_rows = {}

        if spreadsheet_id and sheet_name:
            # Targeted sync — only fetch the one sheet that changed
            logging.info(f"Targeted sync for {spreadsheet_id}/{sheet_name}")

            # Check if this source is excel or gsheet
            source_type = "gsheet"
            for s in get_all_sources():
                if s["id"] == spreadsheet_id:
                    source_type = s.get("type", "gsheet")
                    break

            if source_type == "excel":
                source_data = sync_excel_source(client, spreadsheet_id, sheet_name)
            else:
                source_data = fetch_sheet(client, spreadsheet_id, sheet_name, START_ROW)

            for row in source_data:
                id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                if not id_val:
                    continue
                all_source_rows[id_val] = row
                id_to_source[id_val] = spreadsheet_id

        else:
            # Full sync — process all sources in batches of 50
            logging.info("Full sync — fetching all sources in batches...")
            sources   = get_all_sources()
            completed = 0

            for batch in chunk_sources(sources, size=50):
                for source in batch:
                    if source.get("type") == "excel":
                        source_data = sync_excel_source(client, source["id"], source["sheet"])
                    else:
                        source_data = fetch_sheet(client, source["id"], source["sheet"], START_ROW)

                    for row in source_data:
                        id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                        if not id_val:
                            continue
                        all_source_rows[id_val] = row
                        id_to_source[id_val] = source["id"]

                    completed += 1
                    logging.info(f"Fetched {completed}/{len(sources)} sources")

                # Pause between batches to avoid rate limits
                if completed < len(sources):
                    logging.info("Batch done, pausing 15 seconds...")
                    time.sleep(15)

        logging.info(f"Source rows fetched: {len(all_source_rows)}")

        new_rows      = []
        updates       = []
        ids_to_remove = []

        if spreadsheet_id and sheet_name:
            # Targeted sync — only check IDs that belong to this source
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
            # Full sync — check all IDs
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

        # Find new IDs
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

            # Re-read master after deletions
            _, master_data = fetch_master(client)
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

        return jsonify({
            "status": "ok",
            "new": len(new_rows),
            "updates": len(updates),
            "removed": len(ids_to_remove)
        }), 200

    except Exception as e:
        logging.error(f"Sync error: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/diagnose", methods=["GET"])
def diagnose():
    client = get_client()
    try:
        ss          = client.open_by_key("1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt")
        sheets      = ss.worksheets()
        sheet_names = [s.title for s in sheets]
        first_sheet = sheets[0]
        all_rows    = first_sheet.get_all_values()

        return jsonify({
            "all_tabs":                sheet_names,
            "first_tab_name":          first_sheet.title,
            "total_rows_in_first_tab": len(all_rows),
            "row_13":                  all_rows[12] if len(all_rows) >= 13 else "EMPTY",
            "row_1":                   all_rows[0]  if all_rows else "EMPTY"
        })
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/", methods=["GET"])
def health():
    return "OK", 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
