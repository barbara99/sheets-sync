import gspread
from google.oauth2.service_account import Credentials
import logging
import json
import os
from flask import Flask, request, jsonify

app = Flask(__name__)

logging.basicConfig(level=logging.INFO)

START_ROW    = 23
ID_COL       = 0
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_client():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds      = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return gspread.authorize(creds)

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

@app.route("/sync", methods=["POST"])
def sync():
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

        # If a specific sheet was edited, only fetch that one
        # Otherwise fetch all sources (fallback full sync)
        if spreadsheet_id and sheet_name:
            logging.info(f"Targeted sync for {spreadsheet_id}/{sheet_name}")
            sources_to_fetch = [{"id": spreadsheet_id, "sheet": sheet_name}]
        else:
            logging.info("Full sync — no specific sheet provided")
            sources_to_fetch = get_all_sources()

        # Fetch only the relevant source(s)
        all_source_rows = {}
        for source in sources_to_fetch:
            source_data = fetch_sheet(client, source["id"], source["sheet"], START_ROW)
            for row in source_data:
                id_val = row[ID_COL].strip() if len(row) > ID_COL else ""
                if not id_val:
                    continue
                all_source_rows[id_val] = row

        logging.info(f"Source rows fetched: {len(all_source_rows)}")

        new_rows      = []
        updates       = []
        ids_to_remove = []

        # Only check IDs that came from this source for removal
        source_ids = set(all_source_rows.keys())

        # Find IDs in master that came from this source but no longer exist
        if spreadsheet_id and sheet_name:
            # Only remove IDs that belong to this specific source
            # We know which IDs came from this source because we just fetched it
            # IDs in master that are NOT in the source fetch = deleted from this source
            master_ids_from_source = set()
            try:
                # Re-fetch to get full picture of what this source had before
                # We use the source fetch we already did
                master_ids_from_source = source_ids
            except:
                pass

            for id_val, index in id_to_index.items():
                if id_val in source_ids:
                    # This ID exists in source — check for updates
                    source_row  = all_source_rows[id_val]
                    master_row  = master_data[index]
                    max_len     = max(len(source_row), len(master_row))
                    source_norm = source_row + [""] * (max_len - len(source_row))
                    master_norm = master_row + [""] * (max_len - len(master_row))
                    if source_norm != master_norm:
                        updates.append((index, source_row))
        else:
            # Full sync — check all IDs for removal
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


def get_all_sources():
    # Full list of all 600 sources for fallback full sync
    return [
        {"id": "1PaYgXe2fzKkR-y-CXnei0RM2fQbCUlpVKoFAShjpX7w", "sheet": "Sheet1"},
        {"id": "1bIsyZ2cF-uAFI3fa7G8xGL98ZK8o5Ra5WWbV0FRLD88", "sheet": "Sheet1"},
        # add all 600 sources here
    ]
# Process in batches of 50
def chunk_sources(sources, size=50):
    for i in range(0, len(sources), size):
        yield sources[i:i + size]

@app.route("/", methods=["GET"])
def health():
    return "OK", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
