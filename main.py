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
DEPT_COL     = 3  # Column D = "Department" (0-indexed)
MASTER_SS_ID = "1BkMncGrq2o26CF77x7ppuyM0xlOEJSA6xNeu7T5CIHQ"
MASTER_SHEET = "Sheet7"
MASTER_START = 13
HEADER_ROW   = 12
BUFFER_FILE  = "/tmp/sync_buffer.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

MASTER_COLUMNS = [
    "Incident ID",
    "Incident State",
    "Branch",
    "Department",
    "No.",
    "Incident Description or Narration",
    "Incident Reported by",
    "Incident Documented By",
    "Date Logged",
    "Date Completed",
    "Duration",
    "Assigned To",
    "Status",
    "Comments"
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

# ── Column mapping ────────────────────────────────────────────────────────────

def normalize_header(h):
    return str(h).strip().lower()

def map_row_to_master_columns(row, header_row):
    """
    Align a source row to MASTER_COLUMNS using header names.
    Any column not found in the source is left blank.
    """
    source_map = {}
    for i, h in enumerate(header_row):
        key = normalize_header(h)
        if key:
            source_map[key] = row[i] if i < len(row) else ""

    return [source_map.get(normalize_header(col), "") for col in MASTER_COLUMNS]

# ── Parsers ───────────────────────────────────────────────────────────────────

def parse_excel(file_bytes, sheet_name, start_row, header_row_num=None):
    import openpyxl
    if header_row_num is None:
        header_row_num = HEADER_ROW

    workbook = openpyxl.load_workbook(file_bytes, data_only=True)

    if sheet_name and sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
    else:
        ws = workbook.active

    max_col  = ws.max_column
    headers  = []
    all_rows = []

    for i, row in enumerate(ws.iter_rows(min_row=header_row_num, max_col=max_col, values_only=True)):
        row_as_strings = [str(cell) if cell is not None else "" for cell in row]
        actual_row_num = header_row_num + i

        if actual_row_num == header_row_num:
            headers = row_as_strings
            continue

        if actual_row_num < start_row:
            continue

        if not any(cell.strip() for cell in row_as_strings):
            continue

        mapped = map_row_to_master_columns(row_as_strings, headers)
        all_rows.append(mapped)

    return all_rows

def parse_csv(content, start_row, header_row_num=None):
    import csv
    if header_row_num is None:
        header_row_num = HEADER_ROW

    try:
        text = content.decode("utf-8")
    except:
        text = content.decode("latin-1")

    reader   = csv.reader(text.splitlines())
    all_rows = list(reader)

    headers   = all_rows[header_row_num - 1] if len(all_rows) >= header_row_num else []
    data_rows = all_rows[start_row - 1:]
    data_rows = [row for row in data_rows if any(cell.strip() for cell in row)]
    mapped    = [map_row_to_master_columns(row, headers) for row in data_rows]
    return mapped

# ── Fetchers ──────────────────────────────────────────────────────────────────

def fetch_sheet(client, spreadsheet_id, sheet_name, start_row, header_row_num=None):
    if header_row_num is None:
        header_row_num = HEADER_ROW

    for attempt in range(3):
        try:
            ss       = client.open_by_key(spreadsheet_id)
            sheet    = ss.worksheet(sheet_name)
            all_rows = sheet.get_all_values()

            headers   = all_rows[header_row_num - 1] if len(all_rows) >= header_row_num else []
            data_rows = all_rows[start_row - 1:]
            data_rows = [row for row in data_rows if any(cell.strip() for cell in row)]
            mapped    = [map_row_to_master_columns(row, headers) for row in data_rows]
            return mapped

        except Exception as e:
            if "429" in str(e) or "quota" in str(e).lower():
                logging.warning(f"Rate limited, waiting 30s... (attempt {attempt + 1})")
                time.sleep(30)
            else:
                logging.warning(f"Failed to fetch {spreadsheet_id}/{sheet_name}: {e}")
                return []
    return []

def fetch_excel_source(file_id, sheet_name, start_row, header_row_num=None):
    try:
        drive_service = get_drive_service()
        request_obj   = drive_service.files().get_media(fileId=file_id)
        file_bytes    = io.BytesIO(request_obj.execute())
        return parse_excel(file_bytes, sheet_name, start_row, header_row_num)
    except Exception as e:
        logging.error(f"Failed to read Excel file {file_id}: {e}")
        return []

def detect_and_fetch_url(url, sheet_name, start_row, header_row_num=None):
    """Download file from any URL and auto-detect type."""

    # Google Sheets URL → use Sheets API
    if "docs.google.com/spreadsheets/d/" in url:
        match = re.search(r"spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        if match:
            spreadsheet_id = match.group(1)
            logging.info(f"Google Sheets URL detected, using API: {spreadsheet_id}")
            client = get_client()
            return fetch_sheet(client, spreadsheet_id, sheet_name or "Sheet1", start_row, header_row_num)
        else:
            logging.error(f"Could not extract ID from Google Sheets URL: {url}")
            return []

    # Google Drive file URL → use Drive API
    if "drive.google.com/file/d/" in url:
        match = re.search(r"file/d/([a-zA-Z0-9_-]+)", url)
        if match:
            file_id = match.group(1)
            logging.info(f"Google Drive URL detected, using Drive API: {file_id}")
            return fetch_excel_source(file_id, sheet_name, start_row, header_row_num)
        else:
            logging.error(f"Could not extract ID from Drive URL: {url}")
            return []

    # Everything else — download and auto-detect
    try:
        import requests

        logging.info(f"Downloading from URL: {url}")
        response = requests.get(url, timeout=60, allow_redirects=True)

        if response.status_code != 200:
            logging.error(f"Failed to download {url}: HTTP {response.status_code}")
            return []

        content_type = response.headers.get("Content-Type", "").lower()
        url_lower    = url.lower()
        logging.info(f"Content-Type: {content_type}")

        if any(x in content_type for x in ["excel", "spreadsheetml", "openxmlformats"]) or \
           any(url_lower.endswith(x) for x in [".xlsx", ".xls"]):
            logging.info("Detected: Excel")
            return parse_excel(io.BytesIO(response.content), sheet_name, start_row, header_row_num)

        elif "csv" in content_type or url_lower.endswith(".csv"):
            logging.info("Detected: CSV")
            return parse_csv(response.content, start_row, header_row_num)

        else:
            logging.info("Unknown type — trying Excel first, then CSV")
            try:
                result = parse_excel(io.BytesIO(response.content), sheet_name, start_row, header_row_num)
                if result:
                    return result
            except:
                pass
            try:
                result = parse_csv(response.content, start_row, header_row_num)
                if result:
                    return result
            except:
                pass
            logging.error(f"Could not parse file from {url}")
            return []

    except Exception as e:
        logging.error(f"Failed to fetch URL source {url}: {e}")
        return []

# ── All-sheets helpers ────────────────────────────────────────────────────────

def get_all_gsheet_names(client, spreadsheet_id):
    try:
        ss = client.open_by_key(spreadsheet_id)
        return [ws.title for ws in ss.worksheets()]
    except Exception as e:
        logging.error(f"Failed to get sheet names for {spreadsheet_id}: {e}")
        return []

def process_workbook_all_sheets(workbook, skip_sheets, start_row, header_row_num):
    """Read all tabs from an openpyxl workbook, map columns, return combined rows."""
    all_rows = []
    for name in workbook.sheetnames:
        if name in skip_sheets:
            logging.info(f"  Skipping tab '{name}'")
            continue
        try:
            ws       = workbook[name]
            max_col  = ws.max_column
            headers  = []
            tab_rows = []

            for i, row in enumerate(ws.iter_rows(min_row=header_row_num, max_col=max_col, values_only=True)):
                row_as_strings = [str(cell) if cell is not None else "" for cell in row]
                actual_row_num = header_row_num + i

                if actual_row_num == header_row_num:
                    headers = row_as_strings
                    continue

                if actual_row_num < start_row:
                    continue

                if not any(cell.strip() for cell in row_as_strings):
                    continue

                mapped = map_row_to_master_columns(row_as_strings, headers)
                tab_rows.append(mapped)

            logging.info(f"  Tab '{name}': {len(tab_rows)} rows")
            all_rows.extend(tab_rows)

        except Exception as e:
            logging.warning(f"  Failed to read tab '{name}': {e}")
            continue

    return all_rows

def fetch_all_sheets_gsheet_batch(client, spreadsheet_id, skip_sheets, start_row, header_row_num):
    """
    Fetch ALL tabs from a Google Sheet in a SINGLE batch API call.
    Much faster than fetching tab by tab — avoids timeout on large sheets.
    """
    all_rows = []
    try:
        ss          = client.open_by_key(spreadsheet_id)
        worksheets  = ss.worksheets()
        sheet_names = [ws.title for ws in worksheets if ws.title not in skip_sheets]
        logging.info(f"All-sheets batch: {len(sheet_names)} tabs in {spreadsheet_id}")

        if not sheet_names:
            return []

        ranges    = [f"'{name}'!A1:ZZ" for name in sheet_names]
        result    = ss.values_batch_get(ranges)
        responses = result.get("valueRanges", [])

        for i, name in enumerate(sheet_names):
            try:
                tab_values = responses[i].get("values", []) if i < len(responses) else []
                if not tab_values:
                    logging.info(f"  Tab '{name}': 0 rows (empty)")
                    continue

                headers   = tab_values[header_row_num - 1] if len(tab_values) >= header_row_num else []
                data_rows = tab_values[start_row - 1:]
                data_rows = [row for row in data_rows if any(str(cell).strip() for cell in row)]
                mapped    = [map_row_to_master_columns(row, headers) for row in data_rows]
                logging.info(f"  Tab '{name}': {len(mapped)} rows")
                all_rows.extend(mapped)
            except Exception as e:
                logging.warning(f"  Failed to process tab '{name}': {e}")
                continue

    except Exception as e:
        logging.error(f"Failed to batch fetch {spreadsheet_id}: {e}")

    return all_rows

def fetch_all_sheets(client, source):
    """Fetch data from ALL tabs in a source and combine into one list."""
    import openpyxl
    start_row      = source.get("start_row", 13)
    header_row_num = source.get("header_row", HEADER_ROW)
    skip_sheets    = source.get("skip_sheets", [])
    all_rows       = []

    if "url" in source:
        url = source["url"]

        # Google Sheets URL → batch fetch
        if "docs.google.com/spreadsheets/d/" in url:
            match = re.search(r"spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            if match:
                return fetch_all_sheets_gsheet_batch(
                    client, match.group(1), skip_sheets, start_row, header_row_num
                )
            else:
                logging.error(f"Could not extract ID from Google Sheets URL: {url}")
                return []

        # Google Drive Excel URL
        elif "drive.google.com/file/d/" in url:
            match = re.search(r"file/d/([a-zA-Z0-9_-]+)", url)
            if match:
                file_id       = match.group(1)
                drive_service = get_drive_service()
                file_bytes    = io.BytesIO(drive_service.files().get_media(fileId=file_id).execute())
                workbook      = openpyxl.load_workbook(file_bytes, data_only=True, read_only=True)
                logging.info(f"All-sheets: {len(workbook.sheetnames)} tabs in Drive file {file_id}")
                all_rows      = process_workbook_all_sheets(workbook, skip_sheets, start_row, header_row_num)
                workbook.close()
            else:
                logging.error(f"Could not extract ID from Drive URL: {url}")
            return all_rows

        # SharePoint / OneDrive / direct download URL
        else:
            import requests
            logging.info(f"Downloading file for all-sheets processing: {url}")
            response = requests.get(url, timeout=60, allow_redirects=True)
            if response.status_code == 200:
                content_type = response.headers.get("Content-Type", "").lower()
                if "html" in content_type:
                    logging.error("URL returned an HTML page — link may not be publicly accessible")
                    return []
                workbook = openpyxl.load_workbook(io.BytesIO(response.content), data_only=True, read_only=True)
                logging.info(f"All-sheets: {len(workbook.sheetnames)} tabs from URL")
                all_rows = process_workbook_all_sheets(workbook, skip_sheets, start_row, header_row_num)
                workbook.close()
            else:
                logging.error(f"Failed to download {url}: HTTP {response.status_code}")
            return all_rows

    elif source.get("type") == "excel":
        drive_service = get_drive_service()
        file_bytes    = io.BytesIO(drive_service.files().get_media(fileId=source["id"]).execute())
        workbook      = openpyxl.load_workbook(file_bytes, data_only=True, read_only=True)
        logging.info(f"All-sheets: {len(workbook.sheetnames)} tabs in Drive Excel {source['id']}")
        all_rows      = process_workbook_all_sheets(workbook, skip_sheets, start_row, header_row_num)
        workbook.close()
        return all_rows

    else:
        # Native Google Sheet by ID → batch fetch
        return fetch_all_sheets_gsheet_batch(
            client, source["id"], skip_sheets, start_row, header_row_num
        )

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

# ── Department splitter ───────────────────────────────────────────────────────

def split_by_department(client):
    """
    Read the master sheet and create/update one tab per department.
    Each tab gets the same header row (row 12) plus its filtered rows.
    Rows with no department go to 'Unassigned'.
    Tabs for departments that no longer exist are cleared but kept.
    """
    try:
        ss       = client.open_by_key(MASTER_SS_ID)
        master   = ss.worksheet(MASTER_SHEET)
        all_rows = master.get_all_values()

        if len(all_rows) < MASTER_START:
            logging.info("Department split: master sheet has no data rows yet")
            return

        header_row = all_rows[HEADER_ROW - 1]     # row 12 (index 11)
        data_rows  = all_rows[MASTER_START - 1:]   # row 13 onwards
        data_rows  = [row for row in data_rows if any(cell.strip() for cell in row)]

        # Group rows by department
        dept_map = {}
        for row in data_rows:
            dept = row[DEPT_COL].strip() if len(row) > DEPT_COL else ""
            if not dept:
                dept = "Unassigned"
            if dept not in dept_map:
                dept_map[dept] = []
            dept_map[dept].append(row)

        logging.info(f"Department split: {len(dept_map)} departments — {list(dept_map.keys())}")

        existing_titles = [ws.title for ws in ss.worksheets()]

        # Sheets we must never touch
        protected = {MASTER_SHEET}

        for dept, rows in dept_map.items():
            sheet_title = dept[:100]  # Google Sheets tab name limit

            # Skip if it would overwrite a protected sheet
            if sheet_title in protected:
                logging.warning(f"  Skipping '{sheet_title}' — name conflicts with a protected sheet")
                continue

            try:
                if sheet_title in existing_titles:
                    ws = ss.worksheet(sheet_title)
                    ws.clear()
                    logging.info(f"  Cleared existing tab '{sheet_title}'")
                else:
                    ws = ss.add_worksheet(
                        title=sheet_title,
                        rows=max(len(rows) + 20, 100),
                        cols=len(header_row)
                    )
                    logging.info(f"  Created new tab '{sheet_title}'")

                # Rows 1–11 blank (to match master layout), row 12 = header, row 13+ = data
                blank_rows = [[""] * len(header_row)] * (HEADER_ROW - 1)
                write_data = blank_rows + [header_row] + rows

                # Pad all rows to the same width
                max_cols   = max(len(r) for r in write_data)
                write_data = [r + [""] * (max_cols - len(r)) for r in write_data]

                ws.update("A1", write_data, value_input_option="RAW")
                logging.info(f"  Written {len(rows)} rows to tab '{sheet_title}'")

                # Small pause to avoid hitting write quota
                time.sleep(1)

            except Exception as e:
                logging.error(f"  Failed to update tab '{sheet_title}': {e}")
                continue

        logging.info("Department split complete.")

    except Exception as e:
        logging.error(f"Department split error: {e}")

# ── Source router ─────────────────────────────────────────────────────────────

def get_source_data(client, source):
    start_row      = source.get("start_row", 13)
    header_row_num = source.get("header_row", HEADER_ROW)

    if source.get("all_sheets"):
        return fetch_all_sheets(client, source)

    if "url" in source:
        return detect_and_fetch_url(source["url"], source.get("sheet"), start_row, header_row_num)
    elif source.get("type") == "excel":
        return fetch_excel_source(source["id"], source["sheet"], start_row, header_row_num)
    else:
        return fetch_sheet(client, source["id"], source["sheet"], start_row, header_row_num)

# ── Source list ───────────────────────────────────────────────────────────────

def get_all_sources():
    return [
        # ── Single sheet sources (headers on row 12, data from row 13) ────
        {
            "id":        "1pmtDOflpJ4ctaVLgp6BZs4zjv0Lhfvzt",  # Tema General Hospital
            "sheet":     "GHIMS Incident Tracker",
            "type":      "excel",
            "start_row": 13
        },
        {
            "id":        "1nAvvgPk0iMysrAx4TX30btTB-eUOcdAT",  # Adabraka Polyclinic
            "sheet":     "GHIMS Incident Tracker",
            "type":      "excel",
            "start_row": 13
        },
        {
            "id":        "1zZ3w_MeD86KqnbKOdNohDK1GLPI43yW7",  # Northern (needs sharing)
            "sheet":     "GHIMS Incident Tracker",
            "type":      "excel",
            "start_row": 13
        },
        {
            "id":        "1e_jbwTS8s7Ah8gw-gXnxhDI8IgV-JcIX",  # Weija, pantang
            "all_sheets":   True,
            "type":      "excel",
            "start_row": 13
        },

        # ── All-sheets sources ─────────────────────────────────────────────
        {
            "url":         "https://sptlgh-my.sharepoint.com/:x:/g/personal/solomon_odame_spagad_com/IQDiKJd6nTFtQ62uo7y7vyc6AcEGCKDR01APBsdVlt4tpRM?download=1",
            "all_sheets":  True,
            "start_row":   13,
            "skip_sheets": []
        },
        {
            "id":          "1NygRyFFrEOUYebY8ds52OPHaQBwV9WaAvsne5VWhg48",  # Abokobi Polyclinic
            "all_sheets":  True,
            "start_row":   2,
            "header_row":  1,
            "skip_sheets": []
        },
        {
            "id":          "1NNwHeVR3v-yWO9BwPR6cx0dDQjzvKttC7ZRcHRGW3HU",  # Ga east
            "all_sheets":  True,
            "start_row":   12,
            "header_row":  11,
            "skip_sheets": ["DEMO INCIDENTS"]
        },
        {
            "id":          "18LJkL85Nqn6GV6PR8jBriLACS7yf1aQN00WmQernG7I",  # Ga West
            "all_sheets":  True,
            "start_row":   11,
            "header_row":  10,
            "skip_sheets": ["DEMO INCIDENTS"]
        },
        {
            "id":          "1OHL2RusmPioUik8sdA_nQG4Hjr1d8fg4mnqpAvTzNlI",  # Ridge
            "all_sheets":  True,
            "start_row":   12,
            "header_row":  11,
            "skip_sheets": ["DEMO INCIDENTS"]
        },
        {
            "url":         "https://onedrive.live.com/:x:/g/personal/88741d10827ae2a0/IQAR1RU-_4mSRJdjMGcr-NjSAZWR6uFqao1CKPYmmUk0CD0?download=1",  # Ahafo region
            "all_sheets":  True,
            "start_row":   13,
            "skip_sheets": []
        },
        {
            "url":         "https://onedrive.live.com/:x:/g/personal/e3f5d5574dc1389a/IQBRQt77xPg3QYtx4xmgpKorAddI8gf7q1NPFVoTLw5VybM?download=1",  # Eastern region
            "all_sheets":  True,
            "start_row":   13,
            "skip_sheets": []
        },
        {
            "id":          "19DYqpFRD9bFlIa9UDQx45U5jSTZDsc0_7MedtuaaI9U",  # Western North
            "all_sheets":  True,
            "start_row":   13,
            "header_row":  12,
            "skip_sheets": ["[FACILITIES DEPARTMENT LISTINGS]", "Common Issues Tracker"]
        },

        # ── Templates for adding more sources ─────────────────────────────
        # Google Drive Excel — single sheet
        # {
        #     "id":        "FILE_ID",
        #     "sheet":     "GHIMS Incident Tracker",
        #     "type":      "excel",
        #     "start_row": 13
        # },
        # Google Sheet — single sheet
        # {
        #     "id":        "SHEET_ID",
        #     "sheet":     "Sheet1",
        #     "start_row": 13
        # },
        # Google Sheet — all tabs, custom header row
        # {
        #     "id":          "SHEET_ID",
        #     "all_sheets":  True,
        #     "start_row":   2,
        #     "header_row":  1,
        #     "skip_sheets": []
        # },
        # SharePoint / OneDrive / direct URL
        # {
        #     "url":         "https://sptlgh-my.sharepoint.com/...?download=1",
        #     "all_sheets":  True,
        #     "start_row":   13,
        #     "skip_sheets": ["Summary", "Dashboard"]
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

        # ── STEP 8: Split master by department ────────────────────────────
        logging.info("Running department split...")
        split_by_department(client)

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

@app.route("/split", methods=["POST"])
def split():
    """Manually trigger the department split without running a full sync."""
    try:
        client = get_client()
        split_by_department(client)
        return jsonify({"status": "ok", "message": "Department split complete"}), 200
    except Exception as e:
        logging.error(f"Split error: {e}")
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
