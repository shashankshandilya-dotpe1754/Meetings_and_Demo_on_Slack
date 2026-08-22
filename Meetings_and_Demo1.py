"""
Meetings_and_Demo.py — refactored

Pulls daily CSV exports (Zoho Analytics emailers) from Gmail, loads them into
Google Sheets, copies/reshapes data between sheets, and applies pivot-table
formatting. Functionally the same job list as the original script, but built
around a few reusable helpers + config tables instead of ~30 copy-pasted
blocks — ~450 lines instead of ~1500.

Fixes vs. the original run:
  1. "Date conversion failed: Invalid value ... for dtype 'str'"
     -> The `csv_df.iloc[:, 0] = pd.to_datetime(...)` block was dead code
        (nothing downstream read csv_df after this point — the real
        column-A date formatting is the Sheets API batchUpdate call right
        above it). Dropped it instead of papering over the crash.
  2. gspread DeprecationWarning on worksheet.update("A1", values)
     -> Every call now uses update(values=..., range_name=...).
  3. pandas DtypeWarning on mixed-type columns while reading the CSVs
     -> pd.read_csv(..., low_memory=False).
  4. The crash that ended the run:
        gspread.exceptions.APIError: [403] The caller does not have permission
     -> That's Google Sheets refusing the service account access to the
        "Enterprise Business and Meetings" spreadsheet — a sharing problem,
        not a code problem. Share that spreadsheet (Editor access) with the
        service-account email in your JSON key file
        (the "client_email" field) and it will resolve.
     -> Regardless, every sheet open/read/write below is now wrapped so one
        inaccessible sheet logs a warning and the rest of the jobs still
        run, instead of the whole script dying partway through.
  5. Off-by-one bug in the "MT / Meeting Reports" number formatting: the
     column list used to pick which Sheets columns to reformat was being
     fed straight into a 1-indexed column-letter function while the values
     themselves were selected with 0-indexed row access — this silently
     formatted the column *before* the intended one (e.g. AF instead of
     AG). Fixed by standardizing every helper on 0-indexed column input.
  6. The email/app-password and the local Windows path to the service
     account JSON were hardcoded in plaintext in the file. Moved to
     environment variables (.env) below — since that password was exposed
     in a shared file, rotate it (Google Account -> Security -> App
     passwords) even after making this change.
  7. PSDEMO_ROI_V2 was being uploaded to its sheet twice in a row inside
     the original script (once inside the email loop, once again right
     after) — now uploaded once.

Section 5 posts the "Summary" tab's report table to Slack as an image that
pixel-matches the sheet — cell colors, merged cells, borders, bold text —
by exporting the exact range as a PDF (Google renders that exactly as it
looks on screen) and converting it to a PNG.

Set these in a .env file next to this script (or your OS environment):
  GOOGLE_SERVICE_ACCOUNT_JSON=/path/to/service-account.json
  GMAIL_USER=you@yourcompany.com
  GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
  SLACK_BOT_TOKEN=xoxb-...
  SLACK_CHANNEL_ID=C0123456789

Extra packages beyond the original script: pymupdf, pillow, slack_sdk, requests
  pip install pymupdf pillow slack_sdk requests --break-system-packages
"""

import os
import re
import io
import imaplib
import email
from email.utils import parsedate_to_datetime
from datetime import datetime, time, timedelta
from io import BytesIO

import pandas as pd
import pytz
import gspread
import requests
import pymupdf
import fitz  # PyMuPDF — pip install pymupdf
from PIL import Image, ImageChops
from slack_sdk import WebClient  # pip install slack_sdk
from slack_sdk.errors import SlackApiError
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from gspread.utils import rowcol_to_a1, a1_to_rowcol
from gspread_formatting import (
    CellFormat, Color, TextFormat, NumberFormat, format_cell_range, batch_updater,
)
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

IST = pytz.timezone("Asia/Kolkata")
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

JSON_KEY_PATH = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]  # path to the service-account JSON file
EMAIL_USER = os.environ["GMAIL_USER"]
EMAIL_PASS = os.environ["GMAIL_APP_PASSWORD"]

creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_PATH, SCOPE)
gc = gspread.authorize(creds)
sheets_api = build("sheets", "v4", credentials=creds)

SLACK_BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]      # xoxb-... ; needs files:write scope
SLACK_CHANNEL_ID = os.environ["SLACK_CHANNEL_ID"]    # e.g. "C0123456789"

imap = imaplib.IMAP4_SSL("imap.gmail.com")
imap.login(EMAIL_USER, EMAIL_PASS)


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------

def log_ok(msg):
    print(f"✅ {msg}")


def log_warn(msg):
    print(f"⚠️ {msg}")


def last_filled_data_row(ws, start_a1, end_a1, header_rows, key_col_offset=0):
    """Row number (1-indexed, absolute) of the last row to include from
    start_a1:end_a1. The first `header_rows` rows are always kept as-is
    (title/label rows, which may contain merged cells that read back
    blank in the non-top-left cells). After that, data rows are kept only
    while the column at key_col_offset (0-indexed from start_a1's column)
    is non-blank — this trims trailing empty rows off the range."""
    start_row, _ = a1_to_rowcol(start_a1)
    values = ws.get(f"{start_a1}:{end_a1}")
    last_row = start_row + header_rows - 1
    for i, row in enumerate(values[header_rows:]):
        cell = row[key_col_offset] if key_col_offset < len(row) else ""
        if not str(cell).strip():
            break
        last_row = start_row + header_rows + i
    return last_row


def export_range_as_pdf(creds, spreadsheet_id, gid, range_a1):
    """Fetch a specific A1 range of a sheet tab as a PDF, rendered exactly
    as Google Sheets displays it (colors, merges, borders, bold text)."""
    access_token = creds.get_access_token().access_token
    params = {
        "format": "pdf", "gid": gid, "range": range_a1,
        "size": "A4", "portrait": "false", "fitw": "true",
        "gridlines": "false", "printtitle": "false", "sheetnames": "false",
        "pagenumbers": "false", "fzr": "false",
        "top_margin": "0.10", "bottom_margin": "0.10",
        "left_margin": "0.10", "right_margin": "0.10",
    }
    resp = requests.get(
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export",
        params=params, headers={"Authorization": f"Bearer {access_token}"},
    )
    resp.raise_for_status()
    return resp.content


def pdf_to_trimmed_png(pdf_bytes, zoom=3, pad=6):
    """Render page 1 of a PDF to a PNG (zoom=3 ~= 216 DPI) and crop away
    the surrounding white margin so it posts cleanly to Slack."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")

    bg = Image.new("RGB", img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    if bbox:
        left, top, right, bottom = bbox
        img = img.crop((
            max(left - pad, 0), max(top - pad, 0),
            min(right + pad, img.width), min(bottom + pad, img.height),
        ))

    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def post_image_to_slack(png_bytes, channel, filename, title=None, follow_up_text=None):
    """Uploads the image, then (if given) posts follow_up_text as its own
    message right after — Slack renders initial_comment ABOVE the file
    preview, so a separate message is what actually lands below it."""
    client = WebClient(token=SLACK_BOT_TOKEN)
    try:
        client.files_upload_v2(channel=channel, file=png_bytes, filename=filename, title=title)
        log_ok(f"Posted '{title or filename}' to Slack.")
        if follow_up_text:
            client.chat_postMessage(channel=channel, text=follow_up_text)
    except SlackApiError as e:
        log_warn(f"Slack upload failed: {e.response.get('error', e)}")


def open_sheet(spreadsheet_ref, tab, by="url"):
    """Open a worksheet. Returns None (with a warning) instead of crashing
    the whole run if the service account can't access it."""
    try:
        ss = gc.open_by_url(spreadsheet_ref) if by == "url" else gc.open_by_key(spreadsheet_ref)
        return ss.worksheet(tab)
    except gspread.exceptions.APIError as e:
        log_warn(f"Can't open '{tab}' ({spreadsheet_ref}): {e}. "
                 f"Share this sheet with the service account's client_email.")
    except Exception as e:
        log_warn(f"Can't open '{tab}' ({spreadsheet_ref}): {e}")
    return None


def fetch_csv_attachment(subject_keyword, start_hour, end_hour, filename=None, filename_suffix=".csv"):
    """Search today's inbox for a Zoho Analytics emailer and return
    (dataframe, received_datetime) for the first matching CSV attachment in
    the given IST time window, or (None, None) if nothing matched."""
    today_str = datetime.now(IST).strftime("%d-%b-%Y")
    imap.select("inbox")
    status, messages = imap.search(
        None, f'(FROM "notifications@zohoanalytics.in" SUBJECT "{subject_keyword}" ON {today_str})'
    )
    start_t, end_t = time(start_hour, 0), time(end_hour, 0)

    for msg_id in messages[0].split():
        _, msg_data = imap.fetch(msg_id, "(RFC822)")
        for response in msg_data:
            if not isinstance(response, tuple):
                continue
            msg = email.message_from_bytes(response[1])
            received = parsedate_to_datetime(msg["Date"]).astimezone(IST)
            if not (start_t <= received.time() <= end_t):
                continue
            for part in msg.walk():
                fname = part.get_filename()
                if "attachment" not in part.get("Content-Disposition", ""):
                    continue
                if filename and fname != filename:
                    continue
                if not filename and not (fname or "").endswith(filename_suffix):
                    continue
                df = pd.read_csv(BytesIO(part.get_payload(decode=True)), low_memory=False)
                df = df.replace([float("inf"), float("-inf")], "").fillna("")
                return df, received
    return None, None


def upload_dataframe(ws, df, mode="data_only", date_column=None):
    """mode='data_only' clears/rewrites rows below the header (keeps sheet
    formatting); mode='full' clears the whole sheet first.
    date_column: 0-based DataFrame column index to parse as a date before
    upload."""
    if ws is None or df is None:
        return
    df = df.copy()

    # Parse the given column as a date and write it back as a dd/mm/yyyy
    # string. USER_ENTERED input makes Sheets recognize that string as a
    # real date value (sortable, filterable) rather than plain text.
    # NOTE: writing a raw python/pandas datetime object here instead of a
    # string will fail — gspread has to JSON-encode the request body, and
    # datetime objects aren't JSON serializable.
    if date_column is not None and date_column < len(df.columns):
        col_name = df.columns[date_column]
        parsed = pd.to_datetime(df[col_name], dayfirst=True, errors="coerce")
        df[col_name] = parsed.dt.strftime("%d/%m/%Y").fillna("")

    values = [df.columns.tolist()] + df.values.tolist()
    if mode == "full":
        ws.clear()
    else:
        end_cell = rowcol_to_a1(len(df) + 1, len(df.columns))
        ws.batch_clear([f"A2:{end_cell}"])
    ws.update(values=values, range_name="A1", value_input_option="USER_ENTERED")


def to_number(val):
    s = str(val).strip() if val is not None else ""
    if s == "":
        return ""
    s = s.replace(",", "")
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return val


def col_letter(idx_0based):
    return rowcol_to_a1(1, idx_0based + 1).rstrip("0123456789")


def format_number_columns(ws, cols_0indexed, pattern="0", start_row=2):
    if ws is None:
        return
    fmt = CellFormat(numberFormat=NumberFormat(type="NUMBER", pattern=pattern))
    for col in cols_0indexed:
        letter = col_letter(col)
        format_cell_range(ws, f"{letter}{start_row}:{letter}", fmt)


def copy_columns(source_ws, dest_ws, col_indices, numeric_cols=(), dest_range="A1",
                  clear_range=None, value_input_option="USER_ENTERED",
                  format_cols=None, format_pattern="0"):
    """Read source_ws, pick out col_indices (0-based) from every row,
    coerce numeric_cols (0-based, in *source* terms), write into dest_ws.
    format_cols (0-based, in *output* terms) get number formatting."""
    if source_ws is None or dest_ws is None:
        return None
    data = source_ws.get_all_values()
    header = [data[0][i] for i in col_indices]
    rows = []
    for row in data[1:]:
        if not any(row):
            continue
        new_row = []
        for i in col_indices:
            val = row[i].strip() if i < len(row) else ""
            new_row.append(to_number(val) if i in numeric_cols else val)
        rows.append(new_row)
    final_data = [header] + rows

    if clear_range:
        dest_ws.batch_clear([clear_range])
    else:
        dest_ws.clear()
    dest_ws.update(values=final_data, range_name=dest_range, value_input_option=value_input_option)

    if format_cols:
        format_number_columns(dest_ws, format_cols, pattern=format_pattern)
    return final_data


def numeric_coerce_columns(data, col_indices):
    """In-place-ish float coercion of specific 0-indexed columns across a
    get_all_values() table; every other column/value is left untouched."""
    header, rows = data[0], [r[:] for r in data[1:]]
    for row in rows:
        for c in col_indices:
            if c < len(row):
                try:
                    row[c] = float(row[c].replace(",", "")) if row[c] else ""
                except (ValueError, AttributeError):
                    pass
    return [header] + rows


def push_full_sheet(dest_ws, data, number_cols=(), pattern="0"):
    if dest_ws is None or data is None:
        return
    dest_ws.clear()
    dest_ws.update(values=data, range_name="A1")
    if number_cols:
        format_number_columns(dest_ws, number_cols, pattern=pattern)


# --------------------------------------------------------------------------
# 1. Pull today's CSVs out of Gmail and load them into their sheets
# --------------------------------------------------------------------------

CSV_JOBS = [
    dict(name="FPI_ROI_V2", subject="Daily_Emailer (FPI_ROI_V2)", start=10, end=20,
         sheet="https://docs.google.com/spreadsheets/d/1x0zHzPNHtOEKtSKN1xNl6QvcYErKt0mUCExuYHDn3NM/edit?gid=172134459#gid=172134459",
         tab="Sheet1", mode="data_only", date_column=0, date_pattern="mmmm, yyyy"),
    dict(name="FPI_PL_ROI_V2", subject="Daily_Emailer (FPI_PL_ROI_V2)", start=10, end=20,
         sheet="https://docs.google.com/spreadsheets/d/15sFZYd7RNsJBHOTuMHxGf-ErK77N5UdTpku2A6DQZBw/edit?gid=1552552952#gid=1552552952",
         tab="Sheet1", mode="data_only"),
    dict(name="Meetings_ROI_V2", subject="Daily_Emailer (Meetings_ROI_V2)", start=10, end=20,
         filename="Meetings_ROI_V2.csv",
         sheet="https://docs.google.com/spreadsheets/d/1F7_urKM32CfjLSpqNanIJo3yCjxkwr6L-ffpwx8FzFk/edit?gid=2006961588#gid=2006961588",
         tab="Sheet1", mode="data_only", date_column=3, date_pattern="dd/mm/yyyy"),
    dict(name="PSDEMO_ROI_V2", subject="Daily_Emailer (PSDEMO_ROI_V2)", start=10, end=20,
         sheet="https://docs.google.com/spreadsheets/d/1jO7PQshTrzZ-2cfq6U-a--2uTBI0zI7VjSxrwVdcumI/edit?gid=1508333307#gid=1508333307",
         tab="Sheet1", mode="data_only", date_column=26, date_pattern="dd/mm/yyyy"),
    dict(name="Deal_H/F_Flag_V2", subject="Daily_Emailer (DEAL_H/F_Flag_ROI_V2)", start=10, end=20,
         sheet="https://docs.google.com/spreadsheets/d/1QbQ5UkLBMIWrqWhy5tvYGDyF_vfzF1g-BJ5zS0dCo5M/edit?gid=1340281528",
         tab="Sheet1", mode="full"),
    dict(name="Inbound Leads/Deals/Demos", subject="Daily_Emailer (Inbound Leads/Deals/SelfDemo)", start=10, end=19,
         filename="Inbound_Lead_-_Calls_-_Deals_-_Demos.csv",
         sheet="https://docs.google.com/spreadsheets/d/12KGodyEKsxsuBHKSfC3Lsn2Ud4PO99t30YrFDpDU-AU/edit?gid=1591501821#gid=1591501821",
         tab="Worksheet_Demos", mode="data_only"),
]

for job in CSV_JOBS:
    df, received = fetch_csv_attachment(job["subject"], job["start"], job["end"],
                                         filename=job.get("filename"))
    if df is None:
        log_warn(f"No valid email with a CSV found today for {job['name']} "
                 f"between {job['start']}:00 and {job['end']}:00.")
        continue
    log_ok(f"{job['name']}.csv uploaded from email received at {received.strftime('%I:%M %p')}")

    ws = open_sheet(job["sheet"], job["tab"])
    upload_dataframe(ws, df, mode=job["mode"], date_column=job.get("date_column"))
    if ws:
        log_ok(f"{job['name']}.csv file imported successfully to Google Sheet!")

    if ws and job.get("date_column") is not None:
        spreadsheet_id = job["sheet"].split("/d/")[1].split("/")[0]
        date_col = job["date_column"]
        date_pattern = job.get("date_pattern", "dd/mm/yyyy")
        sheets_api.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{
                "repeatCell": {
                    "range": {"sheetId": ws._properties["sheetId"], "startRowIndex": 1,
                              "startColumnIndex": date_col, "endColumnIndex": date_col + 1},
                    "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": date_pattern}}},
                    "fields": "userEnteredFormat.numberFormat",
                }
            }]},
        ).execute()
        log_ok(f"Column {col_letter(date_col)} formatted as '{date_pattern}'")


# --------------------------------------------------------------------------
# 2. Copy/reshape data between sheets
# --------------------------------------------------------------------------

try:
    src = open_sheet("https://docs.google.com/spreadsheets/d/15sFZYd7RNsJBHOTuMHxGf-ErK77N5UdTpku2A6DQZBw/edit#gid=2128889199", "WorkingSheet")
    dst = open_sheet("https://docs.google.com/spreadsheets/d/1x0zHzPNHtOEKtSKN1xNl6QvcYErKt0mUCExuYHDn3NM/edit#gid=1792673476", "FPI_PL_Import")
    copy_columns(src, dst, col_indices=[1, 0, 22, 23, 24, 36], numeric_cols={24}, format_cols=[4])
    log_ok("Data copied to FPI_ROI_V2 as FPI_PL_Import.")
except Exception as e:
    log_warn(f"FPI_PL_Import copy failed: {e}")

try:
    src = open_sheet("https://docs.google.com/spreadsheets/d/1QbQ5UkLBMIWrqWhy5tvYGDyF_vfzF1g-BJ5zS0dCo5M/edit#gid=852851606", "Worksheet")
    dst = open_sheet("https://docs.google.com/spreadsheets/d/1F7_urKM32CfjLSpqNanIJo3yCjxkwr6L-ffpwx8FzFk/edit#gid=2053354170", "DEAL_Import")
    copy_columns(src, dst, col_indices=[0, 1, 9], numeric_cols={0, 1, 9}, format_cols=[0, 1, 2])
    log_ok("DEAL_Import copied and formatted successfully to Meetings_ROI_V2 ✅")
except Exception as e:
    log_warn(f"DEAL_Import -> Meetings_ROI_V2 copy failed: {e}")

try:
    src = open_sheet("1QbQ5UkLBMIWrqWhy5tvYGDyF_vfzF1g-BJ5zS0dCo5M", "Worksheet", by="key")
    dst = open_sheet("1jO7PQshTrzZ-2cfq6U-a--2uTBI0zI7VjSxrwVdcumI", "DEAL_Import", by="key")
    copy_columns(src, dst, col_indices=[11, 48], numeric_cols={48}, clear_range="A1:B", format_cols=[1])
    log_ok("DEAL_Import copied and formatted successfully to PSDEMO_ROI_V2 ✅")
except Exception as e:
    log_warn(f"DEAL_Import -> PSDEMO_ROI_V2 copy failed: {e}")

FPI_COLS = [16, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39]  # Q, Z..AN
try:
    src = open_sheet("https://docs.google.com/spreadsheets/d/1x0zHzPNHtOEKtSKN1xNl6QvcYErKt0mUCExuYHDn3NM/edit#gid=374018598", "Worksheet")
    dst = open_sheet("https://docs.google.com/spreadsheets/d/1QbQ5UkLBMIWrqWhy5tvYGDyF_vfzF1g-BJ5zS0dCo5M/edit?gid=723379314#gid=723379314", "FPI_Import")
    copy_columns(src, dst, col_indices=FPI_COLS, numeric_cols=set(FPI_COLS), clear_range="A1:P")
    log_ok("FPI_Import copied and formatted successfully to Deal_H/F_Flag_V2 ✅")
except Exception as e:
    log_warn(f"FPI_Import copy failed: {e}")

try:
    src = open_sheet("https://docs.google.com/spreadsheets/d/1x0zHzPNHtOEKtSKN1xNl6QvcYErKt0mUCExuYHDn3NM/edit?gid=374018598#gid=374018598", "Worksheet")
    dst = open_sheet("https://docs.google.com/spreadsheets/d/1QbQ5UkLBMIWrqWhy5tvYGDyF_vfzF1g-BJ5zS0dCo5M/edit?gid=723379314#gid=723379314", "FPI_Import")
    copy_columns(src, dst, col_indices=[16, 3, 4, 5, 6, 7, 8, 9], numeric_cols={16, 3, 4, 5, 6, 7, 8, 9},
                 dest_range="S1", clear_range="S1:Z")
    log_ok("FPI_Import (With Collection Amount) copied and formatted successfully to Deal_H/F_Flag_V2 ✅")
except Exception as e:
    log_warn(f"FPI_Import (With Collection Amount) copy failed: {e}")


# ----------------------------------------------------------------------------------
# 3. Post the "PreSale Demo Summary" report table to Slack — pixel-matches the sheet
# ----------------------------------------------------------------------------------

try:
    ws = open_sheet(
        "https://docs.google.com/spreadsheets/d/1jO7PQshTrzZ-2cfq6U-a--2uTBI0zI7VjSxrwVdcumI/edit?pli=1&gid=93459055#gid=93459055",
        "Summary",
    )
    if ws is None:
        raise RuntimeError("Summary sheet not accessible")

    # B5:P7 = title/date/column-header rows (may contain merged cells);
    # from row 8 on, keep only rows whose Email column (B) is non-blank —
    # trims the trailing blank rows below the actual data.
    last_row = last_filled_data_row(ws, "B5", "R30", header_rows=3, key_col_offset=0)
    range_a1 = f"B5:R{last_row}"

    spreadsheet_id = "1jO7PQshTrzZ-2cfq6U-a--2uTBI0zI7VjSxrwVdcumI"
    pdf_bytes = export_range_as_pdf(creds, spreadsheet_id, ws.id, range_a1)
    png_bytes = pdf_to_trimmed_png(pdf_bytes)

    today_label = datetime.now(IST).strftime("%d %b %Y")
    yesterday_str = (datetime.now(IST) - timedelta(days=1)).strftime("%d/%m/%Y")
    post_image_to_slack(
        png_bytes, SLACK_CHANNEL_ID,
        filename=f"presale_demo_summary_{datetime.now(IST).strftime('%Y%m%d')}.png",
        title="PreSale Demo Summary",
        follow_up_text=f"Hi Team,\nPlease find the PreSales Demos conducted till yesterday i.e {yesterday_str}",
    )
except Exception as e:
    log_warn(f"Slack Summary report failed: {e}")
    
    
# --------------------------------------------------------------------------------------
# 4. Post the "Meeting and Demo Summary" report table to Slack — pixel-matches the sheet
# --------------------------------------------------------------------------------------

try:
    ws = open_sheet(
        "https://docs.google.com/spreadsheets/d/1F7_urKM32CfjLSpqNanIJo3yCjxkwr6L-ffpwx8FzFk/edit?pli=1&gid=481185227#gid=481185227",
        "Summary",
    )
    if ws is None:
        raise RuntimeError("Summary sheet not accessible")

    # B5:P7 = title/date/column-header rows (may contain merged cells);
    # from row 8 on, keep only rows whose Email column (B) is non-blank —
    # trims the trailing blank rows below the actual data.
    last_row = last_filled_data_row(ws, "B5", "M30", header_rows=3, key_col_offset=0)
    range_a1 = f"B5:M{last_row}"

    spreadsheet_id = "1F7_urKM32CfjLSpqNanIJo3yCjxkwr6L-ffpwx8FzFk"
    pdf_bytes = export_range_as_pdf(creds, spreadsheet_id, ws.id, range_a1)
    png_bytes = pdf_to_trimmed_png(pdf_bytes)

    today_label = datetime.now(IST).strftime("%d %b %Y")
    yesterday_str = (datetime.now(IST) - timedelta(days=1)).strftime("%d/%m/%Y")
    post_image_to_slack(
        png_bytes, SLACK_CHANNEL_ID,
        filename=f"meeting_demo_summary_{datetime.now(IST).strftime('%Y%m%d')}.png",
        title="Meeting and Demo Summary",
        follow_up_text=f"Hi Team,\nPlease find the Meetings and Demos conducted till yesterday i.e {yesterday_str}. Meeting Count are based on stages ('Approved and Closed'  , 'Completed'  , 'Demo Completed'  , 'Open'  , 'System Closed')",
    )
except Exception as e:
    log_warn(f"Slack Summary report failed: {e}")


imap.logout()
print("Done.")
