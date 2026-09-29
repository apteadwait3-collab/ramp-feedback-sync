"""
Normalize scraped feedback rows into the sheet schema, dedup against what's
already there, batch-write to Google Sheets, and format each sheet so the
feedback is actually readable (frozen header, sized columns, wrapped text,
banded rows).

Layout: one worksheet per event month ("May 2026", "June 2026", ...) rather
than one flat sheet — with 160+ events across many months each holding
several feedback rows, a single sheet became unreadable. An "Overview" tab
holds the startup demo row and a per-month row-count summary; "RunLog"
holds run history, same as before.
"""

import hashlib
import json
import os
import re
import time
from collections import defaultdict
from datetime import datetime, timezone

import gspread
from dateutil import parser as dateparser
from google.oauth2.service_account import Credentials

# Event columns, then the MSME respondent's identity fields (as requested),
# then the feedback itself. One row per respondent: every question the
# form asked is one "Question: Answer" line inside the single Feedback
# cell (source_client.format_feedback_text), newline-separated so Sheets'
# wrap-text formatting renders each question on its own visual line
# instead of one run-on paragraph. No separate "Rating" column — the old
# one only ever caught a rating when a form happened to use word-scale
# answers (Excellent/Good/...) rather than the numeric 1-5 scale most
# forms actually use, so it was blank almost everywhere; every answer,
# rating or otherwise, is already visible inline in Feedback now.
SCHEMA_COLUMNS = [
    "Event ID",
    "Component",
    "Event Name",
    "Event Date",
    "Month",
    "Enrollment ID",
    "Name",
    "Mobile Number",
    "Email",
    "District",
    "Feedback",
    "Scraped At",
]

# (width in pixels) — narrow for short fields, wide for the feedback text.
COLUMN_WIDTHS = [110, 160, 220, 110, 100, 110, 160, 120, 200, 120, 420, 170]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

OVERVIEW_SHEET = "Overview"
WEEKLY_REPORTS_SHEET = "Weekly Reports"  # historical archive from the old Monday schedule — left untouched, no longer written to
DAILY_REPORTS_SHEET = "Daily Reports"
ADMIN_CHECK_LOG_SHEET = "Post-Event Admin Log"

ADMIN_CHECK_LOG_HEADER = [
    "Run Timestamp", "Event ID", "Event Name", "Event Date", "Registered",
    "Attended", "Attendance Sheet", "Event Photos", "Status", "Details",
]

# Resolved against each live worksheet header at runtime. ``Feedback`` is the
# field present in the current workbook; the earlier names make a future
# explicit post-event column take precedence without silently reading the
# wrong field.
POST_EVENT_COMPLETION_COLUMNS = (
    "Post-event Data",
    "Post-event Data Entry",
    "Data Entry Status",
    "Feedback",
)

HEADER_COLOR = {"red": 0.11, "green": 0.27, "blue": 0.53}  # dark blue
BAND_COLOR = {"red": 0.94, "green": 0.96, "blue": 1.0}  # very light blue


# A single Sheets cell caps at 50,000 chars. This class of bug (a per-item
# join silently growing past that under real-world volume) has bitten this
# project once already — write_run_log's error join crashed live at 218
# errors — so every place that concatenates a variable number of items into
# one cell caps its length through this one helper rather than trusting the
# input to stay small.
CELL_CHAR_LIMIT = 45000
ERROR_SUMMARY_CHAR_LIMIT = 2000
RUN_LOG_FILE = "output/run.log"


def _capped(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + f"… [truncated, {len(text)} chars total]"


def _normalized_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _find_header(header: list[str], candidates: tuple[str, ...], required: bool = True):
    by_normalized = {_normalized_header(value): (index, value) for index, value in enumerate(header)}
    for candidate in candidates:
        match = by_normalized.get(_normalized_header(candidate))
        if match:
            return match
    if required:
        raise ValueError(
            f"Sheet schema is missing one of {candidates}; actual headers: {header}"
        )
    return None


def _with_retry(fn, *args, retries=4, base_delay=2, **kwargs):
    """The Sheets API occasionally returns a transient 5xx under sustained
    use (confirmed live: a bare 500 mid-run) — retry those with backoff.
    Client errors (403 permission, 400 bad request, ...) are real problems,
    not transient, so they're raised immediately instead of being retried."""
    last_err = None
    for attempt in range(retries):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code
            if status < 500:
                raise
            last_err = e
            time.sleep(base_delay * (2**attempt))
    raise last_err


def derive_month(event_date: str) -> str:
    """Best-effort "Month YYYY" from whatever date string the API returns —
    this is also the worksheet name each row gets routed to.

    Confirmed live against real mssidc/getialist responses: EventDate comes
    back as "MM/DD/YYYY HH:MM:SS" (American, month-first) — e.g.
    "05/09/2026 00:00:00" means May 9, not September 5. An earlier version
    of this function assumed day-first (dayfirst=True), which silently
    swapped month/day whenever the day-of-month was <=12 (so 05/09 read as
    September, 05/12 as December) — confirmed by cross-checking every
    distinct Event Date against its derived Month on a live run, where
    every event was actually in May. Fixed to month-first; the ASP.NET
    `/Date(epoch_ms)/` wrapper is still handled in case a different
    endpoint ever uses it, and the column is left blank (routed to an
    "Unknown" sheet) rather than guessing if neither form parses.
    """
    if not event_date:
        return ""
    s = str(event_date).strip()

    m = re.match(r"/Date\((\d+)\)/", s)
    if m:
        try:
            dt = datetime.fromtimestamp(int(m.group(1)) / 1000, tz=timezone.utc)
            return dt.strftime("%B %Y")
        except (ValueError, OSError):
            return ""

    try:
        dt = dateparser.parse(s, dayfirst=False)
        return dt.strftime("%B %Y")
    except (ValueError, OverflowError, dateparser.ParserError):
        return ""


def normalize_row(event: dict, respondent: dict, feedback_text: str) -> dict:
    """Map one (event, respondent, feedback) pair onto the fixed schema —
    one row per respondent.

    Field provenance (see source_client.py docstring for confidence levels):
    - Event ID / Component / Event Name / Event Date: from the IA event
      list (mssidc/getialist, the endpoint behind the "/view_event"
      screen), fields confirmed via VieweventComponent's own
      displayedColumns + template bindings in the app bundle.
    - Enrollment ID / Name / Mobile Number / Email / District: from the
      per-event respondent list (industryassociation/viewfeedback) —
      fields confirmed via EventfeedbacklistComponent's own Excel-export
      mapping in the app bundle.
    - Feedback: every question this respondent answered, newline-joined —
      from source_client.answers_by_question + format_feedback_text.
    """
    event_id = event.get("eventId") or event.get("uniqueEventId") or ""
    component = event.get("componentName", "")
    event_name = event.get("eventName", "")
    event_date = event.get("eventDate", "")

    return {
        "Event ID": event_id,
        "Component": component,
        "Event Name": event_name,
        "Event Date": event_date,
        "Month": derive_month(event_date),
        "Enrollment ID": respondent.get("enrollmentId") or respondent.get("EnrollmentId") or "",
        "Name": respondent.get("name") or respondent.get("udyamRegistrationNumber") or "",
        "Mobile Number": respondent.get("ownersMobileNumber", ""),
        "Email": respondent.get("email", ""),
        "District": respondent.get("districtName", ""),
        "Feedback": feedback_text,
        "Scraped At": datetime.now(timezone.utc).isoformat(),
    }


def dedup_key(row: dict) -> str:
    """Event ID + Enrollment ID is a genuinely unique key for one person's
    one submission to one event — falls back to a hash of name+feedback
    when Enrollment ID is missing (PROMPT.md's suggested fallback)."""
    if row.get("Enrollment ID"):
        basis = f"{row['Event ID']}|{row['Enrollment ID']}"
    else:
        basis = "|".join([str(row["Event ID"]), str(row["Name"]), str(row["Feedback"])])
    return hashlib.sha256(basis.encode("utf-8")).hexdigest()


class SheetsSync:
    def __init__(self, service_account_json: str, sheet_id: str):
        creds = Credentials.from_service_account_file(service_account_json, scopes=SCOPES)
        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open_by_key(sheet_id)
        self._worksheets = {}  # sheet name -> gspread.Worksheet, cached per run

    def _get_or_create(self, name: str):
        if name in self._worksheets:
            return self._worksheets[name]
        try:
            ws = self.spreadsheet.worksheet(name)
        except gspread.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(name, rows=1000, cols=len(SCHEMA_COLUMNS))
        self._worksheets[name] = ws
        return ws

    def _existing_keys(self, ws) -> set:
        values = _with_retry(ws.get_all_values)
        if not values or not values[0]:
            return set()
        header, *rows = values
        keys = set()
        for r in rows:
            row_dict = dict(zip(header, r))
            if all(c in row_dict for c in SCHEMA_COLUMNS):
                keys.add(dedup_key(row_dict))
        return keys

    def _ensure_header(self, ws):
        # get_all_values() returns [[]] (not []) for a genuinely blank
        # sheet, which is still truthy — check the first row instead.
        values = _with_retry(ws.get_all_values)
        if not values or not values[0]:
            _with_retry(ws.append_row, SCHEMA_COLUMNS, value_input_option="USER_ENTERED")
            self._apply_formatting(ws)

    def _apply_formatting(self, ws):
        """One-time formatting applied when a month sheet's header is first
        written: frozen header row, bold/colored header, sized columns,
        wrapped feedback text, and banded rows so it reads as a report."""
        n_cols = len(SCHEMA_COLUMNS)
        last_col_letter = gspread.utils.rowcol_to_a1(1, n_cols).rstrip("1")
        sheet_id = ws.id

        _with_retry(ws.freeze, rows=1)

        _with_retry(ws.format, f"A1:{last_col_letter}1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": HEADER_COLOR,
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        })

        feedback_col = SCHEMA_COLUMNS.index("Feedback") + 1
        feedback_col_letter = gspread.utils.rowcol_to_a1(1, feedback_col).rstrip("1")
        _with_retry(ws.format, f"{feedback_col_letter}2:{feedback_col_letter}", {
            "wrapStrategy": "WRAP",
            "verticalAlignment": "TOP",
        })

        requests = []
        for i, width in enumerate(COLUMN_WIDTHS):
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1,
                    },
                    "properties": {"pixelSize": width},
                    "fields": "pixelSize",
                }
            })
        # ws.clear() wipes cell values but not formatting/banding, so a
        # sheet that's been cleared-and-reused already has a banded range
        # from a prior run — re-adding one errors out. Only add it if this
        # sheet doesn't already have one.
        meta = _with_retry(self.spreadsheet.fetch_sheet_metadata)
        sheet_meta = next(s for s in meta["sheets"] if s["properties"]["sheetId"] == sheet_id)
        if not sheet_meta.get("bandedRanges"):
            requests.append({
                "addBanding": {
                    "bandedRange": {
                        "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 5000},
                        "rowProperties": {
                            "firstBandColor": {"red": 1, "green": 1, "blue": 1},
                            "secondBandColor": BAND_COLOR,
                        },
                    }
                }
            })
        if requests:
            _with_retry(self.spreadsheet.batch_update, {"requests": requests})

    def write_demo_row(self):
        """Append a note to the Overview sheet immediately, before scraping
        starts, so there's a visible sign of life right away."""
        ws = self._get_or_create(OVERVIEW_SHEET)
        values = _with_retry(ws.get_all_values)
        if not values or not values[0]:
            _with_retry(ws.append_row, ["Month", "Feedback Rows", "Last Updated"], value_input_option="USER_ENTERED")
            _with_retry(ws.format, "A1:C1", {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": HEADER_COLOR,
            })
            _with_retry(ws.freeze, rows=1)
        now = datetime.now(timezone.utc).isoformat()
        _with_retry(ws.append_row, [
            "Automation started",
            "",
            f"Scrape started at {now}. Per-month sheets will appear shortly.",
        ], value_input_option="USER_ENTERED")

    def sync(self, rows: list[dict]) -> int:
        """Dedup and batch-append new rows, routed to a per-month worksheet
        based on each row's Month (blank Month routes to "Unknown"). Returns
        the total count of rows actually added across all months touched."""
        by_month = defaultdict(list)
        for row in rows:
            by_month[row["Month"] or "Unknown"].append(row)

        total_added = 0
        for month, month_rows in by_month.items():
            ws = self._get_or_create(month)
            self._ensure_header(ws)
            existing = self._existing_keys(ws)

            new_rows = []
            seen_in_batch = set()
            for row in month_rows:
                key = dedup_key(row)
                if key in existing or key in seen_in_batch:
                    continue
                seen_in_batch.add(key)
                new_rows.append([row[c] for c in SCHEMA_COLUMNS])

            if new_rows:
                _with_retry(ws.append_rows, new_rows, value_input_option="USER_ENTERED")
            total_added += len(new_rows)
        return total_added

    def read_post_event_rows(self, events: list[dict]) -> dict[str, list[dict]]:
        """Read the existing Sheet field that tracks post-event completion.

        Only month worksheets needed by ``events`` are fetched. Headers are
        resolved case/punctuation-insensitively instead of assuming column
        positions. The returned rows contain the shared participant identity
        fields plus the discovered completion value for post_event_checks.py.
        A schema mismatch fails loudly rather than turning every participant
        into a misleading pending item.
        """
        wanted_months = {derive_month(event.get("eventDate", "")) for event in events}
        wanted_months.discard("")
        available = {worksheet.title for worksheet in self.spreadsheet.worksheets()}
        month_titles = sorted(wanted_months & available)
        if not month_titles:
            return {}

        ranges = [f"'{title.replace(chr(39), chr(39) * 2)}'!A:Z" for title in month_titles]
        response = _with_retry(self.spreadsheet.values_batch_get, ranges)
        rows_by_event = defaultdict(list)

        for title, value_range in zip(month_titles, response.get("valueRanges", [])):
            values = value_range.get("values") or []
            if not values:
                continue
            header = values[0]
            event_col, _ = _find_header(header, ("Event ID", "EventId"))
            name_col, _ = _find_header(header, ("Name", "Participant Name"))
            mobile_col, _ = _find_header(header, ("Mobile Number", "Mobile"))
            email_col, _ = _find_header(header, ("Email", "Email Address"))
            completion_col, completion_name = _find_header(
                header,
                POST_EVENT_COMPLETION_COLUMNS,
            )

            def value(row, index):
                return row[index] if index < len(row) else ""

            for row in values[1:]:
                event_id = str(value(row, event_col)).strip()
                if not event_id:
                    continue
                rows_by_event[event_id].append({
                    "name": value(row, name_col),
                    "mobile": value(row, mobile_col),
                    "email": value(row, email_col),
                    "completion_column": completion_name,
                    "completion_value": value(row, completion_col),
                    "worksheet": title,
                })

        return dict(rows_by_event)

    def replace_month(self, month: str, rows: list[dict]) -> int:
        """Clear one worksheet and rewrite it from scratch with a fresh row
        set — unlike sync(), which only appends rows not already present,
        this discards whatever was already in that tab. Used only by the
        one-time in-place format migration (migrate_feedback_format.py
        --in-place), never by the ongoing incremental sync."""
        ws = self._get_or_create(month)
        _with_retry(ws.clear)
        self._ensure_header(ws)

        seen = set()
        new_rows = []
        for row in rows:
            key = dedup_key(row)
            if key in seen:
                continue
            seen.add(key)
            new_rows.append([row[c] for c in SCHEMA_COLUMNS])

        if new_rows:
            _with_retry(ws.append_rows, new_rows, value_input_option="USER_ENTERED")
        return len(new_rows)

    def update_overview(self):
        """Refresh the Overview tab's per-month row counts. Cheap enough to
        call once at the end of a run rather than after every event."""
        ws = self._get_or_create(OVERVIEW_SHEET)
        month_sheets = [
            s for s in self.spreadsheet.worksheets()
            if s.title not in (
                OVERVIEW_SHEET,
                "RunLog",
                WEEKLY_REPORTS_SHEET,
                DAILY_REPORTS_SHEET,
                ADMIN_CHECK_LOG_SHEET,
            )
        ]
        rows = [["Month", "Feedback Rows", "Last Updated"]]
        now = datetime.now(timezone.utc).isoformat()
        for s in sorted(month_sheets, key=lambda s: s.title):
            count = max(0, len(_with_retry(s.get_all_values)) - 1)
            rows.append([s.title, count, now])
        ws.clear()
        _with_retry(ws.append_rows, rows, value_input_option="USER_ENTERED")
        _with_retry(ws.format, "A1:C1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": HEADER_COLOR,
        })

    def write_run_log(self, events_scanned: int, new_rows: int, errors: list[str]):
        """Write one summary row to the RunLog tab. The cell only ever gets
        a short preview (first 5 errors, capped hard at
        ERROR_SUMMARY_CHAR_LIMIT) — the full, untruncated list always goes
        to the local run.log file first, unconditionally, so nothing is
        lost even if the Sheets write itself fails."""
        timestamp = datetime.now(timezone.utc).isoformat()

        error_summary = ""
        if errors:
            os.makedirs(os.path.dirname(RUN_LOG_FILE), exist_ok=True)
            with open(RUN_LOG_FILE, "a") as f:
                f.write(f"\n=== Run at {timestamp} ({len(errors)} errors) ===\n")
                for e in errors:
                    f.write(f"{e}\n")

            preview = "; ".join(errors[:5])
            if len(errors) > 5:
                preview += f" …and {len(errors) - 5} more (see {RUN_LOG_FILE})"
            error_summary = _capped(preview, ERROR_SUMMARY_CHAR_LIMIT)

        try:
            log_ws = self.spreadsheet.worksheet("RunLog")
        except gspread.WorksheetNotFound:
            log_ws = self.spreadsheet.add_worksheet("RunLog", rows=1000, cols=4)
            _with_retry(log_ws.append_row, ["Timestamp", "Events Scanned", "New Rows Added", "Errors"])

        _with_retry(log_ws.append_row, [
            timestamp,
            events_scanned,
            new_rows,
            error_summary,
        ], value_input_option="USER_ENTERED")

    def write_admin_check_log(self, entries: list[dict]):
        """Append one row per event checked by admin_completion_check.py's
        noon run — a full audit trail ("every log about whatever is
        happening") separate from the feedback-sync RunLog, since this
        checks entirely different portal data (registration/attendance
        counts + document uploads, not feedback forms).

        Each entry: {event_id, event_name, event_date, registered, attended,
        has_attendance_sheet, has_event_photos, status, details}.
        """
        if not entries:
            return
        try:
            ws = self.spreadsheet.worksheet(ADMIN_CHECK_LOG_SHEET)
        except gspread.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                ADMIN_CHECK_LOG_SHEET, rows=1000, cols=len(ADMIN_CHECK_LOG_HEADER)
            )
            _with_retry(ws.append_row, ADMIN_CHECK_LOG_HEADER)
            last_col = gspread.utils.rowcol_to_a1(1, len(ADMIN_CHECK_LOG_HEADER)).rstrip("1")
            _with_retry(ws.format, f"A1:{last_col}1", {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": HEADER_COLOR,
            })
            _with_retry(ws.freeze, rows=1)

        timestamp = datetime.now(timezone.utc).isoformat()
        rows = [[
            timestamp,
            e.get("event_id", ""),
            e.get("event_name", ""),
            e.get("event_date", ""),
            e.get("registered", ""),
            e.get("attended", ""),
            "Yes" if e.get("has_attendance_sheet") else "No",
            "Yes" if e.get("has_event_photos") else "No",
            e.get("status", ""),
            _capped(e.get("details", ""), CELL_CHAR_LIMIT),
        ] for e in entries]
        _with_retry(ws.append_rows, rows, value_input_option="USER_ENTERED")

    def write_daily_report(self, run_date_label: str, month_label: str, summary_events: list[dict],
                            pending_items: list[dict], errors: list[str], html_body: str, plain_body: str):
        """Upsert one row in Daily Reports: that day's event
        numbers, the post-event pending-items count, plus an exact archive
        of the digest email that was sent, so a past day's report can be
        looked up later without digging through inboxes. Both the
        structured numbers (Summary JSON / Pending Items JSON) and the
        rendered email (Email HTML / Email Text) are kept, each capped
        through _capped like every other variable-length cell in this
        file."""
        timestamp = datetime.now(timezone.utc).isoformat()
        total_events = len(summary_events)
        total_enrolled = sum(e["enrolled"] for e in summary_events if e["enrolled"] is not None)
        total_attended = sum(e["attended"] for e in summary_events if e["attended"] is not None)
        total_feedback = sum(e["feedbackCount"] for e in summary_events)

        error_summary = "; ".join(errors[:5])
        if len(errors) > 5:
            error_summary += f" …and {len(errors) - 5} more"
        error_summary = _capped(error_summary, ERROR_SUMMARY_CHAR_LIMIT)

        header = [
            "Timestamp", "Date", "Month", "Events", "Registered", "Attended",
            "Feedback", "Pending Events", "Errors", "Summary JSON",
            "Pending Items JSON", "Email HTML", "Email Text",
        ]
        try:
            ws = self.spreadsheet.worksheet(DAILY_REPORTS_SHEET)
        except gspread.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(DAILY_REPORTS_SHEET, rows=1000, cols=len(header))
            _with_retry(ws.append_row, header)
            last_col = gspread.utils.rowcol_to_a1(1, len(header)).rstrip("1")
            _with_retry(ws.format, f"A1:{last_col}1", {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": HEADER_COLOR,
            })
            _with_retry(ws.freeze, rows=1)

        report_row = [
            timestamp, run_date_label, month_label, total_events, total_enrolled,
            total_attended, total_feedback, len(pending_items), error_summary,
            _capped(json.dumps(summary_events), CELL_CHAR_LIMIT),
            _capped(json.dumps(pending_items), CELL_CHAR_LIMIT),
            _capped(html_body, CELL_CHAR_LIMIT),
            _capped(plain_body, CELL_CHAR_LIMIT),
        ]

        # Idempotent archive: a second run on the same IST date replaces the
        # existing report row rather than appending a duplicate.
        values = _with_retry(ws.get_all_values)
        existing_row = None
        if values and values[0]:
            date_index = values[0].index("Date") if "Date" in values[0] else None
            if date_index is not None:
                for row_number, row in enumerate(values[1:], start=2):
                    if date_index < len(row) and row[date_index] == run_date_label:
                        existing_row = row_number
                        break

        if existing_row is None:
            _with_retry(ws.append_row, report_row, value_input_option="USER_ENTERED")
        else:
            last_col = gspread.utils.rowcol_to_a1(1, len(header)).rstrip("1")
            _with_retry(
                ws.update,
                [report_row],
                range_name=f"A{existing_row}:{last_col}{existing_row}",
                value_input_option="USER_ENTERED",
            )
