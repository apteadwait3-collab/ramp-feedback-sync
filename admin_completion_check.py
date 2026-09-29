"""
Noon entrypoint: a second, independent daily check from the feedback-sync
job in daily_report.py (which now runs Monday-only — see README).

Every day at 12:00 IST this opens the same manual RAMP login as every other
entrypoint here, then checks exactly YESTERDAY's events (not today's — an
event happening today hasn't had time for post-event admin work yet, and
not older backlog either — see below) for whether the team has done the
post-event portal admin work:

- Registered MSME data entered (enrolled-MSME count > 0)
- Attended MSME data entered (attended count > 0)
- Attendance sheet uploaded (View Document, EventAttendance)
- Event photos uploaded (View Document, EventImages)

This does NOT check MSME feedback forms — that's daily_report.py's job,
unchanged, against the Sheet's Feedback column. This check is purely
against live portal counts/documents, no Sheet cross-referencing.

One email is sent on every run, no matter what: if no events happened
yesterday, a short "no events" email; otherwise a digest — events missing
anything get a reminder section listing exactly what's missing, events
that are fully complete get a section with the actual numbers. Every event
checked (and a heartbeat row on a no-events day) is also logged as its own
row in the "Post-Event Admin Log" Sheet tab (sheets_sync.write_admin_check_log),
so there's a full daily audit trail regardless of outcome.

Deliberately no retry-until-fixed/backlog tracking (unlike
post_event_checks.py): scope is always exactly yesterday, once. An event
that's still incomplete two days later won't be re-flagged by this job —
that's a explicit product decision, not an oversight.
"""

import argparse
import os
import sys
import time
from datetime import datetime, timedelta

if os.name == "nt":
    import msvcrt
else:
    import fcntl

from dotenv import load_dotenv

import auth
import emailer
from notifier import notify
from post_event_checks import IST, check_view_documents, is_cancelled, parse_event_date
from sheets_sync import SheetsSync
from source_client import PersistentFailure, RampClient

OUTPUT_DIR = "output"
RUN_LOCK_FILE = os.path.join(OUTPUT_DIR, "admin_check.lock")


def event_happened_on(event: dict, target_date) -> bool:
    event_date = parse_event_date(event.get("eventDate", ""))
    return bool(event_date and event_date == target_date and not is_cancelled(event))


def check_event(client: RampClient, event: dict, errors: list) -> dict:
    """Run all four checks for one event. Returns a result dict always —
    a failed sub-check is reported as an error string rather than raising,
    so one bad event never stops the run."""
    event_id = str(event.get("eventId") or event.get("uniqueEventId") or "?")
    unique_event_id = event.get("uniqueEventId")
    result = {
        "event_id": event_id,
        "event_name": event.get("eventName", ""),
        "event_date": event.get("eventDate", ""),
        "registered": None,
        "attended": None,
        "has_attendance_sheet": None,
        "has_event_photos": None,
        "missing": [],
        "check_failed": False,
    }

    if not unique_event_id:
        result["check_failed"] = True
        result["missing"] = ["missing uniqueEventId, could not check"]
        errors.append(f"Event {event_id}: missing uniqueEventId, skipping admin check")
        return result

    try:
        result["registered"] = client.get_enrolled_count(unique_event_id)
        if result["registered"] == 0:
            result["missing"].append("Registered MSME data not entered")
    except PersistentFailure as exc:
        result["check_failed"] = True
        result["missing"].append("Registered MSME check failed")
        errors.append(f"Event {event_id}: registered-count check failed ({exc})")

    try:
        result["attended"] = client.get_attended_count(unique_event_id)
        if result["attended"] == 0:
            result["missing"].append("Attended MSME data not entered")
    except PersistentFailure as exc:
        result["check_failed"] = True
        result["missing"].append("Attended MSME check failed")
        errors.append(f"Event {event_id}: attended-count check failed ({exc})")

    try:
        docs = check_view_documents(client, unique_event_id)
        result["has_attendance_sheet"] = docs["has_attendance_photos"]
        result["has_event_photos"] = docs["has_event_photos"]
        if not docs["has_attendance_photos"]:
            result["missing"].append("Attendance sheet not uploaded")
        if not docs["has_event_photos"]:
            result["missing"].append("Event photos not uploaded")
    except PersistentFailure as exc:
        result["check_failed"] = True
        result["missing"].append("View Document check failed")
        errors.append(f"Event {event_id}: View Document check failed ({exc})")

    return result


def fmt_count(value):
    return "?" if value is None else str(value)


def render_no_events_email(yesterday_label: str):
    html_body = (
        f"<h2>RAMP Post-Event Admin Check — {yesterday_label}</h2>"
        f"<p>No events happened on {yesterday_label}. Nothing to check today.</p>"
    )
    plain_body = (
        f"RAMP Post-Event Admin Check — {yesterday_label}\n\n"
        f"No events happened on {yesterday_label}. Nothing to check today."
    )
    return html_body, plain_body


def render_email(missing_events: list[dict], complete_events: list[dict], yesterday_label: str):
    parts_html = [f"<h2>RAMP Post-Event Admin Check — {yesterday_label}</h2>"]
    parts_plain = [f"RAMP Post-Event Admin Check — {yesterday_label}", ""]

    if missing_events:
        parts_html.append("<h3>⚠ Needs attention</h3><ul>")
        parts_plain.append("NEEDS ATTENTION:")
        for e in missing_events:
            reasons = "; ".join(e["missing"])
            parts_html.append(
                f"<li><b>{e['event_name']}</b> ({e['event_id']}, {e['event_date']}): {reasons}</li>"
            )
            parts_plain.append(f"- {e['event_name']} ({e['event_id']}, {e['event_date']}): {reasons}")
        parts_html.append("</ul>")
        parts_plain.append("")

    if complete_events:
        parts_html.append("<h3>✅ Complete</h3><ul>")
        parts_plain.append("COMPLETE:")
        for e in complete_events:
            line = (
                f"Registered: {fmt_count(e['registered'])}, Attended: {fmt_count(e['attended'])}, "
                f"Attendance sheet: uploaded, Event photos: uploaded"
            )
            parts_html.append(
                f"<li><b>{e['event_name']}</b> ({e['event_id']}, {e['event_date']}) — {line}</li>"
            )
            parts_plain.append(f"- {e['event_name']} ({e['event_id']}, {e['event_date']}) — {line}")
        parts_html.append("</ul>")
        parts_plain.append("")

    return "\n".join(parts_html), "\n".join(parts_plain)


def _run(dry_run: bool = False, limit_events: int = None):
    load_dotenv()
    login_timeout = int(os.getenv("ADMIN_CHECK_LOGIN_TIMEOUT_SECONDS", "1800"))
    delay_seconds = float(os.getenv("REQUEST_DELAY_SECONDS", "2"))
    max_retries = int(os.getenv("MAX_RETRIES", "3"))

    sync = None
    if not dry_run:
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
        if not sheet_id or not service_account_json:
            print("ERROR: GOOGLE_SHEET_ID and GOOGLE_SERVICE_ACCOUNT_JSON must be set in .env")
            sys.exit(1)
        sync = SheetsSync(service_account_json, sheet_id)

    notify("Log in to RAMP now for the noon post-event admin check.", title="RAMP Post-Event Check")
    session = auth.interactive_login(timeout_seconds=login_timeout)
    print(f"Logged in as {session['user_type']} ({session['name']}).")

    client = RampClient(token=session["token"], delay_seconds=delay_seconds, max_retries=max_retries)
    start = time.monotonic()
    errors = []
    today = datetime.now(IST).date()
    yesterday = today - timedelta(days=1)
    yesterday_label = yesterday.strftime("%b %d, %Y")

    try:
        events = client.get_all_events()
        if limit_events:
            events = events[:limit_events]
        print(f"Found {len(events)} event(s) total.")

        candidates = [e for e in events if event_happened_on(e, yesterday)]
        candidates.sort(key=lambda e: str(e.get("eventId") or e.get("uniqueEventId") or ""))
        print(f"{len(candidates)} event(s) happened on {yesterday_label}.")

        missing_events, complete_events, log_entries = [], [], []
        for i, event in enumerate(candidates, start=1):
            result = check_event(client, event, errors)
            is_complete = not result["missing"]
            (complete_events if is_complete else missing_events).append(result)
            log_entries.append({
                **result,
                "status": "Complete" if is_complete else ("Check Failed" if result["check_failed"] else "Missing"),
                "details": "; ".join(result["missing"]) if result["missing"] else "All checks passed",
            })
            print(f"  [{i}/{len(candidates)}] Event {result['event_id']}: "
                  f"{'complete' if is_complete else ', '.join(result['missing'])}")
    finally:
        client.close()

    elapsed = time.monotonic() - start

    if not candidates:
        html_body, plain_body = render_no_events_email(yesterday_label)
    else:
        html_body, plain_body = render_email(missing_events, complete_events, yesterday_label)

    if dry_run:
        print(f"\n--- DRY RUN: nothing written to Sheets, no email sent ---\n")
        print(plain_body)
        print(f"\nDone in {elapsed/60:.1f} min. {len(errors)} error(s).")
        return

    if not candidates:
        log_entries = [{
            "event_id": "—", "event_name": "", "event_date": "",
            "registered": "", "attended": "", "has_attendance_sheet": False,
            "has_event_photos": False, "status": "No Events",
            "details": f"No events happened on {yesterday_label}.",
        }]
    sync.write_admin_check_log(log_entries)

    if not candidates:
        subject = f"RAMP Post-Event Check — no events on {yesterday_label}"
    elif missing_events and complete_events:
        subject = f"RAMP Post-Event Check — {len(missing_events)} need attention, {len(complete_events)} complete ({yesterday_label})"
    elif missing_events:
        subject = f"RAMP Post-Event Check — {len(missing_events)} event(s) need attention ({yesterday_label})"
    else:
        subject = f"RAMP Post-Event Check — all caught up ({yesterday_label})"

    recipients_env = "ADMIN_CHECK_RECIPIENTS" if os.getenv("ADMIN_CHECK_RECIPIENTS") else "REPORT_RECIPIENTS"
    emailer.send_report_email(subject, html_body, plain_body, recipients_env=recipients_env)
    print("Post-event admin check emailed.")
    notify(subject, title="RAMP Post-Event Check")

    print(f"Done in {elapsed/60:.1f} min. {len(errors)} error(s).")


def main():
    parser = argparse.ArgumentParser(description="RAMP noon post-event admin completion check")
    parser.add_argument("--dry-run", action="store_true", help="Run everything read-only: print the result, skip Sheets writes and the email send")
    parser.add_argument("--limit-events", type=int, default=None, help="Only scan the first N events (useful with --dry-run)")
    args = parser.parse_args()

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(RUN_LOCK_FILE, "w") as lock_file:
        if os.name == "nt":
            lock_file.write("1")
            lock_file.flush()
            lock_file.seek(0)
            try:
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError:
                print("Another admin-check run is already active; skipping this duplicate run.")
                return
        else:
            try:
                fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                print("Another admin-check run is already active; skipping this duplicate run.")
                return

        try:
            _run(dry_run=args.dry_run, limit_events=args.limit_events)
        except Exception as e:
            notify(f"RAMP noon admin check FAILED: {e}"[:200], title="RAMP Post-Event Check")
            raise


if __name__ == "__main__":
    main()
