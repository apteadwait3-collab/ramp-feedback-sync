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
import html
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


def _event_date_only(event_date: str) -> str:
    parsed = parse_event_date(event_date)
    return parsed.strftime("%d %b %Y") if parsed else str(event_date)


def _yes_no_html(value):
    if value is None:
        return '<span style="color:#b45309;">Check failed</span>'
    if value:
        return '<span style="color:#15803d;">&#10003; Uploaded</span>'
    return '<span style="color:#b91c1c;">&#10007; Missing</span>'


def _yes_no_plain(value):
    return "Check failed" if value is None else ("Uploaded" if value else "MISSING")


def _count_html(value):
    if value is None:
        return '<span style="color:#b45309;">?</span>'
    if value == 0:
        return '<span style="color:#b91c1c;font-weight:bold;">0</span>'
    return f"<b>{value}</b>"


def render_no_events_email(yesterday_label: str):
    html_body = (
        '<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;color:#1f2937;">'
        '<h2 style="margin:0 0 4px;color:#1e3a8a;">RAMP Post-Event Digest</h2>'
        f'<div style="color:#6b7280;margin-bottom:16px;">Events held on {yesterday_label}</div>'
        '<div style="background:#f3f4f6;border-radius:6px;padding:14px 16px;">'
        f'No events were held on <b>{yesterday_label}</b>. Nothing to check today.</div>'
        '<p style="color:#9ca3af;font-size:12px;margin-top:20px;">Automated daily check &middot; MCCIA RAMP</p>'
        '</div>'
    )
    plain_body = (
        f"RAMP POST-EVENT DIGEST - events held on {yesterday_label}\n\n"
        f"No events were held on {yesterday_label}. Nothing to check today."
    )
    return html_body, plain_body


def render_email(missing_events: list[dict], complete_events: list[dict], yesterday_label: str):
    total = len(missing_events) + len(complete_events)
    all_events = sorted(missing_events + complete_events, key=lambda e: e["event_id"])
    n_bad, n_ok = len(missing_events), len(complete_events)
    banner_color, banner_bg = ("#b91c1c", "#fef2f2") if n_bad else ("#15803d", "#f0fdf4")
    banner_text = (
        f"{n_bad} of {total} event(s) need attention" if n_bad else f"All {total} event(s) fully complete"
    )

    th = 'style="text-align:left;padding:8px 10px;background:#1e3a8a;color:#fff;font-size:12px;white-space:nowrap;"'
    rows_html = []
    for e in all_events:
        ok = not e["missing"]
        status = (
            '<span style="color:#15803d;font-weight:bold;">&#10003; Complete</span>' if ok
            else '<span style="color:#b91c1c;font-weight:bold;">&#9888; Needs attention</span>'
        )
        td = 'style="padding:8px 10px;border-bottom:1px solid #e5e7eb;font-size:13px;vertical-align:top;"'
        rows_html.append(
            f"<tr><td {td}><b>{html.escape(e['event_name'])}</b><br>"
            f'<span style="color:#6b7280;font-size:11px;">{html.escape(e["event_id"])} &middot; {_event_date_only(e["event_date"])}</span></td>'
            f"<td {td} align='center'>{_count_html(e['registered'])}</td>"
            f"<td {td} align='center'>{_count_html(e['attended'])}</td>"
            f"<td {td}>{_yes_no_html(e['has_attendance_sheet'])}</td>"
            f"<td {td}>{_yes_no_html(e['has_event_photos'])}</td>"
            f"<td {td}>{status}</td></tr>"
        )

    action_html = ""
    plain_actions = []
    if missing_events:
        items = "".join(
            f"<li style='margin-bottom:6px;'><b>{html.escape(e['event_name'])}</b> "
            f"({html.escape(e['event_id'])}): {html.escape('; '.join(e['missing']))}</li>"
            for e in missing_events
        )
        action_html = (
            '<h3 style="margin:22px 0 8px;color:#b91c1c;font-size:15px;">Action required</h3>'
            '<div style="background:#fef2f2;border-left:4px solid #b91c1c;padding:10px 16px;font-size:13px;">'
            "Please update the RAMP portal for the following:"
            f'<ul style="margin:8px 0 0;padding-left:18px;">{items}</ul></div>'
        )
        plain_actions = ["", "ACTION REQUIRED - please update the RAMP portal:"] + [
            f"  * {e['event_name']} ({e['event_id']}): {'; '.join(e['missing'])}" for e in missing_events
        ]

    html_body = (
        '<div style="font-family:Arial,Helvetica,sans-serif;max-width:760px;color:#1f2937;">'
        '<h2 style="margin:0 0 4px;color:#1e3a8a;">RAMP Post-Event Digest</h2>'
        f'<div style="color:#6b7280;margin-bottom:16px;">Events held on {yesterday_label}</div>'
        f'<div style="background:{banner_bg};border-left:4px solid {banner_color};padding:12px 16px;margin-bottom:18px;">'
        f'<span style="color:{banner_color};font-weight:bold;font-size:15px;">{banner_text}</span><br>'
        f'<span style="font-size:13px;">{n_ok} complete &middot; {n_bad} incomplete &middot; {total} total</span></div>'
        '<table role="presentation" cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">'
        f"<tr><th {th}>Event</th><th {th}>Registered</th><th {th}>Attended</th>"
        f"<th {th}>Attendance sheet</th><th {th}>Event photos</th><th {th}>Status</th></tr>"
        + "".join(rows_html) + "</table>"
        + action_html
        + '<p style="color:#9ca3af;font-size:12px;margin-top:22px;">Automated daily check &middot; MCCIA RAMP. '
        "Counts and uploads are read live from the RAMP portal.</p></div>"
    )

    plain = [f"RAMP POST-EVENT DIGEST - events held on {yesterday_label}", "",
             f"{banner_text}  ({n_ok} complete, {n_bad} incomplete, {total} total)", ""]
    for e in all_events:
        plain += [
            f"* {e['event_name']} [{e['event_id']}, {_event_date_only(e['event_date'])}]",
            f"    Registered: {fmt_count(e['registered'])}   Attended: {fmt_count(e['attended'])}",
            f"    Attendance sheet: {_yes_no_plain(e['has_attendance_sheet'])}   "
            f"Event photos: {_yes_no_plain(e['has_event_photos'])}",
            f"    Status: {'Complete' if not e['missing'] else 'NEEDS ATTENTION'}",
            "",
        ]
    plain += plain_actions
    return html_body, "\n".join(plain)


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
            if not args.dry_run:
                send_failure_alert(e)
            raise


def send_failure_alert(error: Exception):
    """A run that dies (login never detected, portal down, ...) otherwise
    produces no email at all, which looks identical to a quiet day."""
    try:
        load_dotenv()
        today_label = datetime.now(IST).strftime("%b %d, %Y")
        reason = f"{type(error).__name__}: {error}"
        html_body = (
            f"<h2>RAMP noon check did NOT run — {today_label}</h2>"
            f"<p>No post-event report was produced today.</p><p><b>Reason:</b> {reason}</p>"
            "<p>If it says no login was detected, log in inside the Chromium window the "
            "script opens (not your normal browser) before the 30-minute timeout.</p>"
        )
        plain_body = (
            f"RAMP noon check did NOT run — {today_label}\n\n"
            f"No post-event report was produced today.\nReason: {reason}\n"
            "If it says no login was detected, log in inside the Chromium window the "
            "script opens (not your normal browser) before the 30-minute timeout."
        )
        env = "ADMIN_CHECK_ALERT_RECIPIENTS" if os.getenv("ADMIN_CHECK_ALERT_RECIPIENTS") else "ADMIN_CHECK_RECIPIENTS"
        emailer.send_report_email(
            f"RAMP noon check FAILED — no report today ({today_label})",
            html_body, plain_body, recipients_env=env,
        )
    except Exception as alert_error:
        print(f"Could not send failure alert: {alert_error}")


if __name__ == "__main__":
    main()
