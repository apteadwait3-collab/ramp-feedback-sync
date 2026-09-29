"""
Daily entrypoint: one manual login, then everything else is automatic —
scrape+sync any new feedback into Sheets, build a "last 7 days" analysis,
run the two post-event completion checks (enrolled MSME data entry, View
Document photos) against every past event, and email one digest to
REPORT_RECIPIENTS. Meant to be fired by the com.mccia.ramp-daily-report
LaunchAgent every night at 23:00 IST; run it by hand any time with
`python daily_report.py` (or `--dry-run` to preview without touching
Sheets or sending mail — see main() below).

Auth model is unchanged from main.py (see auth.py / RECON.md): a human
completes the RAMP login + CAPTCHA in a real browser window this script
opens. A desktop notification fires when that window opens since this is
now started by a scheduler rather than a person sitting at the terminal.

The Sheets sync half reuses main.py's exact checkpoint-driven scrape loop
(scrape_event / load_checkpoint / save_checkpoint) — a daily run behaves
like a normal full run there. The "last 7 days" summary is independent: it
always fetches fresh feedback + enrolled/attended counts for events whose
eventDate falls in the last REPORT_WINDOW_DAYS, regardless of checkpoint
state, since the email needs current numbers even if an event was already
synced in an earlier run. The post-event checks are independent again —
see post_event_checks.py for the "already passed" / resolved-checkpoint
logic that keeps them from re-scanning all of history every night.
"""

import argparse
import html
import json
import os
import subprocess
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
import post_event_checks
from main import load_checkpoint, save_checkpoint, scrape_event
from notifier import notify as _notify
from post_event_checks import parse_event_date
from sheets_sync import SheetsSync
from source_client import PersistentFailure, RampClient

REPORT_WINDOW_DAYS = 7

OUTPUT_DIR = "output"
LAST_SENT_FILE = os.path.join(OUTPUT_DIR, "last_digest_sent.json")
RUN_LOCK_FILE = os.path.join(OUTPUT_DIR, "daily_report.lock")


def notify(message: str):
    _notify(message, title="RAMP Daily Report")


def already_sent_today(today) -> bool:
    """Idempotency guard for the digest send: running the job twice in one
    IST calendar day (e.g. a manual re-run after a scheduled one) must not
    re-email or double-archive the report. Sheets/feedback sync stays safe
    to re-run on its own (Enrollment-ID dedup, the scrape checkpoint, and
    post_event_checks' resolved-checkpoint) — this is the one genuinely
    non-idempotent action (an SMTP send) that needs its own guard."""
    if not os.path.exists(LAST_SENT_FILE):
        return False
    try:
        with open(LAST_SENT_FILE) as f:
            state = json.load(f)
        return state.get("date") == today.isoformat()
    except (OSError, json.JSONDecodeError, TypeError):
        return False


def mark_sent_today(today):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    tmp_path = f"{LAST_SENT_FILE}.tmp"
    with open(tmp_path, "w") as f:
        json.dump({"date": today.isoformat()}, f)
    os.replace(tmp_path, LAST_SENT_FILE)


def in_report_window(event_date: str, today) -> bool:
    d = parse_event_date(event_date)
    if not d:
        return False
    return today - timedelta(days=REPORT_WINDOW_DAYS) <= d <= today


def build_week_summary(client: RampClient, week_events: list, errors: list):
    """Full feedback + enrolled/attended fetch for this week's events only
    (expected to be a handful), independent of the Sheets checkpoint.

    Returns (summary_events, feedback_rows): summary_events carries the
    per-event enrolled/attended/feedback counts shown in the email;
    feedback_rows is the flat list of individual respondent rows that
    actually contain feedback, used only to compute the total feedback
    count. A respondent who's merely on the feedback roster with nothing
    filled in doesn't count as "gave feedback".
    """
    summary = []
    feedback_rows = []
    for event in week_events:
        event_id = event.get("eventId") or event.get("uniqueEventId") or "?"
        guid = event.get("uniqueEventId")
        rows = scrape_event(client, event, errors)
        given_rows = [r for r in rows if r["Feedback"].strip()]
        feedback_rows.extend(given_rows)

        enrolled = attended = None
        if guid:
            try:
                enrolled = client.get_enrolled_count(guid)
            except PersistentFailure as e:
                errors.append(f"{event_id}: enrolled count failed ({e})")
            try:
                attended = client.get_attended_count(guid)
            except PersistentFailure as e:
                errors.append(f"{event_id}: attended count failed ({e})")

        summary.append({
            "eventId": event_id,
            "component": event.get("componentName", ""),
            "eventName": event.get("eventName", ""),
            "eventDate": event.get("eventDate", ""),
            "enrolled": enrolled,
            "attended": attended,
            "feedbackCount": len(given_rows),
        })
    return summary, feedback_rows


# --- Email rendering -------------------------------------------------------
#
# Fully inline styles, no <style> block, no CSS variables/media queries —
# this goes through real inboxes (Gmail web/app, Outlook desktop's Word
# rendering engine), not a browser, so only the lowest-common-denominator
# subset of CSS is safe.

NAVY = "#14304f"
INK = "#0b0b0b"
INK_2 = "#52514e"
MUTED = "#898781"
BORDER = "#e5e5e0"
ACCENT = "#eb6834"
WARN = "#9a6a00"


def esc(v) -> str:
    return html.escape(str(v)) if v is not None else ""


def fmt_count(n) -> str:
    return f"{n:,}" if n is not None else "—"


def stat_card_html(value, label: str) -> str:
    """One card in the top stat row — width is a percentage so 5 cards
    divide the row evenly regardless of email client."""
    return f"""<td width="20%" style="padding:6px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {BORDER};border-radius:8px;background:#ffffff;">
        <tr><td align="center" style="padding:16px 6px;">
          <div style="font-size:19px;font-weight:800;color:{INK};font-family:Arial,Helvetica,sans-serif;line-height:1.2;">{esc(value)}</div>
          <div style="font-size:10px;color:{MUTED};margin-top:5px;font-family:Arial,Helvetica,sans-serif;text-transform:uppercase;letter-spacing:0.03em;">{esc(label)}</div>
        </td></tr>
      </table>
    </td>"""


def event_header_html(ev: dict) -> str:
    """Section header for one event — name + identity, with its own
    enrolled/attended/feedback line, sitting above that event's feedback
    cards so each event is visually its own group."""
    date_only = esc(ev["eventDate"]).split(" ")[0]
    return f"""<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:{NAVY};border-radius:8px;margin:0 0 8px;">
      <tr><td style="padding:12px 16px;">
        <div style="font-size:14px;font-weight:700;color:#ffffff;font-family:Arial,Helvetica,sans-serif;">{esc(ev['eventName'])}</div>
        <div style="font-size:11px;color:#ffffff;opacity:0.82;margin-top:3px;font-family:Arial,Helvetica,sans-serif;">{esc(ev['eventId'])} &middot; {date_only} &middot; {fmt_count(ev['enrolled'])} registered &middot; {fmt_count(ev['attended'])} attended &middot; {ev['feedbackCount']} feedback</div>
      </td></tr>
    </table>"""


def event_section_html(ev: dict) -> str:
    return f'<div style="margin-bottom:20px;">{event_header_html(ev)}</div>'


GOOD = "#1e7d32"
BAD = "#b3261e"

# How many missing-MSME names to list inline before collapsing the rest —
# same defensive-cap spirit as sheets_sync._capped / source_client's
# FEEDBACK_CHAR_LIMIT, just for a list instead of a string.
MAX_MSMES_SHOWN = 15


def pending_event_html(item: dict) -> str:
    """One pending event's card in the "Pending Post-Event Items" section.
    Event photos and attendance-sheet photos are always both stated
    explicitly (Present/Missing), not just shown when missing — the task
    this section serves wants both flags visible per event regardless."""
    date_only = esc(item["eventDate"]).split(" ")[0]
    msmes = item["missing_msmes"]

    if msmes is None:
        msme_html = (
            f'<div style="margin-top:8px;font-size:12px;color:{WARN};font-family:Arial,Helvetica,sans-serif;">'
            f'MSME data-entry check: <strong>Failed; will retry.</strong></div>'
        )
    elif msmes:
        shown = msmes[:MAX_MSMES_SHOWN]
        li_items = "".join(
            f'<li>{esc(m["name"] or "(no name)")} ({esc(m["udyam"] or "no Udyam #")})'
            f' &mdash; {esc(m.get("reason", "Incomplete"))}</li>'
            for m in shown
        )
        if len(msmes) > MAX_MSMES_SHOWN:
            li_items += f'<li>…and {len(msmes) - MAX_MSMES_SHOWN} more</li>'
        msme_html = (
            f'<div style="margin-top:8px;font-size:12px;color:{INK_2};font-family:Arial,Helvetica,sans-serif;">'
            f'<strong>MSMEs missing post-event data ({len(msmes)}):</strong>'
            f'<ul style="margin:4px 0 0 18px;padding:0;">{li_items}</ul></div>'
        )
    else:
        msme_html = (
            f'<div style="margin-top:8px;font-size:12px;color:{GOOD};font-family:Arial,Helvetica,sans-serif;">'
            f'MSME data entry: complete.</div>'
        )

    def flag(missing: bool | None, label: str) -> str:
        if missing is None:
            return f'{label}: <strong style="color:{WARN};">Check failed; will retry</strong>'
        color = BAD if missing else GOOD
        state = "Missing" if missing else "Present"
        return f'{label}: <strong style="color:{color};">{state}</strong>'

    photos_line = (
        flag(item["missing_event_photos"], "Event photos")
        + " &middot; "
        + flag(item["missing_attendance_photos"], "Attendance sheet photos")
    )

    return f"""<div style="border:1px solid {BORDER};border-radius:8px;padding:12px 16px;margin-bottom:10px;background:#fffaf6;">
      <div style="font-size:13px;font-weight:700;color:{INK};font-family:Arial,Helvetica,sans-serif;">{esc(item['eventName'])}</div>
      <div style="font-size:11px;color:{MUTED};margin-top:2px;font-family:Arial,Helvetica,sans-serif;">{esc(item['eventId'])} &middot; {date_only}</div>
      <div style="font-size:12px;color:{INK_2};margin-top:8px;font-family:Arial,Helvetica,sans-serif;">{photos_line}</div>
      {msme_html}
    </div>"""


def pending_items_html(pending_items: list) -> str:
    """Always renders a visible section — an explicit "all clear" line
    when pending_items is empty, never an omitted section, so it's clear
    from the email itself that the check actually ran."""
    if not pending_items:
        body = (
            f'<p style="padding:16px;text-align:center;color:{GOOD};'
            f'font-family:Arial,Helvetica,sans-serif;font-size:13px;margin:0;border:1px solid {BORDER};'
            f'border-radius:8px;background:#f5fbf6;">All clear &mdash; no pending post-event items.</p>'
        )
    else:
        body = "".join(pending_event_html(item) for item in pending_items)
    return f"""<div style="margin-top:24px;">
      <div style="font-size:14px;font-weight:700;color:{INK};font-family:Arial,Helvetica,sans-serif;margin-bottom:12px;">Pending Post-Event Items</div>
      {body}
    </div>"""


def pending_items_plain(pending_items: list) -> list:
    lines = ["", "Pending Post-Event Items:"]
    if not pending_items:
        lines.append("All clear — no pending post-event items.")
        return lines
    for item in pending_items:
        date_only = str(item["eventDate"]).split(" ")[0]
        lines.append("")
        lines.append(f"== {item['eventName']} ({item['eventId']}) — {date_only} ==")
        def state(value):
            if value is None:
                return "Check failed; will retry"
            return "Missing" if value else "Present"

        lines.append(f"  Event photos: {state(item['missing_event_photos'])}")
        lines.append(f"  Attendance sheet photos: {state(item['missing_attendance_photos'])}")
        msmes = item["missing_msmes"]
        if msmes is None:
            lines.append("  MSME data-entry check: Failed; will retry")
        elif msmes:
            lines.append(f"  MSMEs missing post-event data ({len(msmes)}):")
            for m in msmes[:MAX_MSMES_SHOWN]:
                lines.append(
                    f"    - {m['name'] or '(no name)'} ({m['udyam'] or 'no Udyam #'})"
                    f" — {m.get('reason', 'Incomplete')}"
                )
            if len(msmes) > MAX_MSMES_SHOWN:
                lines.append(f"    …and {len(msmes) - MAX_MSMES_SHOWN} more")
        else:
            lines.append("  MSME data entry: complete.")
    return lines


def render_email(week_label: str, month_label: str, summary_events: list, feedback_rows: list,
                  pending_items: list, errors: list, total_synced: int):
    total_events = len(summary_events)
    total_enrolled = sum(e["enrolled"] for e in summary_events if e["enrolled"] is not None)
    total_attended = sum(e["attended"] for e in summary_events if e["attended"] is not None)

    total_feedback = len(feedback_rows)

    cards_html = "".join([
        stat_card_html(month_label, "Month"),
        stat_card_html(total_events, "Events Happened"),
        stat_card_html(fmt_count(total_enrolled), "Registered"),
        stat_card_html(fmt_count(total_attended), "Attended"),
        stat_card_html(fmt_count(total_feedback), "Gave Feedback"),
    ])

    events_sorted = sorted(summary_events, key=lambda e: e["eventDate"])
    events_html = "".join(
        event_section_html(ev) for ev in events_sorted
    ) or (
        f'<p style="padding:20px;text-align:center;color:{MUTED};'
        f'font-family:Arial,Helvetica,sans-serif;font-size:13px;margin:0;border:1px solid {BORDER};'
        f'border-radius:8px;background:#fcfcfb;">No events fell in this window.</p>'
    )

    error_html = ""
    if errors:
        error_html = (
            f'<p style="margin:18px 0 0;font-size:12px;color:{WARN};'
            f'font-family:Arial,Helvetica,sans-serif;">{len(errors)} issue(s) occurred while collecting '
            f'counts/feedback this run &mdash; see run.log for details.</p>'
        )

    html_body = f"""<html><body style="margin:0;padding:0;background:#f2f2ee;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2ee;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="700" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid {BORDER};">
  <tr><td style="background:{NAVY};padding:20px 24px;">
    <div style="color:#ffffff;font-size:18px;font-weight:800;font-family:Arial,Helvetica,sans-serif;">MCCIA RAMP Feedback &mdash; Daily Digest</div>
    <div style="color:#ffffff;opacity:0.8;font-size:12px;margin-top:4px;font-family:Arial,Helvetica,sans-serif;">Last 7 days: {esc(week_label)} &middot; Industry Association scope</div>
  </td></tr>
  <tr><td style="padding:20px 24px 4px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
      <tr>{cards_html}</tr>
    </table>
  </td></tr>
  <tr><td style="padding:14px 24px 24px;">
    <div style="font-size:14px;font-weight:700;color:{INK};font-family:Arial,Helvetica,sans-serif;margin-bottom:12px;">Events &amp; Feedback (Last 7 Days)</div>
    {events_html}
    {error_html}
    <p style="margin:20px 0 0;font-size:12px;color:{MUTED};font-family:Arial,Helvetica,sans-serif;">
      {total_synced} new feedback row(s) were synced to the
      <a href="https://docs.google.com/spreadsheets/d/{esc(os.getenv('GOOGLE_SHEET_ID', ''))}" style="color:{ACCENT};">Feedback MSSIDC</a>
      sheet this run.
    </p>
    {pending_items_html(pending_items)}
  </td></tr>
</table>
</td></tr>
</table>
</body></html>"""

    plain_lines = [
        f"MCCIA RAMP Feedback — Daily Digest — Last 7 days: {week_label} ({month_label})",
        "",
        f"Events Happened: {total_events}",
        f"Registered: {fmt_count(total_enrolled)}",
        f"Attended: {fmt_count(total_attended)}",
        f"Gave Feedback: {fmt_count(total_feedback)}",
        "",
        "Events & Feedback (Last 7 Days):",
    ]
    for ev in events_sorted:
        date_only = str(ev["eventDate"]).split(" ")[0]
        plain_lines.append("")
        plain_lines.append(
            f"== {ev['eventName']} ({ev['eventId']}) — {date_only} — "
            f"{fmt_count(ev['enrolled'])} registered, {fmt_count(ev['attended'])} attended, "
            f"{ev['feedbackCount']} feedback =="
        )
    if not events_sorted:
        plain_lines.append("No events fell in this window.")
    plain_lines.append("")
    plain_lines.append(f"{total_synced} new feedback row(s) synced to the Feedback MSSIDC sheet this run.")
    if errors:
        plain_lines.append(f"{len(errors)} issue(s) occurred — see run.log for details.")
    plain_lines.extend(pending_items_plain(pending_items))

    return html_body, "\n".join(plain_lines)


# --- Orchestration -----------------------------------------------------


def _run(dry_run: bool = False, limit_events: int = None):
    load_dotenv()
    login_timeout = int(os.getenv("LOGIN_TIMEOUT_SECONDS", "1800"))
    delay_seconds = float(os.getenv("REQUEST_DELAY_SECONDS", "2"))
    max_retries = int(os.getenv("MAX_RETRIES", "3"))

    # Even --dry-run needs read access because the MSME data-entry check is
    # explicitly based on the existing Sheet's schema and tracked value.
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sheet_id or not service_account_json:
        print("ERROR: GOOGLE_SHEET_ID and GOOGLE_SERVICE_ACCOUNT_JSON must be set in .env")
        sys.exit(1)
    sync = SheetsSync(service_account_json, sheet_id)

    notify("Log in to RAMP now to run tonight's feedback sync + post-event checks.")
    session = auth.interactive_login(timeout_seconds=login_timeout)
    print(f"Logged in as {session['user_type']} ({session['name']}).")

    client = RampClient(token=session["token"], delay_seconds=delay_seconds, max_retries=max_retries)
    start = time.monotonic()
    errors = []
    today = datetime.now(post_event_checks.IST).date()
    total_added = 0
    pending_items = []

    try:
        events = client.get_all_events()
        if limit_events:
            events = events[:limit_events]
        print(f"Found {len(events)} event(s) total.")

        week_events = [e for e in events if in_report_window(e.get("eventDate", ""), today)]
        print(f"{len(week_events)} event(s) fall in the report window (last {REPORT_WINDOW_DAYS} days).")
        summary_events, feedback_rows = build_week_summary(client, week_events, errors)

        if not dry_run:
            done_ids = load_checkpoint()

            def checkpoint_key(e):
                return e.get("eventId") or e.get("uniqueEventId")

            pending = [e for e in events if checkpoint_key(e) not in done_ids]
            print(f"{len(pending)} event(s) not yet checkpointed; syncing to Sheets...")
            for i, event in enumerate(pending, start=1):
                rows = scrape_event(client, event, errors)
                if rows:
                    total_added += sync.sync(rows)
                key = checkpoint_key(event)
                if key:
                    done_ids.add(key)
                    save_checkpoint(done_ids)
                print(f"  [{i}/{len(pending)}] synced")

        # --dry-run ignores the resolved-checkpoint, same convention as
        # main.py's --dry-run ignoring the scrape checkpoint — every past
        # event gets (re)checked so a preview run shows the full picture.
        resolved_ids = post_event_checks.load_resolved() if not dry_run else set()
        sheet_rows_by_event = sync.read_post_event_rows(events)
        print("Running post-event completion checks (enrolled MSME data entry + View Document photos)...")
        pending_items, resolved_ids = post_event_checks.run_post_event_checks(
            client,
            events,
            sheet_rows_by_event,
            resolved_ids,
            today,
            errors,
        )
        print(f"{len(pending_items)} past event(s) with pending post-event items.")
        if not dry_run:
            post_event_checks.save_resolved(resolved_ids)
    finally:
        client.close()

    today_label = today.strftime("%b %d, %Y")
    week_start_label = (today - timedelta(days=REPORT_WINDOW_DAYS)).strftime("%b %d")
    week_label = f"{week_start_label} – {today_label}"
    month_label = today.strftime("%B %Y")

    html_body, plain_body = render_email(
        week_label, month_label, summary_events, feedback_rows, pending_items, errors, total_added
    )

    elapsed = time.monotonic() - start

    if dry_run:
        print("\n--- DRY RUN: nothing written to Sheets, no email sent ---\n")
        print(plain_body)
        print(f"\nDone in {elapsed/60:.1f} min. {len(errors)} error(s).")
        return

    sync.update_overview()
    sync.write_run_log(events_scanned=len(events), new_rows=total_added, errors=errors)

    if already_sent_today(today):
        print(f"Digest already sent today ({today.isoformat()}) — skipping duplicate send.")
    else:
        sync.write_daily_report(today_label, month_label, summary_events, pending_items, errors, html_body, plain_body)
        emailer.send_report_email(
            subject=f"RAMP Feedback — Daily Digest ({today_label})",
            html_body=html_body,
            plain_fallback=plain_body,
        )
        mark_sent_today(today)
        print("Daily digest emailed.")
        notify("Daily RAMP digest sent.")

    print(f"Done in {elapsed/60:.1f} min. {len(errors)} error(s), {total_added} new row(s) synced.")


def main():
    parser = argparse.ArgumentParser(description="RAMP daily digest: feedback sync + post-event completion checks")
    parser.add_argument("--dry-run", action="store_true", help="Run everything read-only: print the digest, skip Sheets writes and the email send")
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
                print("Another daily RAMP run is already active; skipping this duplicate run.")
                return
        else:
            try:
                fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                print("Another daily RAMP run is already active; skipping this duplicate run.")
                return

        try:
            _run(dry_run=args.dry_run, limit_events=args.limit_events)
        except Exception as e:
            notify(f"RAMP daily run FAILED: {e}"[:200])
            raise


if __name__ == "__main__":
    main()
