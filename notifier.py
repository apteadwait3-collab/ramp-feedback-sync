"""
Best-effort desktop notification (macOS + Windows) shared by every
unattended scheduled entrypoint (daily_report.py, admin_completion_check.py)
so a login window opening on an otherwise-unwatched machine doesn't go
unnoticed. Never allowed to break the run it's reporting on.
"""

import html
import subprocess
import sys


def notify(message: str, title: str = "RAMP Automation"):
    try:
        if sys.platform == "darwin":
            subprocess.run(
                [
                    "osascript", "-e",
                    f'display notification "{message}" with title "{title}" sound name "Glass"',
                ],
                check=False, timeout=5,
            )
        elif sys.platform == "win32":
            escaped_title = html.escape(title, quote=True)
            escaped_message = html.escape(message, quote=True)
            ps_script = (
                "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null; "
                "[Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null; "
                "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null; "
                "$xml = New-Object Windows.Data.Xml.Dom.XmlDocument; "
                "$xml.LoadXml('<toast><visual><binding template=\"ToastGeneric\">"
                "<text>" + escaped_title + "</text><text>" + escaped_message + "</text>"
                "</binding></visual></toast>'); "
                "$toast = New-Object Windows.UI.Notifications.ToastNotification $xml; "
                "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('RAMP Automation').Show($toast)"
            )
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                check=False, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
    except Exception:
        pass
