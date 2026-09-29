# Wrapper the Windows Scheduled Task calls every night at 22:50 IST — ten
# minutes before run_daily.ps1 opens the RAMP login window.
Set-Location $PSScriptRoot
New-Item -ItemType Directory -Force -Path output | Out-Null
& "$PSScriptRoot\venv\Scripts\python.exe" login_reminder.py *>> output\login_reminder_output.log
