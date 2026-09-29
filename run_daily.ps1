# Wrapper the Windows Scheduled Task calls every night at 23:00 IST.
# Windows equivalent of run_daily.sh: fixed working directory, venv
# python, and log redirection all in one obvious place.
Set-Location $PSScriptRoot
New-Item -ItemType Directory -Force -Path output | Out-Null
& "$PSScriptRoot\venv\Scripts\python.exe" daily_report.py *>> output\daily_report_output.log
