# Wrapper the Windows Scheduled Task calls every day at 12:00 PM IST for
# the noon post-event admin completion check (registered/attended counts +
# attendance-sheet/event-photo uploads). Independent of run_daily.ps1
# (feedback sync + Sheet write), which is now Monday-only.
Set-Location $PSScriptRoot
New-Item -ItemType Directory -Force -Path output | Out-Null
& "$PSScriptRoot\venv\Scripts\python.exe" admin_completion_check.py *>> output\admin_check_output.log
