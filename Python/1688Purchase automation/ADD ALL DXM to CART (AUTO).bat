@echo off
setlocal EnableExtensions

REM ============================================================
REM Pipeline:
REM   1) Run DXM export+audit (dxm_export_and_audit.py)
REM   2) If DXM produced a NEW picklist xlsx in this run, run add_to_cart_http_1688.py
REM ============================================================

cd /d D:\JoeProgramFiles\Automation

REM IMPORTANT: set this to your actual picklist folder
REM If your picklists are saved elsewhere, change this path.
set "PICKLIST_FOLDER=D:\JoeProgramFiles\Automation\Batch_added_to_cart"

REM Capture pipeline start time (epoch seconds)
for /f %%i in ('powershell -NoProfile -Command "[int](Get-Date -UFormat %%s)"') do set "PIPELINE_START_EPOCH=%%i"

REM Run DXM export + audit
python dxm_export_and_audit.py

REM ------------------------------------------------------------
REM After DXM: check whether a NEW picklist xlsx was created/modified AFTER pipeline start
REM exit code meanings:
REM   0 = found new xlsx
REM   2 = no xlsx exists
REM   3 = latest xlsx is older than pipeline start (=> DXM exported nothing new)
REM ------------------------------------------------------------
powershell -NoProfile -Command ^
  "$d=$env:PICKLIST_FOLDER; $t=[int]$env:PIPELINE_START_EPOCH; " ^
  "if(-not (Test-Path $d)){ exit 2 } " ^
  "$f=Get-ChildItem $d -Filter *.xlsx -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1; " ^
  "if($null -eq $f){ exit 2 } " ^
  "$epoch=[int][double]::Parse((Get-Date $f.LastWriteTime -UFormat %%s)); " ^
  "if($epoch -lt $t){ exit 3 } else { exit 0 }"

if errorlevel 3 (
  echo [INFO] DXM exported no new pick list in this run. Skipping add_to_cart.
  goto :eof
)
if errorlevel 2 (
  echo [INFO] No pick list xlsx exists in PICKLIST_FOLDER. Skipping add_to_cart.
  goto :eof
)

REM Allow add_to_cart to auto-confirm if latest xlsx is fresh
set "AUTO_CONFIRM_LATEST=1"
REM Optional: set a tighter window, e.g. 10 minutes
REM set "AUTO_CONFIRM_WINDOW_SEC=600"

python add_to_cart_http_1688.py wholesale

endlocal
