@echo off
REM ===========================================================================
REM  bridge-service.cmd  --  restart-loop supervisor for the Hermes Excel bridge
REM ---------------------------------------------------------------------------
REM  requires elevation? NO. Per-user.
REM
REM  Launches the bridge (via run-bridge.cmd), waits for it to exit, and
REM  relaunches it. Includes a "health guard / stand-down" so it backs off
REM  (sleeps longer) when the bridge is intentionally down/unhealthy rather
REM  than hot-looping a crash. Restart events are appended with timestamps to
REM  the restart log UNDER THE DATA DIR (never the served web root).
REM
REM  Config block -- a deployer can retarget these. If INSTALL_DIR/DATA_DIR are
REM  not pre-set in the environment, they are derived/defaulted here so the
REM  script is runnable standalone. The Scheduled Task registered by
REM  register-task.ps1 launches this file.
REM ===========================================================================

setlocal EnableDelayedExpansion

REM --- Config (override via environment before launch) ---------------------
if not defined INSTALL_DIR set "INSTALL_DIR=%LOCALAPPDATA%\hermes\excel-addin"
if not defined HERMES_EXCEL_DATA_DIR set "HERMES_EXCEL_DATA_DIR=%LOCALAPPDATA%\hermes\excel-addin\data"
if not defined HEALTHY_BACKOFF_SECONDS set "HEALTHY_BACKOFF_SECONDS=3"
if not defined UNHEALTHY_BACKOFF_SECONDS set "UNHEALTHY_BACKOFF_SECONDS=30"
if not defined MAX_BACKOFF_SECONDS set "MAX_BACKOFF_SECONDS=120"

set "RUN_BRIDGE=%INSTALL_DIR%\run-bridge.cmd"
set "RESTART_LOG=%HERMES_EXCEL_DATA_DIR%\bridge-restarts.log"

if not exist "%RUN_BRIDGE%" (
  echo [bridge-service] ERROR: run-bridge.cmd not found at "%RUN_BRIDGE%"
  exit /b 1
)
if not exist "%HERMES_EXCEL_DATA_DIR%" mkdir "%HERMES_EXCEL_DATA_DIR%" 2>nul

REM --- Singleton guard -----------------------------------------------------
REM Task Scheduler can leave an old cmd supervisor orphaned when its wscript
REM parent is stopped. Two supervisors then alternate killing each other's
REM node process. Own an atomic lock directory and record this cmd.exe PID;
REM a new launch exits when the owner is alive, or reclaims a stale lock.
set "SUPERVISOR_LOCK=%HERMES_EXCEL_DATA_DIR%\bridge-supervisor.lock"
set "SUPERVISOR_PID="
for /f "usebackq tokens=*" %%P in (`powershell -NoProfile -Command "$ps=Get-CimInstance Win32_Process -Filter ('ProcessId=' + $PID); $forCmd=Get-CimInstance Win32_Process -Filter ('ProcessId=' + $ps.ParentProcessId); $forCmd.ParentProcessId"`) do set "SUPERVISOR_PID=%%P"
if not defined SUPERVISOR_PID (
  echo [bridge-service] ERROR: could not determine supervisor PID
  exit /b 1
)

2>nul mkdir "%SUPERVISOR_LOCK%"
if not errorlevel 1 goto lock_acquired

set "EXISTING_PID="
if exist "%SUPERVISOR_LOCK%\owner.pid" set /p EXISTING_PID=<"%SUPERVISOR_LOCK%\owner.pid"
if defined EXISTING_PID (
  powershell -NoProfile -Command "$p=Get-CimInstance Win32_Process -Filter ('ProcessId=' + $env:EXISTING_PID) -ErrorAction SilentlyContinue; if ($p -and $p.Name -eq 'cmd.exe' -and $p.CommandLine -like '*bridge-service.cmd*') { exit 0 }; exit 1"
  if not errorlevel 1 (
    echo [bridge-service] Supervisor PID !EXISTING_PID! already owns the lock; exiting duplicate.
    exit /b 0
  )
)

rmdir /s /q "%SUPERVISOR_LOCK%" 2>nul
2>nul mkdir "%SUPERVISOR_LOCK%"
if errorlevel 1 (
  echo [bridge-service] ERROR: could not reclaim stale supervisor lock
  exit /b 1
)

:lock_acquired
>"%SUPERVISOR_LOCK%\owner.pid" echo !SUPERVISOR_PID!

call :log "supervisor starting (install=%INSTALL_DIR%)"

set "BACKOFF=%HEALTHY_BACKOFF_SECONDS%"

:loop
  REM Capture start time so we can tell a quick crash from a long, healthy run.
  set "START_TICK=%TIME%"
  call :log "launching bridge"

  REM run-bridge.cmd runs node in the foreground; this CALL blocks until exit.
  call "%RUN_BRIDGE%"
  set "RC=!ERRORLEVEL!"

  call :log "bridge exited rc=!RC!"

  REM --- Health guard / stand-down ----------------------------------------
  REM  If the bridge stayed up a while it was "healthy enough"; restart fast.
  REM  If it died almost immediately (likely intentionally down / misconfig /
  REM  Node missing), back off with an increasing delay so we don't hot-loop.
  call :ranged_seconds "%START_TICK%" elapsed

  if !elapsed! GEQ 20 (
    set "BACKOFF=%HEALTHY_BACKOFF_SECONDS%"
    call :log "ran !elapsed!s before exit; fast restart in !BACKOFF!s"
  ) else (
    set /a "BACKOFF=BACKOFF*2"
    if !BACKOFF! LSS %UNHEALTHY_BACKOFF_SECONDS% set "BACKOFF=%UNHEALTHY_BACKOFF_SECONDS%"
    if !BACKOFF! GTR %MAX_BACKOFF_SECONDS% set "BACKOFF=%MAX_BACKOFF_SECONDS%"
    call :log "quick exit (!elapsed!s); standing down, backoff !BACKOFF!s"
  )

  REM timeout is interruptible; /nobreak avoids accidental keypress skip noise.
  timeout /t !BACKOFF! /nobreak >nul 2>nul
  goto loop

REM --- helpers --------------------------------------------------------------
:log
  REM %~1 = message. Appends "YYYY-MM-DD HH:MM:SS  message" to the restart log.
  set "MSG=%~1"
  for /f "usebackq tokens=*" %%T in (`powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"`) do set "STAMP=%%T"
  >>"%RESTART_LOG%" echo !STAMP!  !MSG!
  echo [bridge-service] !STAMP!  !MSG!
  exit /b 0

:ranged_seconds
  REM Crude elapsed-seconds since %~1 (HH:MM:SS[.cc]) to now. Good enough to
  REM distinguish "crashed instantly" from "ran for a while". Returns in %2.
  for /f "usebackq tokens=*" %%E in (`powershell -NoProfile -Command "$s=[datetime]::Parse('%~1'); $n=Get-Date; $d=[int]($n - $s).TotalSeconds; if ($d -lt 0) { $d += 86400 }; $d"`) do set "%~2=%%E"
  exit /b 0
