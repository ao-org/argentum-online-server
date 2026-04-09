@echo off
REM ============================================================
REM  Register Argentum20 as a Windows Event Log source with
REM  a generic message file so Event Viewer can display
REM  descriptions for any Event ID.
REM
REM  Must be run as Administrator (one time per machine).
REM ============================================================

set SOURCE_NAME=Argentum20
set LOG_NAME=Application
set REG_KEY=HKLM\SYSTEM\CurrentControlSet\Services\EventLog\%LOG_NAME%\%SOURCE_NAME%

REM Use Windows' built-in generic message DLL (supports %1 string insertion)
set MSG_DLL=%SystemRoot%\System32\EventCreate.exe

echo Registering event source: %SOURCE_NAME%
echo Registry key: %REG_KEY%
echo Message file: %MSG_DLL%
echo.

REM Create the registry key and set EventMessageFile
reg add "%REG_KEY%" /v EventMessageFile /t REG_EXPAND_SZ /d "%MSG_DLL%" /f
if %errorlevel% neq 0 (
    echo [ERROR] Failed to set EventMessageFile. Are you running as Administrator?
    goto :end
)

REM TypesSupported = 7 means Information (1) + Warning (2) + Error (4)
reg add "%REG_KEY%" /v TypesSupported /t REG_DWORD /d 7 /f
if %errorlevel% neq 0 (
    echo [ERROR] Failed to set TypesSupported.
    goto :end
)

echo.
echo [OK] Event source "%SOURCE_NAME%" registered successfully.
echo      Event Viewer will now show descriptions for all Event IDs.
echo      Restart the Event Log service or reboot for changes to take effect.

:end
pause
