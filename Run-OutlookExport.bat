@echo off
setlocal

set "SCRIPT=%~dp0Export-OutlookSender.ps1"
set "OUTPUT=%~dp0outlook-export"

if not exist "%SCRIPT%" (
    set "SCRIPT=%CD%\Export-OutlookSender.ps1"
    set "OUTPUT=%CD%\outlook-export"
)

if not exist "%SCRIPT%" (
    echo Error: Export-OutlookSender.ps1 was not found.
    echo Expected path: "%SCRIPT%"
    echo Run this BAT from the Mix folder, or keep both files in the same folder.
    pause
    exit /b 1
)

set "SENDER="
set /p "SENDER=Sender email address: "
if not defined SENDER (
    echo Error: Sender email address is required.
    pause
    exit /b 1
)

set "START_DATE="
set "END_DATE="
echo.
echo Enter dates as YYYY-MM-DD. Leave both blank for incremental export.
set /p "START_DATE=Start date (blank = no lower limit): "
set /p "END_DATE=End date (blank = now; includes the whole day): "

echo.
echo Exporting messages ...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& $env:SCRIPT -Sender $env:SENDER -OutputPath $env:OUTPUT -StartDate $env:START_DATE -EndDate $env:END_DATE -Recurse"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" echo Export failed with exit code %EXIT_CODE%.
pause
exit /b %EXIT_CODE%
