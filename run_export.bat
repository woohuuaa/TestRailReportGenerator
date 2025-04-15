@echo off
set SCRIPT_DIR=C:\TestRailReportGenerator
set SCRIPT_FILE=exportTestRailCases.js
set LOG_FILE=%SCRIPT_DIR%\raw_data\export_log.txt

REM Create raw_data folder if it doesn’t exist
if not exist "%SCRIPT_DIR%\raw_data" (
    mkdir "%SCRIPT_DIR%\raw_data"
)

REM Run the script with Node and log output
cd /d %SCRIPT_DIR%
echo Running TestRail export at %date% %time% >> "%LOG_FILE%"
node "%SCRIPT_FILE%" >> "%LOG_FILE%" 2>&1
echo Done. >> "%LOG_FILE%"