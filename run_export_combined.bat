@echo off
set SCRIPT_DIR=C:\TestRailReportGenerator
set EXPORT_SCRIPT=exportTestRailCases.js
set COMBINE_SCRIPT=generateCombinedStatus_newold.js
set LOG_FILE=%SCRIPT_DIR%\raw_data\export_log.txt

REM Create raw_data folder if it doesn’t exist
if not exist "%SCRIPT_DIR%\raw_data" (
    mkdir "%SCRIPT_DIR%\raw_data"
)

cd /d %SCRIPT_DIR%

echo ============================= >> "%LOG_FILE%"
echo Running TestRail export at %date% %time% >> "%LOG_FILE%"
node "%EXPORT_SCRIPT%" >> "%LOG_FILE%" 2>&1

echo Running Combined Status Generator at %date% %time% >> "%LOG_FILE%"
node "%COMBINE_SCRIPT%" >> "%LOG_FILE%" 2>&1

echo Done. >> "%LOG_FILE%"
echo ============================= >> "%LOG_FILE%"
