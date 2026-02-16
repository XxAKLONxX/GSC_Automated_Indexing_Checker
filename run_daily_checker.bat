@echo off
REM ============================================================================
REM GSC Indexing Checker - Daily Runner
REM This batch file runs the Python script and logs output
REM ============================================================================

REM Change to the script directory
cd /d "%~dp0"

REM Set timestamp for log
set timestamp=%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%%time:~6,2%
set timestamp=%timestamp: =0%

REM Log file
set logfile=run_log_%timestamp%.txt

echo ============================================ > %logfile%
echo GSC Indexing Checker - Daily Run >> %logfile%
echo Started at: %date% %time% >> %logfile%
echo ============================================ >> %logfile%
echo. >> %logfile%

REM Run the Python script
python gsc_auto_indexing_checker.py >> %logfile% 2>&1

REM Check if script succeeded
if %ERRORLEVEL% EQU 0 (
    echo. >> %logfile%
    echo ============================================ >> %logfile%
    echo Script completed successfully >> %logfile%
    echo ============================================ >> %logfile%
) else (
    echo. >> %logfile%
    echo ============================================ >> %logfile%
    echo ERROR: Script failed with error code %ERRORLEVEL% >> %logfile%
    echo ============================================ >> %logfile%
)

REM Keep the window open if running manually (optional - comment out for scheduled runs)
REM pause
