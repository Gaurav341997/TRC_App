@echo off
REM ==============================================================================
REM Windows Task Scheduler Batch Script (DRAFT ONLY)
REM Schedule via Task Scheduler or run manually:
REM schtasks /create /tn "TreasuryRatesRefresh" /tr "%~dp0cron_rates_refresh_windows.bat" /sc daily /st 06:00
REM ==============================================================================
setlocal EnableDelayedExpansion

cd /d "%~dp0.."
echo [%DATE% %TIME%] Starting Treasury Rates comparison refresh...
python manage.py refresh_comparison_rates >> logs\cron_rates.log 2>&1
echo [%DATE% %TIME%] Finished with exit code %ERRORLEVEL%
