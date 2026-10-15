#!/usr/bin/env bash
# ==============================================================================
# Treasury Rates - Automated Comparison Scraper Cron Job Script (DRAFT)
# ==============================================================================
# Instructions:
# 1. Update PROJECT_ROOT to your application's absolute installation directory.
# 2. Update VENV_PATH to your virtual environment if applicable.
# 3. Add to crontab via `crontab -e`:
#    0 6 * * * /path/to/project/scripts/cron_rates_refresh.sh >> /path/to/project/logs/cron_rates.log 2>&1
# ==============================================================================

set -euo pipefail

# Configuration
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "${SCRIPT_DIR}")"
LOG_DIR="${PROJECT_ROOT}/logs"
LOCK_FILE="/tmp/treasury_rates_cron.lock"

mkdir -p "${LOG_DIR}"

echo "=============================================================================="
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Starting Treasury Rates Comparison Cron"
echo "Project Root: ${PROJECT_ROOT}"
echo "=============================================================================="

# Prevent concurrent overlapping runs using a file lock
exec 200>"${LOCK_FILE}"
if ! flock -n 200; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Another instance of rate scraper is already running. Skipping."
    exit 0
fi

# Activate virtual environment if present
if [ -f "${PROJECT_ROOT}/venv/bin/activate" ]; then
    source "${PROJECT_ROOT}/venv/bin/activate"
elif [ -f "${PROJECT_ROOT}/.venv/bin/activate" ]; then
    source "${PROJECT_ROOT}/.venv/bin/activate"
fi

cd "${PROJECT_ROOT}"

# Run comparison refresh via Django management command
python manage.py refresh_comparison_rates

EXIT_CODE=$?

if [ ${EXIT_CODE} -eq 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Rate comparison refresh completed successfully."
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Rate comparison refresh encountered errors (exit code ${EXIT_CODE})."
fi

exit ${EXIT_CODE}
