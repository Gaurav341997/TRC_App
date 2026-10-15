"""Django Management Command to refresh Retail, Bulk, and NRE/NRO rate comparison workbooks.

Designed for automated execution via Cron Job, Celery beat, or Windows Task Scheduler.
DO NOT EXECUTE AUTOMATICALLY - DRAFTED AS REQUESTED.

Usage:
    python manage.py refresh_comparison_rates
    python manage.py refresh_comparison_rates --only retail,bulk
    python manage.py refresh_comparison_rates --skip-scrape
"""

import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand

logger = logging.getLogger("rates.cron")


class Command(BaseCommand):
    help = "Refresh Retail, Bulk, and NRE/NRO TD comparison workbooks (Draft for Cron Job)"

    def add_arguments(self, parser):
        parser.add_argument(
            "--only",
            type=str,
            default="",
            help="Comma-separated modules to refresh: retail,bulk,nre_nro (default: all)",
        )
        parser.add_argument(
            "--skip-scrape",
            action="store_true",
            help="Skip live network scraping and rebuild comparison workbooks from existing CSVs",
        )
        parser.add_argument(
            "--timeout",
            type=int,
            default=900,
            help="Timeout per scraper module in seconds (default: 900)",
        )

    def handle(self, *args, **options):
        project_dir = Path(settings.BASE_DIR)
        comparison_dir = project_dir / "rates" / "comparison_engine"

        available_modules = {
            "retail": {
                "name": "Retail TD Comparison",
                "script": comparison_dir / "retail_main.py",
            },
            "bulk": {
                "name": "Bulk TD Comparison",
                "script": comparison_dir / "bulk_main.py",
            },
            "nre_nro": {
                "name": "NRE/NRO TD Comparison",
                "script": comparison_dir / "nre_nro_main.py",
            },
        }

        only_filter = [
            m.strip().lower()
            for m in options.get("only", "").split(",")
            if m.strip()
        ]
        skip_scrape = options.get("skip_scrape", False)
        timeout = options.get("timeout", 900)

        selected_modules = (
            {k: v for k, v in available_modules.items() if k in only_filter}
            if only_filter
            else available_modules
        )

        self.stdout.write(
            self.style.NOTICE(
                f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Starting comparison rates refresh cron..."
            )
        )

        overall_success = True

        for key, module in selected_modules.items():
            name = module["name"]
            script_path = module["script"]

            if not script_path.exists():
                self.stderr.write(self.style.ERROR(f"Script not found for {name}: {script_path}"))
                overall_success = False
                continue

            cmd = [sys.executable, str(script_path)]
            if skip_scrape:
                cmd.append("--skip-scrape")

            self.stdout.write(f"Executing {name} ({script_path.name})...")
            start_time = datetime.now()

            try:
                proc = subprocess.run(
                    cmd,
                    cwd=str(project_dir),
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                    check=False,
                )
                elapsed = datetime.now() - start_time

                if proc.returncode == 0:
                    self.stdout.write(
                        self.style.SUCCESS(f" [SUCCESS] {name} completed in {elapsed}")
                    )
                else:
                    overall_success = False
                    self.stderr.write(
                        self.style.ERROR(
                            f" [FAILED] {name} exited with code {proc.returncode} in {elapsed}\n"
                            f"STDERR: {proc.stderr[:1000]}\n"
                            f"STDOUT: {proc.stdout[:500]}"
                        )
                    )

            except subprocess.TimeoutExpired:
                overall_success = False
                self.stderr.write(
                    self.style.ERROR(f" [TIMEOUT] {name} timed out after {timeout} seconds")
                )
            except Exception as exc:
                overall_success = False
                self.stderr.write(
                    self.style.ERROR(f" [ERROR] {name} failed: {exc}")
                )

        if overall_success:
            self.stdout.write(
                self.style.SUCCESS(
                    f"[{datetime.now():%Y-%m-%d %H:%M:%S}] All modules refreshed successfully."
                )
            )
        else:
            self.stderr.write(
                self.style.WARNING(
                    f"[{datetime.now():%Y-%m-%d %H:%M:%S}] One or more modules failed during refresh."
                )
            )
            sys.exit(1)
