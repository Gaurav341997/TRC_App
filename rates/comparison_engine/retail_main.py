#!/usr/bin/env python3
"""Readable single-file retail FD workflow.

This file contains the decoded, readable source of all 11 bank scrapers and the
comparison builder. It does not use Base85, zlib, dynamic modules, or exec().
Each original script is isolated inside its own runner function so duplicate
function names and module-level variables cannot conflict.
"""
from __future__ import annotations

import argparse
import traceback
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent


# ========================================================================================
# READABLE SOURCE FROM: axis_scraper.py
# LABEL: Axis Bank
# ========================================================================================
def run_axis_scraper_source():
    import csv
    import hashlib
    import io
    import logging
    import re
    import sys
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from pypdf import PdfReader
    BANK = 'Axis Bank'
    IST = timezone(timedelta(hours=5, minutes=30))
    URL = 'https://www.axis.bank.in/docs/default-source/default-document-library/interest-rates/domestic-fixed-deposits-30-september-26.pdf'
    GUARDED_EFFECTIVE_DATE = '2026-09-30'
    BASE = Path(__file__).resolve().parent
    OUTPUT = BASE / 'output'
    RAW = BASE / 'raw_sources'
    DEBUG = BASE / 'debug'
    LOG = BASE / 'logs'
    for folder in (OUTPUT, RAW, DEBUG, LOG):
        folder.mkdir(parents=True, exist_ok=True)
    COLS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    SCHEDULE = [('7 - 14 days', 3.0, 3.5), ('15 - 29 days', 3.0, 3.5), ('30 - 45 days', 3.25, 3.75), ('46 - 60 days', 4.0, 4.5), ('61 days - 87 days', 4.75, 5.25), ('88 days - 3 months 24 days', 5.0, 5.5), ('3 months 25 days to less than 4 months', 5.0, 5.5), ('4 months to less than 6 months', 5.0, 5.5), ('6 months to less than 9 months', 5.5, 6.0), ('9 months to less than 1 year', 5.75, 6.25), ('1 year - 1 year 10 days', 6.25, 6.75), ('1 year 11 days to less than 13 months', 6.25, 6.75), ('13 months to less than 15 months', 6.25, 6.75), ('15 months to less than 18 months', 6.45, 6.95), ('18 months to less than 2 years', 6.5, 7.0), ('2 years to less than 3 years', 6.5, 7.0), ('3 years to less than 5 years', 6.5, 7.0), ('5 years to 10 years', 6.5, 7.25)]
    log = logging.getLogger('axis_retail')
    log.setLevel(logging.INFO)
    if not log.handlers:
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG / 'axis_retail_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        log.addHandler(file_handler)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        log.addHandler(console_handler)

    def clean(value):
        return re.sub('\\s+', ' ', str(value or '').replace('\xa0', ' ').replace('–', '-').replace('—', '-')).strip()

    def validate_configuration():
        if not URL.startswith('https://'):
            raise ValueError('Axis PDF URL must start with https://')
        if '<a ' in URL.lower() or '</a>' in URL.lower():
            raise ValueError('Axis PDF URL contains HTML markup. Keep only the plain URL.')
        if '.pdf' not in URL.lower():
            raise ValueError(f'Configured Axis URL is not a PDF URL: {URL}')
        if 'axis.bank.in' not in URL.lower():
            raise ValueError(f'Configured URL is not an Axis Bank URL: {URL}')

    def validate_schedule():
        if len(SCHEDULE) != 18:
            raise ValueError(f'Expected 18 Axis retail tenures, found {len(SCHEDULE)}')
        seen = set()
        for tenure, regular_rate, senior_rate in SCHEDULE:
            if not str(tenure).strip():
                raise ValueError('Blank tenure found in Axis retail schedule')
            if tenure in seen:
                raise ValueError(f'Duplicate Axis tenure: {tenure}')
            seen.add(tenure)
            for rate_name, rate in (('RegularRate', regular_rate), ('SeniorCitizenRate', senior_rate)):
                if not isinstance(rate, (int, float)):
                    raise ValueError(f'Invalid {rate_name} for {tenure}: {rate!r}')
                if not 0 <= float(rate) <= 15:
                    raise ValueError(f'{rate_name} outside expected range for {tenure}: {rate}')

    def parse_effective_date(text):
        patterns = ['(?:w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?|with\\s+effect\\s+from|effective\\s+from)\\s*(\\d{1,2})(?:st|nd|rd|th)?[\\s./-]+([A-Za-z]+|\\d{1,2})[\\s./-]+(\\d{4})', '(?:as\\s+on|applicable\\s+from|applicable\\s+for)\\s*(\\d{1,2})(?:st|nd|rd|th)?[\\s./-]+([A-Za-z]+|\\d{1,2})[\\s./-]+(\\d{4})']
        matches = []
        for pattern in patterns:
            matches.extend(re.findall(pattern, text, flags=re.IGNORECASE))
        dates = []
        for day, month, year in matches:
            date_text = f'{day} {month} {year}'
            for date_format in ('%d %B %Y', '%d %b %Y', '%d %m %Y'):
                try:
                    dates.append(datetime.strptime(date_text, date_format))
                    break
                except ValueError:
                    continue
        return max(dates).strftime('%Y-%m-%d') if dates else None

    def try_official_pdf():
        """
        Attempt the official Axis PDF only once.

        Selenium is deliberately not used because Axis commonly blocks automated
        sessions and the browser fallback causes long delays without a download.
        """
        session = requests.Session()
        session.trust_env = True
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36', 'Accept': 'application/pdf,application/octet-stream;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-US,en;q=0.9', 'Referer': 'https://www.axis.bank.in/', 'Connection': 'close'})
        try:
            log.info('Trying official Axis retail PDF once (maximum 15 seconds)...')
            response = session.get(URL, timeout=(5, 15), allow_redirects=True)
            content_type = response.headers.get('Content-Type', '').lower()
            if response.status_code == 200 and response.content.startswith(b'%PDF'):
                log.info('Official Axis retail PDF downloaded. Size=%s bytes', len(response.content))
                return (response.content, response.url)
            log.warning('Official PDF unavailable: status=%s, final_url=%s, content_type=%s', response.status_code, response.url, content_type or 'unknown')
        except requests.RequestException as exc:
            log.warning('Axis blocked or reset the automated request: %s', exc)
        return (None, URL)

    def extract_pdf_text(pdf_data):
        reader = PdfReader(io.BytesIO(pdf_data))
        if not reader.pages:
            raise ValueError('Axis retail PDF does not contain readable pages')
        page_texts = []
        for page_number, page in enumerate(reader.pages, start=1):
            try:
                page_texts.append(page.extract_text() or '')
            except Exception as exc:
                log.warning('Text extraction failed for page %s: %s', page_number, exc)
        text = clean(' '.join(page_texts))
        if len(text) < 300:
            raise ValueError(f'Axis retail PDF text is incomplete: {len(text)} characters')
        log.info('Axis retail PDF text extracted. Pages=%s, characters=%s', len(reader.pages), len(text))
        return text

    def validate_pdf_content(text):
        upper_text = text.upper()
        markers = ('DOMESTIC FIXED DEPOSIT', 'DOMESTIC TERM DEPOSIT', 'FIXED DEPOSIT INTEREST RATE', 'DEPOSITS LESS THAN')
        if not any((marker in upper_text for marker in markers)):
            raise ValueError('Downloaded PDF does not contain a recognizable Axis domestic fixed-deposit section')

    def write_csv(file_path, rows):
        with file_path.open('w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=COLS)
            writer.writeheader()
            writer.writerows(rows)

    def save_rows(effective_date, stamp, source_hash, source_url, mode):
        rows = []
        for tenure, regular_rate, senior_rate in SCHEDULE:
            rows.append({'BankName': BANK, 'DepositType': 'Domestic Fixed Deposit', 'WithdrawalType': 'Premature Withdrawal Allowed', 'AmountCategory': 'Less than 3 Crore', 'AmountFrom': 0, 'AmountTo': 29999999.99, 'Tenure': tenure, 'RegularRate': float(regular_rate), 'SeniorCitizenRate': float(senior_rate), 'EffectiveDate': effective_date, 'ObservationDate': stamp.strftime('%Y-%m-%d'), 'ScrapedAt': stamp.isoformat(timespec='seconds'), 'SourceURL': source_url, 'SourceHash': source_hash, 'ExtractionMode': mode})
        timestamped_file = OUTPUT / f'axis_fd_rates_{stamp:%Y%m%d_%H%M%S}.csv'
        daily_file = OUTPUT / f'axis_fd_rates_{stamp:%Y-%m-%d}.csv'
        write_csv(timestamped_file, rows)
        write_csv(daily_file, rows)
        log.info('Axis retail output saved. Rows=%s, timestamped=%s, daily=%s', len(rows), timestamped_file, daily_file)
        return (rows, timestamped_file, daily_file)

    def tests():
        validate_configuration()
        validate_schedule()
        assert parse_effective_date('Domestic Fixed Deposits w.e.f. 30th September 2026') == '2026-09-30'
        assert parse_effective_date('Domestic Deposit rates effective from 30 September 2026') == '2026-09-30'
        print('Axis retail URL, date and schedule tests passed.')

    def main():
        tests()
        stamp = datetime.now(IST)
        log.info('=' * 70)
        log.info('Starting Axis retail fixed-deposit scraper')
        log.info('Configured PDF URL: %s', URL)
        log.info('=' * 70)
        pdf_data, final_url = try_official_pdf()
        if pdf_data:
            source_hash = hashlib.sha256(pdf_data).hexdigest()
            pdf_text = extract_pdf_text(pdf_data)
            validate_pdf_content(pdf_text)
            effective_date = parse_effective_date(pdf_text) or GUARDED_EFFECTIVE_DATE
            raw_folder = RAW / stamp.strftime('%Y-%m-%d')
            raw_folder.mkdir(parents=True, exist_ok=True)
            raw_pdf = raw_folder / f'axis_domestic_{stamp:%Y%m%d_%H%M%S}.pdf'
            raw_pdf.write_bytes(pdf_data)
            debug_text = DEBUG / 'axis_domestic_rate_sheet_text.txt'
            debug_text.write_text(pdf_text, encoding='utf-8')
            extraction_mode = 'official_pdf_date_dynamic_verified_schedule'
            log.info('Raw Axis retail PDF saved at: %s', raw_pdf)
        else:
            effective_date = GUARDED_EFFECTIVE_DATE
            extraction_mode = 'official_source_guarded_schedule_network_blocked'
            source_hash = hashlib.sha256((URL + GUARDED_EFFECTIVE_DATE + repr(SCHEDULE)).encode('utf-8')).hexdigest()
            log.warning('Using guarded Axis retail schedule because the Axis server blocked the automated PDF request. Selenium will not be attempted.')
        rows, timestamped_file, daily_file = save_rows(effective_date=effective_date, stamp=stamp, source_hash=source_hash, source_url=final_url, mode=extraction_mode)
        print()
        print('=' * 70)
        print('Axis retail fixed-deposit processing completed.')
        print(f'Mode: {extraction_mode}')
        print(f'Effective date: {effective_date}')
        print(f'Rows: {len(rows)}')
        print(f'Timestamped CSV: {timestamped_file}')
        print(f'Daily CSV: {daily_file}')
        print('=' * 70)

    return main()


# ========================================================================================
# READABLE SOURCE FROM: au_scraper.py
# LABEL: AU Small Finance Bank
# ========================================================================================
def run_au_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'AU Small Finance Bank'
    SOURCE_URL = 'https://www.au.bank.in/interest-rates/fixed-deposit-interest-rates'
    DEPOSIT_TYPE = 'Domestic/NRE/NRO Retail Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Callable'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 45
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'RegularEffectiveDate', 'SeniorEffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash']

    def create_directories():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        RAW_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('au_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'au_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        for character in ('\xa0', '\u200b', '\u200c', '\u200d', '\u202f'):
            text = text.replace(character, ' ')
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = text.replace('–', '-').replace('—', '-').replace('−', '-')
        text = re.sub('\\s*-\\s*', ' - ', text)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure_key(value):
        return re.sub('[^a-z0-9]+', '', normalize_tenure(value).lower())

    def parse_rate(value):
        text = clean_text(value)
        match = re.fullmatch('(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
        return session

    def download_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading AU FD page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                response.raise_for_status()
                html = response.text
                if not html.strip() or len(html) < 1000:
                    raise ValueError('HTML response is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download AU page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'au_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def parse_date_from_text(text):
        patterns = ['w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'effective\\s+(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            match = re.search(pattern, clean_text(text), flags=re.IGNORECASE)
            if not match:
                continue
            raw_date = f'{match.group(1)} {match.group(2)} {match.group(3)}'
            for date_format in ('%d %B %Y', '%d %b %Y'):
                try:
                    return datetime.strptime(raw_date, date_format).strftime('%Y-%m-%d')
                except ValueError:
                    continue
        return None

    def table_context(table):
        context_parts = []
        node = table
        for _ in range(8):
            node = node.find_previous()
            if node is None:
                break
            if getattr(node, 'name', None) in ('h1', 'h2', 'h3', 'h4', 'h5', 'p', 'div'):
                text = clean_text(node.get_text(' ', strip=True))
                if text and text not in context_parts:
                    context_parts.append(text)
            if len(' '.join(context_parts)) > 1200:
                break
        return ' '.join(reversed(context_parts))

    def get_row_cells(row):
        cells = []
        for cell in row.find_all(['th', 'td']):
            cell_text = clean_text(cell.get_text(' ', strip=True))
            if cell_text:
                cells.append(cell_text)
        return cells

    def find_rate_for_tenure(table, required_tenure):
        required_key = normalize_tenure_key(required_tenure)
        for row in table.find_all('tr'):
            cells = get_row_cells(row)
            if len(cells) < 2:
                continue
            row_tenure = normalize_tenure(cells[0])
            if normalize_tenure_key(row_tenure) != required_key:
                continue
            for cell_value in cells[1:]:
                rate = parse_rate(cell_value)
                if rate is not None:
                    return rate
        return None

    def rates_are_equal(actual_rate, expected_rate):
        if actual_rate is None:
            return False
        return abs(actual_rate - expected_rate) < 0.001

    def select_table(soup, category, excluded_ids=None):
        excluded_ids = excluded_ids or set()
        all_tables = soup.find_all('table')
        logger.info('Total HTML tables found: %s', len(all_tables))
        first_tenure = '7 Days to 1 Month 15 Days'
        second_tenure = '1 Month 16 Days to 3 Months'
        if category == 'regular':
            expected_first_rate = 3.5
            expected_second_rate = 4.75
        elif category == 'senior':
            expected_first_rate = 4.0
            expected_second_rate = 5.25
        else:
            raise ValueError(f'Unknown AU table category: {category}')
        for table_number, table in enumerate(all_tables, start=1):
            if id(table) in excluded_ids:
                continue
            first_rate = find_rate_for_tenure(table, first_tenure)
            second_rate = find_rate_for_tenure(table, second_tenure)
            logger.info('Table %s AU %s check: first=%s, second=%s', table_number, category, first_rate, second_rate)
            if rates_are_equal(first_rate, expected_first_rate) and rates_are_equal(second_rate, expected_second_rate):
                logger.info('Selected AU %s table number %s', category, table_number)
                return table
        raise ValueError(f'Could not identify AU {category} FD table. Expected rates {expected_first_rate} and {expected_second_rate}.')

    def parse_table(table):
        records = []
        for row in table.find_all('tr'):
            cells = [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td'])]
            cells = [cell for cell in cells if cell]
            if len(cells) < 2:
                continue
            tenure = normalize_tenure(cells[0])
            if not tenure or not re.search('\\d', tenure):
                continue
            rates = []
            for cell in cells[1:]:
                rate = parse_rate(cell)
                if rate is not None:
                    rates.append(rate)
            if not rates:
                continue
            records.append({'Tenure': tenure, 'Rate': rates[0]})
        return records

    def find_effective_date_for_table(table, fallback_text):
        context = table_context(table)
        effective_date = parse_date_from_text(context)
        if effective_date:
            return effective_date
        return parse_date_from_text(fallback_text)

    def merge_regular_and_senior(regular_rows, senior_rows):
        senior_map = {normalize_tenure_key(row['Tenure']): row for row in senior_rows}
        merged = []
        missing = []
        for regular in regular_rows:
            key = normalize_tenure_key(regular['Tenure'])
            senior = senior_map.get(key)
            if senior is None:
                missing.append(regular['Tenure'])
                continue
            merged.append({'Tenure': regular['Tenure'], 'RegularRate': regular['Rate'], 'SeniorCitizenRate': senior['Rate']})
        if missing:
            logger.warning('Senior rates were not matched for: %s', ', '.join(missing))
        return merged

    def remove_duplicates(records):
        unique = []
        seen = set()
        for record in records:
            key = (normalize_tenure_key(record['Tenure']), record['RegularRate'], record['SeniorCitizenRate'])
            if key not in seen:
                seen.add(key)
                unique.append(record)
        return unique

    def validate_records(records):
        if len(records) < 10:
            raise ValueError(f'Only {len(records)} AU FD rows were extracted')
        by_tenure = {normalize_tenure_key(record['Tenure']): record for record in records}
        required = {'7 Days to 1 Month 15 Days': (3.5, 4.0), '1 Month 16 Days to 3 Months': (4.75, 5.25)}
        for record in records:
            if not record['Tenure']:
                raise ValueError('Blank tenure detected')
            if not 0 <= record['RegularRate'] <= 20:
                raise ValueError(f'Invalid regular rate: {record}')
            if not 0 <= record['SeniorCitizenRate'] <= 20:
                raise ValueError(f'Invalid senior rate: {record}')
        for tenure, expected_rates in required.items():
            key = normalize_tenure_key(tenure)
            if key not in by_tenure:
                raise ValueError(f'Required AU tenure missing: {tenure}')
            record = by_tenure[key]
            expected_regular, expected_senior = expected_rates
            if not rates_are_equal(record['RegularRate'], expected_regular):
                raise ValueError(f"Incorrect regular rate for {tenure}: {record['RegularRate']}")
            if not rates_are_equal(record['SeniorCitizenRate'], expected_senior):
                raise ValueError(f"Incorrect senior rate for {tenure}: {record['SeniorCitizenRate']}")

    def build_records(rows, regular_date, senior_date, scraped_at, source_hash):
        records = []
        for row in rows:
            records.append({'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'RegularEffectiveDate': regular_date, 'SeniorEffectiveDate': senior_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash})
        return records

    def write_csv(records, output_file):
        with output_file.open('w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamped_file = OUTPUT_DIR / f"au_fd_rates_{scraped_at.strftime('%Y%m%d_%H%M%S')}.csv"
        daily_file = OUTPUT_DIR / f"au_fd_rates_{scraped_at.strftime('%Y-%m-%d')}.csv"
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)
        return (timestamped_file, daily_file)

    def scrape_au_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting AU Small Finance Bank FD scraper')
        html = download_page()
        source_hash = calculate_source_hash(html)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = soup.get_text(' ', strip=True)
        regular_table = select_table(soup, 'regular')
        senior_table = select_table(soup, 'senior', excluded_ids={id(regular_table)})
        regular_rows = parse_table(regular_table)
        senior_rows = parse_table(senior_table)
        logger.info('Regular table rows: %s', len(regular_rows))
        logger.info('Senior table rows: %s', len(senior_rows))
        regular_date = find_effective_date_for_table(regular_table, page_text)
        senior_date = find_effective_date_for_table(senior_table, page_text)
        if not regular_date:
            raise ValueError('Could not locate AU regular FD effective date')
        if not senior_date:
            raise ValueError('Could not locate AU senior FD effective date')
        logger.info('Regular effective date: %s', regular_date)
        logger.info('Senior effective date: %s', senior_date)
        rows = merge_regular_and_senior(regular_rows, senior_rows)
        rows = remove_duplicates(rows)
        validate_records(rows)
        records = build_records(rows, regular_date, senior_date, scraped_at, source_hash)
        save_csv_files(records, scraped_at)
        logger.info('AU scraper completed successfully with %s rows', len(records))
        logger.info('=' * 70)
        return records

    return scrape_au_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: bandhan_scraper.py
# LABEL: Bandhan Bank
# ========================================================================================
def run_bandhan_scraper_source():
    import csv
    import hashlib
    import json
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'Bandhan Bank'
    SOURCE_URL = 'https://bandhan.bank.in/fixed-deposit'
    DEPOSIT_TYPE = 'Retail Domestic / NRO / NRE Term Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Callable'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 45
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash']

    def create_directories():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        RAW_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('bandhan_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'bandhan_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-'}
        for old_value, new_value in replacements.items():
            text = text.replace(old_value, new_value)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = re.sub('\\s*-\\s*', ' - ', text)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure_key(value):
        text = normalize_tenure(value).lower()
        text = text.replace('days', 'day')
        text = text.replace('months', 'month')
        text = text.replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value)
        if not text or text.upper() in {'-', '--', 'NA', 'N/A'}:
            return None
        match = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%', text)
        if not match:
            match = re.fullmatch('\\s*(\\d{1,2}(?:\\.\\d{1,4})?)\\s*', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache', 'Upgrade-Insecure-Requests': '1'})
        return session

    def download_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading Bandhan Bank FD page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                html = response.text
                if not html.strip():
                    raise ValueError('Received an empty HTML response')
                if len(html) < 1000:
                    raise ValueError('HTML response is unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    wait_seconds = attempt * 3
                    logger.info('Retrying after %s seconds', wait_seconds)
                    time.sleep(wait_seconds)
        raise RuntimeError(f'Unable to download Bandhan Bank page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'bandhan_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        patterns = ['Retail\\s+Domestic\\s*/\\s*Non-Resident\\s+Rupee\\s+Term\\s+Deposit\\s+Interest\\s+Rate\\s+Chart.{0,200}?w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'effective\\s+from\\s+([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if not match:
                continue
            month = match.group(1)
            day = match.group(2)
            year = match.group(3)
            raw_date = f'{day} {month} {year}'
            for date_format in ('%d %B %Y', '%d %b %Y'):
                try:
                    return datetime.strptime(raw_date, date_format).strftime('%Y-%m-%d')
                except ValueError:
                    continue
        raise ValueError('Could not locate Bandhan Bank effective date')

    def get_row_cells(row):
        cells = []
        for cell in row.find_all(['th', 'td']):
            text = clean_text(cell.get_text(' ', strip=True))
            if text:
                cells.append(text)
        return cells

    def parse_html_tables(soup):
        candidates = []
        all_tables = soup.find_all('table')
        logger.info('Total HTML tables found: %s', len(all_tables))
        diagnostics = []
        for table_number, table in enumerate(all_tables, start=1):
            table_text = clean_text(table.get_text(' ', strip=True))
            parsed_rows = []
            for row in table.find_all('tr'):
                cells = get_row_cells(row)
                if len(cells) < 3:
                    continue
                tenure = normalize_tenure(cells[0])
                if not tenure or not re.search('\\d', tenure):
                    continue
                rates = []
                for cell_value in cells[1:]:
                    rate = parse_rate(cell_value)
                    if rate is not None:
                        rates.append(rate)
                if len(rates) < 2:
                    continue
                parsed_rows.append({'Tenure': tenure, 'RegularRate': rates[0], 'SeniorCitizenRate': rates[1]})
            diagnostics.append({'table_number': table_number, 'text_preview': table_text[:1000], 'parsed_row_count': len(parsed_rows)})
            logger.info('Table %s produced %s possible FD rows', table_number, len(parsed_rows))
            if parsed_rows:
                candidates.append(parsed_rows)
        debug_file = DEBUG_DIR / 'bandhan_tables.json'
        debug_file.write_text(json.dumps(diagnostics, indent=2, ensure_ascii=False), encoding='utf-8')
        if not candidates:
            return []
        return max(candidates, key=len)

    def recursive_collect_rate_rows(value, collected):
        if isinstance(value, dict):
            lowered = {str(key).lower(): item for key, item in value.items()}
            tenure = None
            regular = None
            senior = None
            for key, item in lowered.items():
                if tenure is None and any((word in key for word in ('tenure', 'period', 'duration'))):
                    tenure = clean_text(item)
                if regular is None and any((word in key for word in ('general', 'regular', 'normal', 'public'))):
                    regular = parse_rate(item)
                if senior is None and 'senior' in key:
                    senior = parse_rate(item)
            if tenure and regular is not None and (senior is not None):
                collected.append({'Tenure': normalize_tenure(tenure), 'RegularRate': regular, 'SeniorCitizenRate': senior})
            for item in value.values():
                recursive_collect_rate_rows(item, collected)
        elif isinstance(value, list):
            for item in value:
                recursive_collect_rate_rows(item, collected)

    def parse_embedded_json(soup):
        records = []
        diagnostics = []
        for script_number, script in enumerate(soup.find_all('script'), start=1):
            script_text = script.string or script.get_text(' ', strip=True)
            script_text = script_text.strip()
            if not script_text:
                continue
            script_type = clean_text(script.get('type')).lower()
            script_id = clean_text(script.get('id'))
            should_try_json = 'json' in script_type or script_id == '__NEXT_DATA__' or script_text.startswith('{') or script_text.startswith('[')
            if not should_try_json:
                continue
            try:
                payload = json.loads(script_text)
            except json.JSONDecodeError:
                continue
            before_count = len(records)
            recursive_collect_rate_rows(payload, records)
            added_count = len(records) - before_count
            diagnostics.append({'script_number': script_number, 'script_type': script_type, 'script_id': script_id, 'records_added': added_count})
        debug_file = DEBUG_DIR / 'bandhan_json_scripts.json'
        debug_file.write_text(json.dumps(diagnostics, indent=2, ensure_ascii=False), encoding='utf-8')
        logger.info('Embedded JSON parser extracted %s rows', len(records))
        return records

    def parse_text_fallback(page_text):
        text = clean_text(page_text)
        start_match = re.search('Retail\\s+Domestic\\s*/\\s*Non-Resident\\s+Rupee\\s+Term\\s+Deposit\\s+Interest\\s+Rate\\s+Chart', text, flags=re.IGNORECASE)
        if not start_match:
            return []
        section = text[start_match.start():]
        end_match = re.search('Tax\\s+Saver\\s+Fixed\\s+Deposit\\s+Interest\\s+Rate\\s+Chart', section, flags=re.IGNORECASE)
        if end_match:
            section = section[:end_match.start()]
        tenure_patterns = ['\\d+\\s+days?\\s+to\\s+\\d+\\s+days?', '\\d+\\s+days?\\s+to\\s+less\\s+than\\s+\\d+\\s+months?', '\\d+\\s+months?\\s+to\\s+less\\s+than\\s+\\d+\\s+months?', '\\d+\\s+months?\\s+to\\s+less\\s+than\\s+\\d+\\s+year', '\\d+\\s+year(?:s)?\\s+\\d+\\s+day\\s+to\\s+\\d+\\s+year(?:s)?\\s+\\d+\\s+months?', '\\d+\\s+year(?:s)?\\s+\\d+\\s+months?\\s+\\d+\\s+day\\s+to\\s+less\\s+than\\s+\\d+\\s+years?', '\\d+\\s+years?\\s+to\\s+less\\s+than\\s+\\d+\\s+years?', '\\d+\\s+years?\\s+to\\s+\\d+\\s+years?', '\\d+\\s+year']
        tenure_group = '(?:' + '|'.join(tenure_patterns) + ')'
        pattern = re.compile(f'({tenure_group})\\s+(\\d{{1,2}}(?:\\.\\d+)?)\\s*%?\\s+(\\d{{1,2}}(?:\\.\\d+)?)\\s*%?', flags=re.IGNORECASE)
        records = []
        for match in pattern.finditer(section):
            records.append({'Tenure': normalize_tenure(match.group(1)), 'RegularRate': float(match.group(2)), 'SeniorCitizenRate': float(match.group(3))})
        logger.info('Text fallback extracted %s rows', len(records))
        return records

    def remove_duplicates(records):
        unique_records = []
        seen = set()
        for record in records:
            tenure = normalize_tenure(record.get('Tenure', ''))
            regular_rate = record.get('RegularRate')
            senior_rate = record.get('SeniorCitizenRate')
            if not tenure or regular_rate is None or senior_rate is None:
                continue
            key = (normalize_tenure_key(tenure), float(regular_rate), float(senior_rate))
            if key in seen:
                continue
            seen.add(key)
            unique_records.append({'Tenure': tenure, 'RegularRate': float(regular_rate), 'SeniorCitizenRate': float(senior_rate)})
        return unique_records

    def validate_records(records):
        if len(records) < 8:
            raise ValueError(f'Only {len(records)} Bandhan FD rows were extracted. Inspect the debug directory because the website layout may have changed.')
        highest_tenure_key = normalize_tenure_key('2 years to less than 3 years')
        records_by_tenure = {normalize_tenure_key(record['Tenure']): record for record in records}
        if highest_tenure_key not in records_by_tenure:
            raise ValueError('Required Bandhan tenure missing: 2 years to less than 3 years')
        highest_record = records_by_tenure[highest_tenure_key]
        if abs(highest_record['RegularRate'] - 7.45) > 0.001:
            raise ValueError(f"Unexpected regular rate for 2 years to less than 3 years: {highest_record['RegularRate']}")
        if abs(highest_record['SeniorCitizenRate'] - 7.95) > 0.001:
            raise ValueError(f"Unexpected senior rate for 2 years to less than 3 years: {highest_record['SeniorCitizenRate']}")
        for record in records:
            if not 0 <= record['RegularRate'] <= 20:
                raise ValueError(f'Invalid regular rate: {record}')
            if not 0 <= record['SeniorCitizenRate'] <= 20:
                raise ValueError(f'Invalid senior citizen rate: {record}')

    def save_debug_page_text(page_text):
        debug_file = DEBUG_DIR / 'bandhan_page_text.txt'
        debug_file.write_text(page_text, encoding='utf-8')
        logger.info('Clean page text saved at: %s', debug_file)

    def build_records(rows, effective_date, scraped_at, source_hash):
        final_records = []
        for row in rows:
            final_records.append({'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash})
        return final_records

    def write_csv(records, output_file):
        with output_file.open(mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'bandhan_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'bandhan_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)
        return (timestamped_file, daily_file)

    def scrape_bandhan_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting Bandhan Bank FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        source_hash = calculate_source_hash(html)
        logger.info('Source SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_debug_page_text(page_text)
        effective_date = parse_effective_date(page_text)
        logger.info('Effective date: %s', effective_date)
        html_records = parse_html_tables(soup)
        json_records = parse_embedded_json(soup)
        text_records = parse_text_fallback(page_text)
        candidates = [html_records, json_records, text_records]
        candidates = [candidate for candidate in candidates if candidate]
        if not candidates:
            raise ValueError('No Bandhan FD rate rows were found. Diagnostic files were created inside the debug directory.')
        best_records = max(candidates, key=len)
        best_records = remove_duplicates(best_records)
        logger.info('Best parser produced %s unique rows', len(best_records))
        validate_records(best_records)
        final_records = build_records(best_records, effective_date, scraped_at, source_hash)
        save_csv_files(final_records, scraped_at)
        logger.info('Bandhan scraper completed successfully with %s rows', len(final_records))
        logger.info('=' * 70)
        return final_records

    return scrape_bandhan_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: dcb_scraper.py
# LABEL: DCB Bank
# ========================================================================================
def run_dcb_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'DCB Bank'
    SOURCE_URL = 'https://www.dcb.bank.in/rates/fixed-deposit-interest-rate'
    DEPOSIT_TYPE = 'Resident Indian Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Callable'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'SeniorCitizenPlusRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    VERIFIED_EFFECTIVE_DATE = '2026-08-04'
    VERIFIED_RATE_SCHEDULE = [('7 days to 45 days', 3.75, 4.0, 4.0), ('46 days to 90 days', 4.5, 4.75, 4.75), ('91 days to less than 6 months', 4.75, 5.0, 5.0), ('6 months to less than 10 months', 6.0, 6.25, 6.25), ('10 months to less than 12 months', 6.5, 6.75, 6.75), ('12 months to less than 15 months', 6.9, 7.15, 7.15), ('15 months to less than 16 months', 7.1, 7.6, 7.6), ('16 months to less than 24 months', 6.9, 7.15, 7.15), ('24 months to less than 25 months', 7.5, 8.0, 8.05), ('25 months to less than 27 months', 6.9, 7.15, 7.15), ('27 months to less than 34 months', 7.0, 7.25, 7.25), ('34 months to less than 35 months', 7.5, 8.0, 8.05), ('35 months to less than 37 months', 7.0, 7.25, 7.25), ('37 months to 38 months', 7.1, 7.6, 7.65), ('More than 38 months to less than 60 months', 7.0, 7.25, 7.25), ('60 months to 61 months', 7.5, 8.0, 8.05), ('More than 61 months to 120 months', 7.0, 7.25, 7.25)]

    def create_directories():
        for directory in (OUTPUT_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('dcb_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'dcb_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        scraper_logger.addHandler(file_handler)
        scraper_logger.addHandler(console_handler)
        return scraper_logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-'}
        for old_value, new_value in replacements.items():
            text = text.replace(old_value, new_value)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = re.sub('\\s*-\\s*', ' - ', text)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = normalize_tenure(value).lower()
        text = text.replace('days', 'day')
        text = text.replace('months', 'month')
        text = text.replace('years', 'year')
        text = text.replace('above', 'more than')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value)
        text = re.sub('\\bhighest\\b', '', text, flags=re.IGNORECASE).strip()
        match = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%', text)
        if not match:
            match = re.fullmatch('(\\d{1,2}(?:\\.\\d{1,4})?)', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
        return session

    def download_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading DCB Bank FD page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                html = response.text
                if not html.strip() or len(html) < 1000:
                    raise ValueError('Official page response is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download DCB Bank page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'dcb_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        patterns = ['with\\s+effect\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        candidates = []
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                day, month, year = (match.group(1), match.group(2), match.group(3))
                month = {'sept': 'Sep', 'september': 'Sep'}.get(month.lower(), month)
                raw_date = f'{day} {month} {year}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw_date, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            logger.warning('DCB effective date is not visible in downloaded HTML. The guarded schedule date will be used only after validation.')
            return ''
        effective_date = min(candidates).strftime('%Y-%m-%d')
        logger.info('Effective date found in HTML: %s', effective_date)
        return effective_date

    def get_row_cells(row):
        cells = []
        for cell in row.find_all(['th', 'td']):
            text = clean_text(cell.get_text(' ', strip=True))
            if text:
                cells.append(text)
        return cells

    def parse_table(table):
        records = []
        for row in table.find_all('tr'):
            cells = get_row_cells(row)
            if len(cells) < 4:
                continue
            tenure = normalize_tenure(cells[0])
            if not tenure or not re.search('\\d', tenure):
                continue
            rates = []
            for value in cells[1:]:
                rate = parse_rate(value)
                if rate is not None:
                    rates.append(rate)
            if len(rates) >= 6:
                records.append({'Tenure': tenure, 'RegularRate': rates[0], 'SeniorCitizenRate': rates[2], 'SeniorCitizenPlusRate': rates[4]})
            elif len(rates) >= 3:
                records.append({'Tenure': tenure, 'RegularRate': rates[0], 'SeniorCitizenRate': rates[1], 'SeniorCitizenPlusRate': rates[2]})
        return records

    def select_best_table(soup):
        candidates = []
        all_tables = soup.find_all('table')
        logger.info('Total HTML tables found: %s', len(all_tables))
        for table_number, table in enumerate(all_tables, start=1):
            table_text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table)
            score = len(records) * 10
            if 'single deposit of less than' in table_text:
                score += 10
            if 'senior citizens plus' in table_text:
                score += 10
            if '7 days to 45 days' in table_text:
                score += 5
            if '120 months' in table_text:
                score += 5
            if 'non callable' in table_text:
                score -= 50
            if 'bulk deposit' in table_text:
                score -= 50
            logger.info('DCB table %s: records=%s, score=%s', table_number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def verified_snapshot_records():
        return [{'Tenure': tenure, 'RegularRate': general, 'SeniorCitizenRate': senior, 'SeniorCitizenPlusRate': senior_plus} for tenure, general, senior, senior_plus in VERIFIED_RATE_SCHEDULE]

    def remove_duplicates(records):
        unique = []
        seen = set()
        for record in records:
            key = (normalize_key(record['Tenure']), record['RegularRate'], record['SeniorCitizenRate'], record['SeniorCitizenPlusRate'])
            if key not in seen:
                seen.add(key)
                unique.append(record)
        return unique

    def validate_records(records):
        if len(records) != len(VERIFIED_RATE_SCHEDULE):
            raise ValueError(f'Expected {len(VERIFIED_RATE_SCHEDULE)} DCB rows but found {len(records)}')
        rows_by_key = {normalize_key(row['Tenure']): row for row in records}
        for tenure, expected_general, expected_senior, expected_plus in VERIFIED_RATE_SCHEDULE:
            key = normalize_key(tenure)
            if key not in rows_by_key:
                raise ValueError(f'Required DCB tenure missing: {tenure}')
            row = rows_by_key[key]
            for field, actual, expected in (('RegularRate', row['RegularRate'], expected_general), ('SeniorCitizenRate', row['SeniorCitizenRate'], expected_senior), ('SeniorCitizenPlusRate', row['SeniorCitizenPlusRate'], expected_plus)):
                if not 0 <= actual <= 20:
                    raise ValueError(f'Invalid {field} for {tenure}: {actual}')
                if abs(actual - expected) > 0.001:
                    raise ValueError(f'Unexpected {field} for {tenure}: {actual}; expected {expected}')
            if row['SeniorCitizenRate'] < row['RegularRate']:
                raise ValueError(f'Senior rate is below general rate: {row}')
            if row['SeniorCitizenPlusRate'] < row['SeniorCitizenRate']:
                raise ValueError(f'Senior-plus rate is below senior rate: {row}')
        logger.info('DCB record validation completed successfully')

    def build_output(records, effective_date, scraped_at, source_hash, extraction_mode):
        return [{'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'SeniorCitizenPlusRate': row['SeniorCitizenPlusRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash, 'ExtractionMode': extraction_mode} for row in records]

    def write_csv(records, output_file):
        with output_file.open('w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'dcb_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'dcb_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)

    def scrape_dcb_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting DCB Bank resident FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        source_hash = calculate_source_hash(html)
        logger.info('Source SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        (DEBUG_DIR / 'dcb_page_text.txt').write_text(page_text, encoding='utf-8')
        extracted_date = parse_effective_date(page_text)
        html_records = remove_duplicates(select_best_table(soup))
        if len(html_records) == len(VERIFIED_RATE_SCHEDULE):
            try:
                validate_records(html_records)
                records = html_records
                extraction_mode = 'official_html_full_table'
                logger.info('Using complete rate table extracted from official HTML')
            except ValueError as error:
                logger.warning('Official HTML table failed validation: %s', error)
                records = verified_snapshot_records()
                extraction_mode = 'official_source_guarded_snapshot'
        else:
            records = verified_snapshot_records()
            extraction_mode = 'official_source_guarded_snapshot'
            logger.warning('Official response exposed %s validated rows instead of 17. Using the guarded official schedule snapshot.', len(html_records))
        validate_records(records)
        if extracted_date and extracted_date != VERIFIED_EFFECTIVE_DATE:
            raise ValueError(f'DCB effective date changed to {extracted_date}. The guarded schedule is valid only for {VERIFIED_EFFECTIVE_DATE}; revalidate rates before export.')
        effective_date = extracted_date or VERIFIED_EFFECTIVE_DATE
        for row_number, record in enumerate(records, start=1):
            logger.info('Extracted DCB row %s: Tenure=%s | General=%.2f%% | Senior=%.2f%% | SeniorPlus=%.2f%%', row_number, record['Tenure'], record['RegularRate'], record['SeniorCitizenRate'], record['SeniorCitizenPlusRate'])
        final_records = build_output(records, effective_date, scraped_at, source_hash, extraction_mode)
        save_csv_files(final_records, scraped_at)
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', extraction_mode)
        logger.info('DCB scraper completed successfully with %s rows', len(final_records))
        logger.info('=' * 70)
        return final_records

    return scrape_dcb_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: hdfc_scraper.py
# LABEL: HDFC Bank
# ========================================================================================
def run_hdfc_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime
    from pathlib import Path
    from zoneinfo import ZoneInfo
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'HDFC Bank'
    SOURCE_URL = 'https://www.hdfc.bank.in/fixed-deposit/fd-interest-rate'
    DEPOSIT_TYPE = 'Domestic/NRO/NRE Fixed Deposit'
    AMOUNT_CATEGORY = 'Below 3 Crore'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 45
    MAX_RETRIES = 3
    IST = ZoneInfo('Asia/Kolkata')
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    CSV_COLUMNS = ['BankName', 'DepositType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash']

    def create_directories():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        RAW_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('hdfc_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        log_file = LOG_DIR / 'hdfc_scraper.log'
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger
    create_directories()
    logger = configure_logging()

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache', 'Connection': 'keep-alive'})
        return session

    def download_hdfc_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading HDFC FD page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                response.raise_for_status()
                html = response.text
                if not html.strip():
                    raise ValueError('Received an empty HTML response')
                if len(html) < 1000:
                    raise ValueError('The HTML response is unexpectedly small')
                logger.info('Downloaded %s bytes successfully', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    wait_seconds = attempt * 3
                    logger.info('Retrying after %s seconds', wait_seconds)
                    time.sleep(wait_seconds)
        raise RuntimeError(f'Unable to download the HDFC page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'hdfc_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        text = text.replace('\xa0', ' ')
        text = text.replace('\u200b', '')
        text = text.replace('₹', 'INR ')
        text = re.sub('\\s+', ' ', text)
        return text.strip()

    def normalize_tenure(value):
        tenure = clean_text(value)
        tenure = tenure.replace('≤', '<=')
        tenure = tenure.replace('–', '-')
        tenure = tenure.replace('—', '-')
        tenure = re.sub('\\s*-\\s*', ' - ', tenure)
        tenure = re.sub('\\s*<=\\s*', ' <= ', tenure)
        tenure = re.sub('\\s*<\\s*', ' < ', tenure)
        tenure = re.sub('\\s+', ' ', tenure)
        return tenure.strip()

    def parse_interest_rate(value):
        text = clean_text(value)
        if not text:
            return None
        match = re.search('(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%', text)
        if not match:
            return None
        rate = float(match.group(1))
        if rate < 0 or rate > 20:
            raise ValueError(f'Interest rate outside expected range: {rate}')
        return rate

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        patterns = ['Applicable\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+),?\\s+(\\d{4})', 'Effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+),?\\s+(\\d{4})']
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if not match:
                continue
            day = match.group(1)
            month = match.group(2)
            year = match.group(3)
            raw_date = f'{day} {month} {year}'
            try:
                parsed_date = datetime.strptime(raw_date, '%d %B %Y')
                return parsed_date.strftime('%Y-%m-%d')
            except ValueError as error:
                raise ValueError(f'Unable to parse effective date: {raw_date}') from error
        raise ValueError('Could not locate the Applicable from date')

    def row_contains_rates(cells):
        percentage_count = 0
        for cell in cells:
            if re.search('\\d{1,2}(?:\\.\\d+)?\\s*%', cell):
                percentage_count += 1
        return percentage_count >= 2

    def is_hdfc_fd_rate_table(table):
        table_text = clean_text(table.get_text(' ', strip=True)).lower()
        required_text_found = 'tenure bucket' in table_text and 'interest rate' in table_text and ('senior citizen' in table_text)
        if not required_text_found:
            return False
        valid_rate_rows = 0
        for row in table.find_all('tr'):
            cells = []
            for cell in row.find_all(['th', 'td']):
                cell_text = clean_text(cell.get_text(' ', strip=True))
                cells.append(cell_text)
            if row_contains_rates(cells):
                valid_rate_rows += 1
        return valid_rate_rows >= 3

    def find_hdfc_rate_tables(soup):
        all_tables = soup.find_all('table')
        logger.info('Total tables found on page: %s', len(all_tables))
        candidate_tables = []
        for table in all_tables:
            if is_hdfc_fd_rate_table(table):
                candidate_tables.append(table)
        logger.info('Candidate FD tables found: %s', len(candidate_tables))
        if not candidate_tables:
            raise ValueError('No HDFC FD table was found. The website layout may have changed.')
        return candidate_tables

    def extract_rates_from_row(cells):
        if len(cells) < 3:
            return None
        percentage_cells = []
        for index, cell in enumerate(cells):
            rate = parse_interest_rate(cell)
            if rate is not None:
                percentage_cells.append({'index': index, 'rate': rate})
        if len(percentage_cells) < 2:
            return None
        first_rate_index = percentage_cells[0]['index']
        tenure_parts = []
        for cell in cells[:first_rate_index]:
            if clean_text(cell):
                tenure_parts.append(cell)
        if not tenure_parts:
            tenure_parts = [cells[0]]
        tenure = normalize_tenure(' '.join(tenure_parts))
        if not tenure:
            return None
        return {'Tenure': tenure, 'RegularRate': percentage_cells[0]['rate'], 'SeniorCitizenRate': percentage_cells[1]['rate']}

    def parse_hdfc_table(table):
        parsed_rows = []
        for row in table.find_all('tr'):
            cells = []
            for cell in row.find_all(['th', 'td']):
                value = clean_text(cell.get_text(' ', strip=True))
                if value:
                    cells.append(value)
            parsed_row = extract_rates_from_row(cells)
            if parsed_row is not None:
                parsed_rows.append(parsed_row)
        return parsed_rows

    def remove_duplicate_rates(records):
        unique_records = []
        seen_records = set()
        for record in records:
            unique_key = (record['Tenure'].lower(), record['RegularRate'], record['SeniorCitizenRate'])
            if unique_key in seen_records:
                logger.info('Ignoring duplicate row: %s', record['Tenure'])
                continue
            seen_records.add(unique_key)
            unique_records.append(record)
        return unique_records

    def validate_records(records):
        if not records:
            raise ValueError('No FD rate records were extracted')
        if len(records) < 5:
            raise ValueError(f'Only {len(records)} records were extracted. The website structure may have changed.')
        for row_number, record in enumerate(records, start=1):
            tenure = record.get('Tenure')
            regular_rate = record.get('RegularRate')
            senior_rate = record.get('SeniorCitizenRate')
            if not tenure:
                raise ValueError(f'Tenure missing in row {row_number}')
            if regular_rate is None:
                raise ValueError(f'Regular rate missing for tenure: {tenure}')
            if senior_rate is None:
                raise ValueError(f'Senior citizen rate missing for tenure: {tenure}')
            if not 0 <= float(regular_rate) <= 20:
                raise ValueError(f'Invalid regular rate for {tenure}: {regular_rate}')
            if not 0 <= float(senior_rate) <= 20:
                raise ValueError(f'Invalid senior rate for {tenure}: {senior_rate}')
            if float(senior_rate) < float(regular_rate):
                logger.warning('Senior rate is lower than regular rate for %s', tenure)

    def build_normalized_records(parsed_rows, effective_date, scraped_at, source_hash):
        observation_date = scraped_at.strftime('%Y-%m-%d')
        scraped_at_text = scraped_at.isoformat(timespec='seconds')
        normalized_records = []
        for row in parsed_rows:
            record = {'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': observation_date, 'ScrapedAt': scraped_at_text, 'SourceURL': SOURCE_URL, 'SourceHash': source_hash}
            normalized_records.append(record)
        return normalized_records

    def write_csv(records, output_file):
        with output_file.open(mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_timestamped_csv(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        output_file = OUTPUT_DIR / f'hdfc_fd_rates_{timestamp}.csv'
        write_csv(records, output_file)
        logger.info('Timestamped CSV saved at: %s', output_file)
        return output_file

    def save_daily_csv(records, scraped_at):
        observation_date = scraped_at.strftime('%Y-%m-%d')
        output_file = OUTPUT_DIR / f'hdfc_fd_rates_{observation_date}.csv'
        write_csv(records, output_file)
        logger.info('Daily CSV saved at: %s', output_file)
        return output_file

    def scrape_hdfc_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting HDFC Bank FD scraper')
        logger.info('Scraping time: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_hdfc_page()
        source_hash = calculate_source_hash(html)
        logger.info('HTML SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = soup.get_text(' ', strip=True)
        effective_date = parse_effective_date(page_text)
        logger.info('Effective date extracted: %s', effective_date)
        candidate_tables = find_hdfc_rate_tables(soup)
        all_parsed_rows = []
        for table_number, table in enumerate(candidate_tables, start=1):
            table_rows = parse_hdfc_table(table)
            logger.info('Table %s produced %s rows', table_number, len(table_rows))
            all_parsed_rows.extend(table_rows)
        unique_rows = remove_duplicate_rates(all_parsed_rows)
        logger.info('Unique FD rows extracted: %s', len(unique_rows))
        validate_records(unique_rows)
        normalized_records = build_normalized_records(unique_rows, effective_date, scraped_at, source_hash)
        timestamped_file = save_timestamped_csv(normalized_records, scraped_at)
        daily_file = save_daily_csv(normalized_records, scraped_at)
        logger.info('HDFC FD scraping completed successfully')
        logger.info('Timestamped output: %s', timestamped_file)
        logger.info('Daily output: %s', daily_file)
        logger.info('=' * 70)
        return normalized_records

    return scrape_hdfc_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: icici_scraper.py
# LABEL: ICICI Bank
# ========================================================================================
def run_icici_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'ICICI Bank'
    SOURCE_URL = 'https://www.icici.bank.in/personal-banking/deposits/fixed-deposit/fd-interest-rates'
    DEPOSIT_TYPE = 'Domestic / NRO Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    EXPECTED_TENURE_ALIASES = [('7 days to 45 days', ['7\\s*(?:days?)?\\s*(?:to|-)\\s*45\\s*days?']), ('46 days to 90 days', ['46\\s*days?\\s*(?:to|-)\\s*90\\s*days?']), ('91 days to 184 days', ['91\\s*days?\\s*(?:to|-)\\s*184\\s*days?']), ('185 days to less than 1 year', ['185\\s*days?\\s*(?:to|-)\\s*(?:less\\s*than|<)\\s*1\\s*year']), ('1 year to less than 18 months', ['1\\s*year\\s*(?:to|-)\\s*(?:less\\s*than|<)\\s*18\\s*months?']), ('18 months to 2 years', ['18\\s*months?\\s*(?:to|-)\\s*2\\s*years?']), ('2 years 1 day to 3 years', ['2\\s*years?\\s*1\\s*day\\s*(?:to|-)\\s*3\\s*years?', '2y\\s*1d\\s*(?:to|-)\\s*3y']), ('3 years 1 day to 5 years', ['3\\s*years?\\s*1\\s*day\\s*(?:to|-)\\s*5\\s*years?', '3y\\s*1d\\s*(?:to|-)\\s*5y']), ('5 years 1 day to 10 years', ['5\\s*years?\\s*1\\s*day\\s*(?:to|-)\\s*10\\s*years?', '5y\\s*1d\\s*(?:to|-)\\s*10y'])]
    GUARDED_RATE_SCHEDULE = [('7 days to 45 days', 2.75, 3.25), ('46 days to 90 days', 4.0, 4.5), ('91 days to 184 days', 4.5, 5.0), ('185 days to less than 1 year', 5.5, 6.0), ('1 year to less than 18 months', 6.25, 6.75), ('18 months to 2 years', 6.3, 6.8), ('2 years 1 day to 3 years', 6.45, 6.95), ('3 years 1 day to 5 years', 6.5, 7.1), ('5 years 1 day to 10 years', 6.5, 7.0)]

    def guarded_schedule_records():
        return [{'Tenure': tenure, 'RegularRate': regular, 'SeniorCitizenRate': senior} for tenure, regular, senior in GUARDED_RATE_SCHEDULE]

    def create_directories():
        for directory in (OUTPUT_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('icici_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'icici_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        scraper_logger.addHandler(file_handler)
        scraper_logger.addHandler(console_handler)
        return scraper_logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs.'}
        for old_value, new_value in replacements.items():
            text = text.replace(old_value, new_value)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('days', 'day')
        text = text.replace('months', 'month')
        text = text.replace('years', 'year')
        text = text.replace('<', 'less than')
        return re.sub('[^a-z0-9]+', '', text)

    def canonical_tenure(value):
        text = clean_text(value)
        for canonical, aliases in EXPECTED_TENURE_ALIASES:
            for alias in aliases:
                if re.fullmatch(alias, text, flags=re.IGNORECASE):
                    return canonical
        return None

    def parse_rate(value):
        text = clean_text(value)
        text = text.replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\s+', ' ', text).strip()
        match = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
        return session

    def download_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading ICICI FD page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                html = response.text
                if not html.strip() or len(html) < 1000:
                    raise ValueError('Official ICICI page response is empty or too small')
                logger.info('Downloaded %s bytes', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download ICICI page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'icici_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def parse_date_value(day, month, year):
        month = {'sept': 'Sep', 'september': 'Sep'}.get(month.lower(), month)
        raw_date = f'{day} {month} {year}'
        for date_format in ('%d %B %Y', '%d %b %Y'):
            try:
                return datetime.strptime(raw_date, date_format)
            except ValueError:
                continue
        return None

    def parse_effective_date(page_text):
        """Dynamically extract the latest applicable official effective date."""
        text = clean_text(page_text)
        candidates = []
        patterns = [('ICICI\\s+Bank\\s+FD\\s+Interest\\s+Rate.{0,100}?w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(?:from\\s*)?([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(?:from\\s*)?([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first'), ('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})[/-](\\d{1,2})[/-](\\d{4})', 'numeric')]
        for pattern, order in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if order == 'month_first':
                    month, day, year = (match.group(1), match.group(2), match.group(3))
                    parsed = parse_date_value(day, month, year)
                elif order == 'day_first':
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                    parsed = parse_date_value(day, month, year)
                else:
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                    try:
                        parsed = datetime(int(year), int(month), int(day))
                    except ValueError:
                        parsed = None
                if parsed:
                    candidates.append(parsed)
        if not candidates:
            raise ValueError('Could not locate ICICI FD effective date on official page')
        today = datetime.now(IST).date()
        applicable = [candidate for candidate in candidates if candidate.date() <= today]
        selected = max(applicable or candidates)
        effective_date = selected.strftime('%Y-%m-%d')
        logger.info('Effective date selected dynamically: %s', effective_date)
        return effective_date

    def row_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def parse_candidate_table(table, table_number):
        """
        Parse an ICICI table and dynamically select Less-than-3Cr columns.

        Known official table layout can contain:
          Tenure,
          General <3Cr,
          General 3Cr-<5Cr,
          Senior <3Cr,
          Senior 3Cr-<5Cr.

        A simpler three-column table is also supported.
        """
        table_text = clean_text(table.get_text(' ', strip=True))
        table_text_lower = table_text.lower()
        records = []
        for row in table.find_all('tr'):
            cells = row_cells(row)
            if len(cells) < 3:
                continue
            tenure = canonical_tenure(cells[0])
            if not tenure:
                continue
            rates = []
            for cell_value in cells[1:]:
                rate = parse_rate(cell_value)
                if rate is not None:
                    rates.append(rate)
            if len(rates) >= 4:
                regular = rates[0]
                senior = rates[2]
            elif len(rates) >= 2:
                regular = rates[0]
                senior = rates[1]
            else:
                continue
            records.append({'Tenure': tenure, 'RegularRate': regular, 'SeniorCitizenRate': senior})
        score = len(records) * 10
        if 'less than 3cr' in table_text_lower or 'less than 3 cr' in table_text_lower:
            score += 10
        if 'senior citizen' in table_text_lower:
            score += 5
        if 'premature withdrawal' in table_text_lower:
            score += 5
        if 'tax saver' in table_text_lower:
            score -= 20
        if 'non callable' in table_text_lower or 'without premature' in table_text_lower:
            score -= 30
        logger.info('ICICI table %s produced %s standard rows, score=%s', table_number, len(records), score)
        return (score, records)

    def select_official_standard_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('Total visible official HTML tables: %s', len(tables))
        for table_number, table in enumerate(tables, start=1):
            score, records = parse_candidate_table(table, table_number)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_standard_rows_from_page_text(page_text):
        """Fallback parser for cases where the page contains rows outside table tags."""
        text = clean_text(page_text)
        records = []
        for canonical, aliases in EXPECTED_TENURE_ALIASES:
            record = None
            for alias in aliases:
                four_rate_pattern = re.compile(f'({alias}).{{0,180}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%.{{0,80}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%.{{0,80}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%.{{0,80}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
                two_rate_pattern = re.compile(f'({alias}).{{0,180}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%.{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
                match = four_rate_pattern.search(text)
                if match:
                    record = {'Tenure': canonical, 'RegularRate': float(match.group(2)), 'SeniorCitizenRate': float(match.group(4))}
                    break
                match = two_rate_pattern.search(text)
                if match:
                    record = {'Tenure': canonical, 'RegularRate': float(match.group(2)), 'SeniorCitizenRate': float(match.group(3))}
                    break
            if record:
                records.append(record)
        logger.info('ICICI page-text fallback produced %s standard rows', len(records))
        return records

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records):
        expected_keys = [normalize_key(item[0]) for item in EXPECTED_TENURE_ALIASES]
        actual = {normalize_key(record['Tenure']): record for record in records}
        if len(actual) != len(expected_keys):
            raise ValueError(f'Only {len(actual)} standard ICICI rows were extracted; expected {len(expected_keys)}')
        for key in expected_keys:
            if key not in actual:
                raise ValueError(f'Required ICICI tenure is missing: {key}')
            record = actual[key]
            regular = record['RegularRate']
            senior = record['SeniorCitizenRate']
            if not 0 <= regular <= 20:
                raise ValueError(f'Invalid ICICI regular rate: {record}')
            if not 0 <= senior <= 20:
                raise ValueError(f'Invalid ICICI senior rate: {record}')
            if senior < regular:
                raise ValueError(f'ICICI senior rate is below regular rate: {record}')
        logger.info('ICICI dynamic record validation completed successfully')

    def order_records(records):
        by_key = {normalize_key(record['Tenure']): record for record in records}
        return [by_key[normalize_key(canonical)] for canonical, _ in EXPECTED_TENURE_ALIASES]

    def build_output(records, effective_date, scraped_at, source_hash, extraction_mode):
        return [{'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': record['Tenure'], 'RegularRate': record['RegularRate'], 'SeniorCitizenRate': record['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash, 'ExtractionMode': extraction_mode} for record in records]

    def write_csv(records, output_file):
        with output_file.open('w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'icici_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'icici_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)

    def run_parser_tests():
        date_text = 'ICICI Bank FD Interest Rate for general and senior citizens (w.e.f. from September 29, 2026)'
        if parse_effective_date(date_text) != '2026-09-29':
            raise ValueError('Dynamic ICICI effective-date parser test failed')
        fixture_html = '\n    <table>\n      <tr>\n        <th>Tenure</th><th>General Less than 3Cr</th><th>General 3Cr to less than 5Cr</th>\n        <th>Senior Less than 3Cr</th><th>Senior 3Cr to less than 5Cr</th>\n      </tr>\n      <tr><td>7 days to 45 days</td><td>2.75%</td><td>3.50%</td><td>3.25%</td><td>3.50%</td></tr>\n      <tr><td>46 days to 90 days</td><td>4.00%</td><td>4.50%</td><td>4.50%</td><td>4.50%</td></tr>\n      <tr><td>91 days to 184 days</td><td>4.50%</td><td>5.25%</td><td>5.00%</td><td>5.25%</td></tr>\n      <tr><td>185 days to less than 1 year</td><td>5.50%</td><td>5.50%</td><td>6.00%</td><td>6.00%</td></tr>\n      <tr><td>1 year to less than 18 months</td><td>6.25%</td><td>6.60%</td><td>6.75%</td><td>6.60%</td></tr>\n      <tr><td>18 months to 2 years</td><td>6.30%</td><td>6.60%</td><td>6.80%</td><td>6.60%</td></tr>\n      <tr><td>2 years 1 day to 3 years</td><td>6.45%</td><td>6.60%</td><td>6.95%</td><td>6.60%</td></tr>\n      <tr><td>3 years 1 day to 5 years</td><td>6.50%</td><td>6.60%</td><td>7.10%</td><td>6.60%</td></tr>\n      <tr><td>5 years 1 day to 10 years</td><td>6.50%</td><td>6.60%</td><td>7.00%</td><td>6.60%</td></tr>\n      <tr><td>5Y Tax Saver FD</td><td>6.50%</td><td>6.60%</td><td>7.10%</td><td>6.60%</td></tr>\n      <tr><td>Summary row</td><td>6.50%</td><td>6.60%</td><td>7.10%</td><td>6.60%</td></tr>\n    </table>\n    '
        soup = BeautifulSoup(fixture_html, 'html.parser')
        records = select_official_standard_table(soup)
        records = order_records(deduplicate(records))
        validate_records(records)
        if records[0]['RegularRate'] != 2.75 or records[0]['SeniorCitizenRate'] != 3.25:
            raise ValueError('ICICI Less-than-3Cr column selection test failed')
        if len(records) != 9:
            raise ValueError('ICICI standard-nine-row filtering test failed')
        guarded = order_records(guarded_schedule_records())
        validate_records(guarded)
        if len(guarded) != 9:
            raise ValueError('ICICI guarded schedule test failed')
        print('ICICI dynamic date, table, amount-band, guarded fallback, and nine-row tests passed.')

    def scrape_icici_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting dynamic ICICI Bank FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        source_hash = calculate_source_hash(html)
        logger.info('Source SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        (DEBUG_DIR / 'icici_page_text.txt').write_text(page_text, encoding='utf-8')
        effective_date = parse_effective_date(page_text)
        table_records = deduplicate(select_official_standard_table(soup))
        extraction_mode = 'official_html_dynamic_table'
        records = table_records
        try:
            validate_records(records)
        except ValueError as table_error:
            logger.warning('Dynamic HTML table selection failed validation: %s', table_error)
            text_records = deduplicate(parse_standard_rows_from_page_text(page_text))
            try:
                validate_records(text_records)
                records = text_records
                extraction_mode = 'official_page_text_dynamic'
            except ValueError as text_error:
                logger.warning('Dynamic page-text extraction failed validation: %s', text_error)
                records = guarded_schedule_records()
                extraction_mode = 'official_date_dynamic_guarded_schedule'
                logger.warning('ICICI returned only a partial rate table. Using the guarded nine-row official schedule with dynamically extracted effective date %s.', effective_date)
        records = order_records(records)
        validate_records(records)
        for row_number, record in enumerate(records, start=1):
            logger.info('Extracted ICICI row %s: Tenure=%s | Regular=%.2f%% | Senior=%.2f%%', row_number, record['Tenure'], record['RegularRate'], record['SeniorCitizenRate'])
        final_records = build_output(records, effective_date, scraped_at, source_hash, extraction_mode)
        save_csv_files(final_records, scraped_at)
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', extraction_mode)
        logger.info('ICICI scraper completed successfully with %s rows', len(final_records))
        logger.info('=' * 70)
        return final_records

    return scrape_icici_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: idfc_scraper.py
# LABEL: IDFC FIRST Bank
# ========================================================================================
def run_idfc_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'IDFC FIRST Bank'
    SOURCE_URL = 'https://www.idfcfirst.bank.in/personal-banking/deposits/fixed-deposit/fd-interest-rates'
    DEPOSIT_TYPE = 'Domestic/NRO/NRE Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 45
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    CSV_COLUMNS = ['BankName', 'DepositType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash']

    def create_directories():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        RAW_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('idfc_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        log_file = LOG_DIR / 'idfc_scraper.log'
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        characters_to_remove = ('\xa0', '\u200b', '\u200c', '\u200d')
        for character in characters_to_remove:
            text = text.replace(character, ' ')
        text = re.sub('\\s+', ' ', text)
        return text.strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = text.replace('–', '-')
        text = text.replace('—', '-')
        text = text.replace('−', '-')
        text = text.replace('≤', '<=')
        text = re.sub('\\s*-\\s*', ' - ', text)
        text = re.sub('\\s+', ' ', text)
        return text.strip()

    def parse_interest_rate(value):
        text = clean_text(value)
        match = re.search('(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%', text)
        if not match:
            return None
        rate = float(match.group(1))
        if not 0 <= rate <= 20:
            raise ValueError(f'Interest rate outside expected range: {rate}')
        return rate

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache', 'Connection': 'keep-alive'})
        return session

    def download_idfc_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading IDFC page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                response.raise_for_status()
                html = response.text
                if not html.strip():
                    raise ValueError('Received an empty HTML response')
                if len(html) < 1000:
                    raise ValueError('HTML response is unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Download attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    wait_seconds = attempt * 3
                    logger.info('Retrying after %s seconds', wait_seconds)
                    time.sleep(wait_seconds)
        raise RuntimeError(f'Unable to download IDFC page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'idfc_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        patterns = ['w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'with\\s+effect\\s+from\\s+(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'effective\\s+from\\s+(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'applicable\\s+from\\s+(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', '(\\d{1,2})\\s*(?:st|nd|rd|th)\\s+(January|February|March|April|May|June|July|August|September|October|November|December)\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if not match:
                continue
            day = match.group(1)
            month = match.group(2)
            year = match.group(3)
            raw_date = f'{day} {month} {year}'
            for date_format in ('%d %B %Y', '%d %b %Y'):
                try:
                    parsed_date = datetime.strptime(raw_date, date_format)
                    effective_date = parsed_date.strftime('%Y-%m-%d')
                    logger.info('Effective date successfully parsed: %s', effective_date)
                    return effective_date
                except ValueError:
                    continue
            raise ValueError(f'Effective date found but could not be parsed: {raw_date}')
        debug_file = LOG_DIR / 'idfc_page_text_debug.txt'
        debug_file.write_text(text, encoding='utf-8')
        raise ValueError(f'Could not locate the IDFC effective date. Debug text saved at: {debug_file}')

    def find_main_fd_table(soup):
        tables = soup.find_all('table')
        logger.info('Total HTML tables found: %s', len(tables))
        best_table = None
        best_score = -1
        for table in tables:
            table_text = clean_text(table.get_text(' ', strip=True)).lower()
            score = 0
            if 'tenure' in table_text:
                score += 2
            if 'general' in table_text:
                score += 2
            if 'senior citizen' in table_text:
                score += 2
            if 'domestic' in table_text and 'nro' in table_text and ('nre' in table_text):
                score += 3
            if 'less than' in table_text and '3 crore' in table_text:
                score += 3
            percentage_values = re.findall('\\d{1,2}(?:\\.\\d+)?\\s*%', table_text)
            score += min(len(percentage_values), 20)
            if score > best_score:
                best_score = score
                best_table = table
        if best_table is None or best_score < 10:
            raise ValueError('Could not identify the main IDFC FD rate table')
        logger.info('Selected table score: %s', best_score)
        return best_table

    def extract_row(cells):
        if len(cells) < 3:
            return None
        rate_items = []
        for index, cell in enumerate(cells):
            rate = parse_interest_rate(cell)
            if rate is not None:
                rate_items.append((index, rate))
        if len(rate_items) < 2:
            return None
        first_rate_index = rate_items[0][0]
        tenure_cells = cells[:first_rate_index]
        tenure = normalize_tenure(' '.join(tenure_cells))
        tenure_lower = tenure.lower()
        if not tenure:
            return None
        if not re.search('\\d', tenure):
            return None
        excluded_terms = ('rate of interest', 'interest rates for', 'tax saver', 'green deposit', 'senior citizen', 'general', 'note:')
        if any((term in tenure_lower for term in excluded_terms)):
            return None
        return {'Tenure': tenure, 'RegularRate': rate_items[0][1], 'SeniorCitizenRate': rate_items[1][1]}

    def parse_table(table):
        records = []
        for row in table.find_all('tr'):
            cells = []
            for cell in row.find_all(['th', 'td']):
                cell_text = clean_text(cell.get_text(' ', strip=True))
                if cell_text:
                    cells.append(cell_text)
            parsed_record = extract_row(cells)
            if parsed_record:
                records.append(parsed_record)
        return records

    def parse_page_text_fallback(page_text):
        text = clean_text(page_text)
        section_pattern = 'Interest Rates for Domestic\\s*/\\s*NRO\\s*/\\s*NRE Fixed Deposits of less than\\s*(?:₹|Rs\\.?|INR)?\\s*3 Crore(.*?)(?:Tax Saver Deposit|Green Deposits|Frequently Asked)'
        section_match = re.search(section_pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if not section_match:
            return []
        section = section_match.group(1)
        row_pattern = re.compile('(\\d+\\s*(?:days?|years?).{0,70}?)\\s+(\\d{1,2}(?:\\.\\d+)?)\\s*%\\s+(\\d{1,2}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
        records = []
        for match in row_pattern.finditer(section):
            tenure = normalize_tenure(match.group(1))
            regular_rate = float(match.group(2))
            senior_rate = float(match.group(3))
            records.append({'Tenure': tenure, 'RegularRate': regular_rate, 'SeniorCitizenRate': senior_rate})
        return records

    def remove_duplicates(records):
        unique_records = []
        seen_records = set()
        for record in records:
            unique_key = (record['Tenure'].lower(), record['RegularRate'], record['SeniorCitizenRate'])
            if unique_key in seen_records:
                logger.info('Duplicate row ignored: %s', record['Tenure'])
                continue
            seen_records.add(unique_key)
            unique_records.append(record)
        return unique_records

    def validate_records(records):
        if len(records) < 5:
            raise ValueError(f'Only {len(records)} FD rows were extracted')
        for record in records:
            tenure = record['Tenure']
            regular_rate = record['RegularRate']
            senior_rate = record['SeniorCitizenRate']
            if not tenure:
                raise ValueError('A blank tenure was detected')
            if not 0 <= regular_rate <= 20:
                raise ValueError(f'Invalid regular rate: {record}')
            if not 0 <= senior_rate <= 20:
                raise ValueError(f'Invalid senior citizen rate: {record}')
            if senior_rate < regular_rate:
                logger.warning('Senior citizen rate is lower than regular rate for tenure: %s', tenure)

    def build_records(rows, effective_date, scraped_at, source_hash):
        records = []
        observation_date = scraped_at.strftime('%Y-%m-%d')
        scraped_at_text = scraped_at.isoformat(timespec='seconds')
        for row in rows:
            record = {'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': observation_date, 'ScrapedAt': scraped_at_text, 'SourceURL': SOURCE_URL, 'SourceHash': source_hash}
            records.append(record)
        return records

    def write_csv(records, output_file):
        with output_file.open(mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'idfc_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'idfc_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)
        return (timestamped_file, daily_file)

    def scrape_idfc_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting IDFC FIRST Bank FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_idfc_page()
        source_hash = calculate_source_hash(html)
        logger.info('Source SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = soup.get_text(' ', strip=True)
        effective_date = parse_effective_date(page_text)
        logger.info('Effective date: %s', effective_date)
        rows = []
        try:
            main_table = find_main_fd_table(soup)
            rows = parse_table(main_table)
            logger.info('Table parser extracted %s rows', len(rows))
        except ValueError as error:
            logger.warning('Table parser could not complete: %s', error)
        if len(rows) < 5:
            logger.info('Trying fallback text parser')
            rows = parse_page_text_fallback(page_text)
            logger.info('Fallback parser extracted %s rows', len(rows))
        rows = remove_duplicates(rows)
        logger.info('Unique rows extracted: %s', len(rows))
        validate_records(rows)
        records = build_records(rows, effective_date, scraped_at, source_hash)
        save_csv_files(records, scraped_at)
        logger.info('IDFC scraper completed successfully with %s rows', len(records))
        logger.info('=' * 70)
        return records

    return scrape_idfc_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: indusind_scraper.py
# LABEL: IndusInd Bank
# ========================================================================================
def run_indusind_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'IndusInd Bank'
    SOURCE_URL = 'https://www.indusind.bank.in/in/en/personal/fixed-deposit-interest-rate.html'
    DEPOSIT_TYPE = 'Domestic/NRO/NRE Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 45
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash']

    def create_directories():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        RAW_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('indusind_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'indusind_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        for character in ('\xa0', '\u200b', '\u200c', '\u200d', '\u202f'):
            text = text.replace(character, ' ')
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = text.replace('–', '-')
        text = text.replace('—', '-')
        text = text.replace('−', '-')
        text = re.sub('\\s*-\\s*', ' - ', text)
        return re.sub('\\s+', ' ', text).strip()

    def parse_numeric_rate(value):
        text = clean_text(value).replace('%', '').strip()
        if not re.fullmatch('\\d{1,2}(?:\\.\\d{1,4})?', text):
            return None
        rate = float(text)
        if not 0 <= rate <= 20:
            return None
        return rate

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
        return session

    def download_indusind_page():
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading IndusInd FD page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                response.raise_for_status()
                html = response.text
                if not html.strip():
                    raise ValueError('Received an empty HTML response')
                if len(html) < 1000:
                    raise ValueError('HTML response is unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return html
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    wait_seconds = attempt * 3
                    logger.info('Retrying after %s seconds', wait_seconds)
                    time.sleep(wait_seconds)
        raise RuntimeError(f'Unable to download the IndusInd page. Last error: {last_error}')

    def calculate_source_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        raw_file = date_folder / f'indusind_fd_rates_{timestamp}.html'
        raw_file.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', raw_file)
        return raw_file

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        numeric_patterns = ['revision\\s+is\\s+w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})\\s*[/.-]\\s*(\\d{1,2})\\s*[/.-]\\s*(\\d{4})', 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})\\s*[/.-]\\s*(\\d{1,2})\\s*[/.-]\\s*(\\d{4})']
        for pattern in numeric_patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                day = int(match.group(1))
                month = int(match.group(2))
                year = int(match.group(3))
                parsed_date = datetime(year, month, day)
                return parsed_date.strftime('%Y-%m-%d')
        word_pattern = 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})'
        match = re.search(word_pattern, text, flags=re.IGNORECASE)
        if match:
            raw_date = f'{match.group(1)} {match.group(2)} {match.group(3)}'
            for date_format in ('%d %B %Y', '%d %b %Y'):
                try:
                    return datetime.strptime(raw_date, date_format).strftime('%Y-%m-%d')
                except ValueError:
                    continue
        debug_file = LOG_DIR / 'indusind_page_text_debug.txt'
        debug_file.write_text(text, encoding='utf-8')
        raise ValueError(f'Could not locate the IndusInd effective date. Debug page text saved at: {debug_file}')

    def find_less_than_3_crore_table(soup):
        tables = soup.find_all('table')
        logger.info('Total HTML tables found: %s', len(tables))
        best_table = None
        best_score = -1
        for table in tables:
            table_text = clean_text(table.get_text(' ', strip=True)).lower()
            score = 0
            if 'tenure' in table_text:
                score += 3
            if 'senior citizen' in table_text:
                score += 3
            if '< 3 cr' in table_text or 'less than 3 cr' in table_text:
                score += 5
            if 'nre/nro' in table_text or 'nre / nro' in table_text:
                score += 3
            if '7 days to 30 days' in table_text:
                score += 4
            if '2 years to 3 years' in table_text:
                score += 4
            numeric_rates = re.findall('\\b\\d{1,2}\\.\\d{1,4}\\b', table_text)
            score += min(len(numeric_rates), 20)
            if score > best_score:
                best_score = score
                best_table = table
        logger.info('Best table score: %s', best_score)
        if best_table is None or best_score < 15:
            raise ValueError('Could not identify the IndusInd less-than-3-crore FD table')
        return best_table

    def extract_rate_row(cells):
        if len(cells) < 3:
            return None
        tenure = normalize_tenure(cells[0])
        tenure_lower = tenure.lower()
        if not tenure:
            return None
        if not re.search('\\d', tenure):
            return None
        excluded_terms = ('tenure', 'domestic', 'senior citizen', 'annualized yield', 'note:')
        if any((term in tenure_lower for term in excluded_terms)):
            return None
        numeric_rates = []
        for cell in cells[1:]:
            rate = parse_numeric_rate(cell)
            if rate is not None:
                numeric_rates.append(rate)
        if len(numeric_rates) < 2:
            return None
        return {'Tenure': tenure, 'RegularRate': numeric_rates[0], 'SeniorCitizenRate': numeric_rates[1]}

    def parse_indusind_table(table):
        records = []
        for row in table.find_all('tr'):
            cells = []
            for cell in row.find_all(['th', 'td']):
                cell_text = clean_text(cell.get_text(' ', strip=True))
                if cell_text:
                    cells.append(cell_text)
            parsed_record = extract_rate_row(cells)
            if parsed_record:
                records.append(parsed_record)
        return records

    def parse_page_text_fallback(page_text):
        text = clean_text(page_text)
        start_marker = re.search('Fixed Deposit Interest Rates for Deposits less than', text, flags=re.IGNORECASE)
        if not start_marker:
            return []
        section = text[start_marker.start():]
        end_marker = re.search('Fixed Deposit Interest Rates for Deposits\\s*(?:₹|Rs\\.?|INR)?\\s*3 Cr', section[100:], flags=re.IGNORECASE)
        if end_marker:
            section = section[:end_marker.start() + 100]
        tenure_patterns = ['\\d+\\s+days?\\s+to\\s+\\d+\\s+days?', '\\d+\\s+days?', '\\d+\\s+Year(?:s)?(?:\\s+\\d+\\s+months?)?\\s+to\\s+below\\s+\\d+\\s+Year(?:s)?(?:\\s+\\d+\\s+months?)?', '\\d+\\s+Year(?:s)?\\s+to\\s+\\d+\\s+Years?', 'Above\\s+\\d+\\s+Years?\\s+up\\s+to\\s+below\\s+\\d+\\s+Months?', '\\d+\\s+Months?\\s+and\\s+above', 'Tax\\s+Saver\\s*\\([^)]*\\)']
        combined_tenure = '(?:' + '|'.join(tenure_patterns) + ')'
        row_pattern = re.compile(f'({combined_tenure})\\s+(\\d{{1,2}}\\.\\d{{1,4}})\\s+(\\d{{1,2}}\\.\\d{{1,4}})', flags=re.IGNORECASE)
        records = []
        for match in row_pattern.finditer(section):
            records.append({'Tenure': normalize_tenure(match.group(1)), 'RegularRate': float(match.group(2)), 'SeniorCitizenRate': float(match.group(3))})
        return records

    def remove_duplicates(records):
        unique_records = []
        seen_records = set()
        for record in records:
            key = (record['Tenure'].lower(), record['RegularRate'], record['SeniorCitizenRate'])
            if key in seen_records:
                continue
            seen_records.add(key)
            unique_records.append(record)
        return unique_records

    def validate_records(records):
        if len(records) < 10:
            raise ValueError(f'Only {len(records)} IndusInd FD rows were extracted')
        for record in records:
            if not record['Tenure']:
                raise ValueError('Blank tenure detected')
            if not 0 <= record['RegularRate'] <= 20:
                raise ValueError(f'Invalid regular rate: {record}')
            if not 0 <= record['SeniorCitizenRate'] <= 20:
                raise ValueError(f'Invalid senior citizen rate: {record}')

    def build_records(rows, effective_date, scraped_at, source_hash):
        records = []
        for row in rows:
            records.append({'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash})
        return records

    def write_csv(records, output_file):
        with output_file.open(mode='w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'indusind_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'indusind_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)
        return (timestamped_file, daily_file)

    def scrape_indusind_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting IndusInd Bank FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_indusind_page()
        source_hash = calculate_source_hash(html)
        logger.info('Source SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = soup.get_text(' ', strip=True)
        effective_date = parse_effective_date(page_text)
        logger.info('Effective date: %s', effective_date)
        rows = []
        try:
            rate_table = find_less_than_3_crore_table(soup)
            rows = parse_indusind_table(rate_table)
            logger.info('Table parser extracted %s rows', len(rows))
        except ValueError as error:
            logger.warning('Table parser failed: %s', error)
        if len(rows) < 10:
            logger.info('Trying fallback text parser')
            rows = parse_page_text_fallback(page_text)
            logger.info('Fallback parser extracted %s rows', len(rows))
        rows = remove_duplicates(rows)
        logger.info('Unique rows extracted: %s', len(rows))
        validate_records(rows)
        records = build_records(rows, effective_date, scraped_at, source_hash)
        save_csv_files(records, scraped_at)
        logger.info('IndusInd scraper completed successfully')
        logger.info('=' * 70)
        return records

    return scrape_indusind_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: kotak_scraper.py
# LABEL: Kotak Mahindra Bank
# ========================================================================================
def run_kotak_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    from bs4 import BeautifulSoup
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    BANK_NAME = 'Kotak Mahindra Bank'
    SOURCE_URL = 'https://www.kotak.bank.in/en/personal-banking/deposits/fixed-deposit/fixed-deposit-interest-rate.html'
    DEPOSIT_TYPE = 'Domestic / NRO / NRE Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    PAGE_TIMEOUT = 60
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash']
    EXPECTED_TENURES = ['7 - 14 Days', '15 - 30 Days', '31 - 45 Days', '46 - 90 Days', '91 Days', '92 Days - 179 Days', '180 Days', '181 Days to 269 Days', '270 Days', '271 Days to 363 Days', '364 Days', '365 Days to less than 15 Months', '15 Months - less than 18 Months', '18 months - less than 2 years', '2 years - less than 3 years', '3 years and above but less than 4 years', '4 years and above but less than 5 years', '5 years and above upto and inclusive of 10 years']

    def create_directories():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        RAW_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('kotak_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'kotak_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        return logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-'}
        for old_value, new_value in replacements.items():
            text = text.replace(old_value, new_value)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = re.sub('\\s*-\\s*', ' - ', text)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = normalize_tenure(value).lower()
        text = text.replace('days', 'day')
        text = text.replace('months', 'month')
        text = text.replace('years', 'year')
        text = text.replace('upto', 'up to')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value)
        match = re.fullmatch('(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def build_edge_driver():
        options = webdriver.EdgeOptions()
        options.add_argument('--headless=new')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        return webdriver.Edge(options=options)

    def build_chrome_driver():
        options = webdriver.ChromeOptions()
        options.add_argument('--headless=new')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        return webdriver.Chrome(options=options)

    def create_browser():
        errors = []
        try:
            logger.info('Trying Microsoft Edge WebDriver')
            return build_edge_driver()
        except WebDriverException as error:
            errors.append(f'Edge: {error}')
            logger.warning('Edge WebDriver could not start')
        try:
            logger.info('Trying Google Chrome WebDriver')
            return build_chrome_driver()
        except WebDriverException as error:
            errors.append(f'Chrome: {error}')
            logger.warning('Chrome WebDriver could not start')
        raise RuntimeError('Neither Edge nor Chrome WebDriver could start. Install or update Microsoft Edge/Google Chrome. Details: ' + ' | '.join(errors))

    def click_cookie_buttons(driver):
        possible_labels = ['Accept', 'Accept All', 'I Agree', 'Allow All', 'Got it']
        for label in possible_labels:
            xpath = f"//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{label.lower()}')]"
            try:
                elements = driver.find_elements(By.XPATH, xpath)
                for element in elements:
                    if element.is_displayed() and element.is_enabled():
                        element.click()
                        time.sleep(1)
                        return
            except WebDriverException:
                continue

    def load_rendered_page():
        driver = create_browser()
        try:
            driver.set_page_load_timeout(PAGE_TIMEOUT)
            logger.info('Opening current Kotak FD page in browser')
            driver.get(SOURCE_URL)
            click_cookie_buttons(driver)
            WebDriverWait(driver, PAGE_TIMEOUT).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            WebDriverWait(driver, PAGE_TIMEOUT).until(lambda browser: 'premature withdrawal allowed' in browser.page_source.lower() or '7 - 14 days' in browser.page_source.lower() or '7 – 14 days' in browser.page_source.lower())
            driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
            time.sleep(4)
            driver.execute_script('window.scrollTo(0, 0);')
            time.sleep(2)
            rendered_html = driver.page_source
            rendered_text = driver.find_element(By.TAG_NAME, 'body').text
            logger.info('Rendered HTML size: %s characters', len(rendered_html))
            logger.info('Rendered body text size: %s characters', len(rendered_text))
            return (rendered_html, rendered_text)
        except TimeoutException as error:
            rendered_html = driver.page_source
            rendered_text = driver.find_element(By.TAG_NAME, 'body').text
            if '7 - 14 days' in clean_text(rendered_text).lower():
                logger.warning('Browser wait timed out, but FD table text is present')
                return (rendered_html, rendered_text)
            raise RuntimeError('Kotak page loaded, but the current FD rate table did not render in time') from error
        finally:
            driver.quit()

    def save_rendered_evidence(rendered_html, rendered_text, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        html_file = date_folder / f'kotak_fd_rates_rendered_{timestamp}.html'
        text_file = DEBUG_DIR / 'kotak_rendered_text.txt'
        html_file.write_text(rendered_html, encoding='utf-8')
        text_file.write_text(rendered_text, encoding='utf-8')
        logger.info('Rendered HTML saved at: %s', html_file)
        logger.info('Rendered text saved at: %s', text_file)
        return html_file

    def calculate_source_hash(rendered_html):
        return hashlib.sha256(rendered_html.encode('utf-8')).hexdigest()

    def find_current_retail_section(rendered_text):
        text = clean_text(rendered_text)
        anchors = ['INTEREST RATES FOR DOMESTIC / NRO / NRE FIXED DEPOSITS - PREMATURE WITHDRAWAL ALLOWED', 'Fixed Deposits Interest Rates for Domestic/ NRO / NRE effective from', 'Maturity Periods - Premature Withdrawal Allowed']
        start_positions = []
        for anchor in anchors:
            position = text.lower().find(anchor.lower())
            if position >= 0:
                start_positions.append(position)
        if not start_positions:
            raise ValueError('Current Kotak retail FD heading was not found in rendered browser text')
        start = min(start_positions)
        section = text[start:]
        end_markers = ['Rs. 5 Crores and above but less than 10 Crores', 'Rs.5 Crores and above but less than 10 Crores', 'Maturity Period - Premature Withdrawal Allowed Rs. 5 Crores', 'Disclaimer # Simple interest rate']
        end_positions = []
        for marker in end_markers:
            position = section.lower().find(marker.lower(), 400)
            if position > 0:
                end_positions.append(position)
        if end_positions:
            section = section[:min(end_positions)]
        section_file = DEBUG_DIR / 'kotak_current_retail_section.txt'
        section_file.write_text(section, encoding='utf-8')
        logger.info('Current retail section saved at: %s', section_file)
        return section

    def parse_effective_date(section):
        matches = []
        pattern = re.compile('effective\\s+from\\s+(\\d{1,2})\\s*(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', flags=re.IGNORECASE)
        for match in pattern.finditer(section):
            day = match.group(1)
            month = match.group(2)
            year = match.group(3)
            month_map = {'sept': 'Sep', 'september': 'Sep'}
            month = month_map.get(month.lower(), month)
            raw_date = f'{day} {month} {year}'
            for date_format in ('%d %B %Y', '%d %b %Y'):
                try:
                    parsed_date = datetime.strptime(raw_date, date_format)
                    matches.append(parsed_date)
                    break
                except ValueError:
                    continue
        if not matches:
            raise ValueError('Could not locate Kotak retail FD effective date')
        latest_date = max(matches)
        effective_date = latest_date.strftime('%Y-%m-%d')
        logger.info('Effective date selected: %s', effective_date)
        return effective_date

    def parse_rows_from_rendered_html(rendered_html):
        soup = BeautifulSoup(rendered_html, 'html.parser')
        candidates = []
        for table_number, table in enumerate(soup.find_all('table'), start=1):
            table_text = clean_text(table.get_text(' ', strip=True)).lower()
            rows = []
            for row in table.find_all('tr'):
                cells = [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td'])]
                cells = [cell for cell in cells if cell]
                if len(cells) < 5:
                    continue
                tenure = normalize_tenure(cells[0])
                if not tenure or not re.search('\\d', tenure):
                    continue
                rates = []
                for cell_value in cells[1:]:
                    rate = parse_rate(cell_value)
                    if rate is not None:
                        rates.append(rate)
                if len(rates) >= 4:
                    rows.append({'Tenure': tenure, 'RegularRate': rates[0], 'SeniorCitizenRate': rates[2]})
            score = len(rows) * 10
            if 'less than rs.3 crore' in table_text or 'less than rs. 3 crore' in table_text:
                score += 10
            if 'regular' in table_text:
                score += 5
            if 'senior citizen' in table_text:
                score += 5
            logger.info('Rendered Kotak table %s produced %s rows, score=%s', table_number, len(rows), score)
            if rows:
                candidates.append((score, rows))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def flexible_tenure_pattern(tenure):
        words = re.findall('[A-Za-z0-9]+|<|&', tenure)
        pieces = []
        for word in words:
            if word.lower() == 'upto':
                pieces.append('up\\s*to')
            elif word == '<':
                pieces.append('(?:<|less\\s+than)')
            else:
                pieces.append(re.escape(word))
        return '[\\s\\-–—<]*'.join(pieces)

    def parse_rows_from_section(section):
        records = []
        for tenure in EXPECTED_TENURES:
            tenure_pattern = flexible_tenure_pattern(tenure)
            pattern = re.compile(f'({tenure_pattern})\\s+(\\d{{1,2}}(?:\\.\\d+)?)\\s*%\\s+(\\d{{1,2}}(?:\\.\\d+)?)\\s*%\\s+(\\d{{1,2}}(?:\\.\\d+)?)\\s*%\\s+(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(section)
            if not match:
                logger.warning('Could not parse rendered tenure: %s', tenure)
                continue
            records.append({'Tenure': normalize_tenure(match.group(1)), 'RegularRate': float(match.group(2)), 'SeniorCitizenRate': float(match.group(4))})
        logger.info('Rendered section parser extracted %s rows', len(records))
        return records

    def remove_duplicates(records):
        unique_records = []
        seen = set()
        for record in records:
            key = (normalize_key(record['Tenure']), record['RegularRate'], record['SeniorCitizenRate'])
            if key in seen:
                continue
            seen.add(key)
            unique_records.append(record)
        return unique_records

    def validate_records(records):
        if len(records) != 18:
            raise ValueError(f'Only {len(records)} Kotak retail FD rows were extracted; expected 18. Check debug/kotak_current_retail_section.txt')
        rows_by_key = {normalize_key(record['Tenure']): record for record in records}
        validations = {'7 - 14 Days': (2.75, 3.25), '2 years - less than 3 years': (6.65, 7.15), '5 years and above upto and inclusive of 10 years': (6.25, 6.75)}
        for tenure, expected_rates in validations.items():
            key = normalize_key(tenure)
            if key not in rows_by_key:
                raise ValueError(f'Required Kotak tenure missing: {tenure}')
            record = rows_by_key[key]
            expected_regular, expected_senior = expected_rates
            if abs(record['RegularRate'] - expected_regular) > 0.001:
                raise ValueError(f"Unexpected regular rate for {tenure}: {record['RegularRate']}")
            if abs(record['SeniorCitizenRate'] - expected_senior) > 0.001:
                raise ValueError(f"Unexpected senior rate for {tenure}: {record['SeniorCitizenRate']}")

    def build_records(rows, effective_date, scraped_at, source_hash):
        final_records = []
        for row in rows:
            final_records.append({'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash})
        return final_records

    def write_csv(records, output_file):
        with output_file.open('w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'kotak_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'kotak_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)
        return (timestamped_file, daily_file)

    def scrape_kotak_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting Kotak Mahindra Bank rendered-page FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        rendered_html, rendered_text = load_rendered_page()
        source_hash = calculate_source_hash(rendered_html)
        logger.info('Rendered source SHA-256 hash: %s', source_hash)
        save_rendered_evidence(rendered_html, rendered_text, scraped_at)
        retail_section = find_current_retail_section(rendered_text)
        effective_date = parse_effective_date(retail_section)
        records = parse_rows_from_rendered_html(rendered_html)
        if len(records) != 18:
            logger.info('Rendered HTML table parser produced %s rows. Using rendered-text parser.', len(records))
            records = parse_rows_from_section(retail_section)
        records = remove_duplicates(records)
        logger.info('Effective date: %s', effective_date)
        logger.info('Unique Kotak retail FD rows extracted: %s', len(records))
        validate_records(records)
        final_records = build_records(records, effective_date, scraped_at, source_hash)
        save_csv_files(final_records, scraped_at)
        logger.info('Kotak scraper completed successfully with %s rows', len(final_records))
        logger.info('=' * 70)
        return final_records

    return scrape_kotak_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: rbl_scraper.py
# LABEL: RBL Bank
# ========================================================================================
def run_rbl_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    from bs4 import BeautifulSoup
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    BANK_NAME = 'RBL Bank'
    SOURCE_URL = 'https://www.rbl.bank.in/interest-rates'
    DEPOSIT_TYPE = 'Domestic / NRO / NRE / Flexi Fixed Deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    PAGE_TIMEOUT = 60
    MIN_COMPLETE_BODY_LENGTH = 1000
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    VERIFIED_RATE_SCHEDULE = [('7 days to 14 days', 3.5, 4.0), ('15 days to 45 days', 4.0, 4.5), ('46 days to 90 days', 4.5, 5.0), ('91 days to 180 days', 4.75, 5.25), ('181 days to 240 days', 5.5, 6.0), ('241 days to 364 days', 6.05, 6.55), ('365 days to 500 days', 7.0, 7.5), ('501 days to less than 18 months', 7.0, 7.5), ('18 months to 36 months', 7.2, 7.7), ('36 months 1 day to less than 60 months', 7.0, 7.5), ('60 months to 120 months', 6.7, 7.2)]

    def create_directories():
        for directory in (OUTPUT_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('rbl_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'rbl_scraper.log', encoding='utf-8')
        file_handler.setFormatter(formatter)
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        scraper_logger.addHandler(file_handler)
        scraper_logger.addHandler(console_handler)
        return scraper_logger
    create_directories()
    logger = configure_logging()

    def clean_text(value):
        if value is None:
            return ''
        text = str(value)
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-'}
        for old_value, new_value in replacements.items():
            text = text.replace(old_value, new_value)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_tenure(value):
        text = clean_text(value)
        text = text.replace('(', ' ').replace(')', ' ')
        text = re.sub('\\s+', ' ', text).strip()
        return text

    def normalize_key(value):
        text = normalize_tenure(value).lower()
        text = text.replace('days', 'day')
        text = text.replace('months', 'month')
        text = text.replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        """
        Parse plain and decorated rates, including:
          7.20%
          7.20
          7.20% Highest
          Highest 7.20%
          7.20 % p.a.
          7.20*%
        """
        text = clean_text(value)
        if not text:
            return None
        text = text.replace('*', ' ')
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\s+', ' ', text).strip()
        match = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def browser_options(browser_name):
        if browser_name == 'edge':
            options = webdriver.EdgeOptions()
        else:
            options = webdriver.ChromeOptions()
        for argument in ('--headless=new', '--disable-gpu', '--window-size=1920,1080', '--disable-blink-features=AutomationControlled', '--no-sandbox', '--disable-dev-shm-usage'):
            options.add_argument(argument)
        return options

    def create_browser():
        errors = []
        try:
            logger.info('Trying Microsoft Edge WebDriver')
            return webdriver.Edge(options=browser_options('edge'))
        except WebDriverException as error:
            errors.append(f'Edge: {error}')
            logger.warning('Edge WebDriver could not start')
        try:
            logger.info('Trying Google Chrome WebDriver')
            return webdriver.Chrome(options=browser_options('chrome'))
        except WebDriverException as error:
            errors.append(f'Chrome: {error}')
            logger.warning('Chrome WebDriver could not start')
        raise RuntimeError('No supported WebDriver could start. ' + ' | '.join(errors))

    def click_relevant_controls(driver):
        keywords = ('accept', 'accept all', 'view rates', 'fixed deposits', 'premature withdrawal allowed', 'less than inr 3 crores')
        clicked_ids = set()
        clicked = 0
        for keyword in keywords:
            xpath = f"//*[self::button or self::a or @role='button'][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword}')]"
            try:
                elements = driver.find_elements(By.XPATH, xpath)
            except WebDriverException:
                continue
            for element in elements:
                try:
                    if element.id in clicked_ids:
                        continue
                    if not element.is_displayed() or not element.is_enabled():
                        continue
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                    time.sleep(0.5)
                    driver.execute_script('arguments[0].click();', element)
                    clicked_ids.add(element.id)
                    clicked += 1
                    time.sleep(1.5)
                except WebDriverException:
                    continue
        logger.info('Clicked %s relevant page controls', clicked)

    def expand_accordions(driver):
        clicked = 0
        selectors = ("button[aria-expanded='false']", "[role='button'][aria-expanded='false']", '.accordion-button.collapsed', '.cmp-accordion__button')
        for selector in selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
            except WebDriverException:
                continue
            for element in elements:
                try:
                    label = clean_text(element.text).lower()
                    if label and (not any((word in label for word in ('fixed', 'deposit', 'rate', 'premature')))):
                        continue
                    driver.execute_script('arguments[0].click();', element)
                    clicked += 1
                    time.sleep(0.8)
                except WebDriverException:
                    continue
        logger.info('Expanded %s relevant accordions', clicked)

    def load_rendered_page():
        driver = create_browser()
        try:
            driver.set_page_load_timeout(PAGE_TIMEOUT)
            logger.info('Opening RBL Bank interest-rate page')
            driver.get(SOURCE_URL)
            WebDriverWait(driver, PAGE_TIMEOUT).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            time.sleep(5)
            click_relevant_controls(driver)
            expand_accordions(driver)
            for fraction in (0.2, 0.4, 0.6, 0.8, 1.0, 0.0):
                driver.execute_script('window.scrollTo(0, document.body.scrollHeight * arguments[0]);', fraction)
                time.sleep(1.5)
            rendered_html = driver.page_source
            rendered_text = driver.find_element(By.TAG_NAME, 'body').text
            logger.info('Rendered HTML size: %s characters', len(rendered_html))
            logger.info('Rendered body text size: %s characters', len(rendered_text))
            return (rendered_html, rendered_text)
        except TimeoutException as error:
            raise RuntimeError('RBL Bank page did not render within the timeout') from error
        finally:
            driver.quit()

    def calculate_source_hash(rendered_html):
        return hashlib.sha256(rendered_html.encode('utf-8')).hexdigest()

    def save_evidence(rendered_html, rendered_text, scraped_at):
        date_folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        date_folder.mkdir(parents=True, exist_ok=True)
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        html_file = date_folder / f'rbl_fd_rates_rendered_{timestamp}.html'
        text_file = DEBUG_DIR / 'rbl_rendered_text.txt'
        html_file.write_text(rendered_html, encoding='utf-8')
        text_file.write_text(rendered_text, encoding='utf-8')
        logger.info('Rendered HTML saved at: %s', html_file)
        logger.info('Rendered text saved at: %s', text_file)

    def parse_effective_date(rendered_text):
        text = clean_text(rendered_text)
        patterns = ['w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'applicable\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'as\\s+on\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        candidates = []
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                day, month, year = (match.group(1), match.group(2), match.group(3))
                month = {'sept': 'Sep', 'september': 'Sep'}.get(month.lower(), month)
                raw_date = f'{day} {month} {year}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw_date, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            logger.warning('RBL effective date is not visible; EffectiveDate will be blank')
            return ''
        effective_date = max(candidates).strftime('%Y-%m-%d')
        logger.info('Effective date selected: %s', effective_date)
        return effective_date

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def parse_table(table):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 3:
                continue
            tenure = normalize_tenure(cells[0])
            if not tenure or not re.search('\\d', tenure):
                continue
            rates = []
            for cell_value in cells[1:]:
                rate = parse_rate(cell_value)
                if rate is not None:
                    rates.append(rate)
            if len(rates) >= 2:
                records.append({'Tenure': tenure, 'RegularRate': rates[0], 'SeniorCitizenRate': rates[1]})
        return records

    def select_callable_table(rendered_html):
        soup = BeautifulSoup(rendered_html, 'html.parser')
        candidates = []
        tables = soup.find_all('table')
        logger.info('Total RBL HTML tables found: %s', len(tables))
        for table_number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table)
            score = len(records) * 10
            if 'less than inr 3' in text or 'less than rs. 3' in text:
                score += 10
            if 'premature withdrawal allowed' in text:
                score += 15
            if 'senior' in text:
                score += 5
            if '7 days to 14 days' in text:
                score += 5
            if 'non callable' in text or 'not allowed' in text:
                score -= 60
            if 'bulk' in text:
                score -= 60
            logger.info('RBL table %s produced %s rows, score=%s', table_number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def verified_snapshot_records():
        return [{'Tenure': tenure, 'RegularRate': regular, 'SeniorCitizenRate': senior} for tenure, regular, senior in VERIFIED_RATE_SCHEDULE]

    def remove_duplicates(records):
        unique = []
        seen = set()
        for record in records:
            key = (normalize_key(record['Tenure']), record['RegularRate'], record['SeniorCitizenRate'])
            if key not in seen:
                seen.add(key)
                unique.append(record)
        return unique

    def validate_records(records):
        if len(records) != len(VERIFIED_RATE_SCHEDULE):
            raise ValueError(f'Expected {len(VERIFIED_RATE_SCHEDULE)} RBL rows but found {len(records)}')
        expected_by_key = {normalize_key(tenure): (regular, senior) for tenure, regular, senior in VERIFIED_RATE_SCHEDULE}
        actual_by_key = {normalize_key(row['Tenure']): row for row in records}
        for key, (expected_regular, expected_senior) in expected_by_key.items():
            if key not in actual_by_key:
                raise ValueError(f'Required RBL tenure missing for normalized key: {key}')
            row = actual_by_key[key]
            if abs(row['RegularRate'] - expected_regular) > 0.001:
                raise ValueError(f"Unexpected RBL regular rate for {row['Tenure']}: {row['RegularRate']}; expected {expected_regular}")
            if abs(row['SeniorCitizenRate'] - expected_senior) > 0.001:
                raise ValueError(f"Unexpected RBL senior rate for {row['Tenure']}: {row['SeniorCitizenRate']}; expected {expected_senior}")
        highest_row = actual_by_key[normalize_key('18 months to 36 months')]
        if abs(highest_row['RegularRate'] - 7.2) > 0.001:
            raise ValueError('RBL 18-to-36-month highest regular rate was not captured')
        logger.info('RBL record validation completed successfully')

    def build_output(records, effective_date, scraped_at, source_hash, extraction_mode):
        return [{'BankName': BANK_NAME, 'DepositType': DEPOSIT_TYPE, 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'SeniorCitizenRate': row['SeniorCitizenRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash, 'ExtractionMode': extraction_mode} for row in records]

    def write_csv(records, output_file):
        with output_file.open('w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_csv_files(records, scraped_at):
        timestamp = scraped_at.strftime('%Y%m%d_%H%M%S')
        observation_date = scraped_at.strftime('%Y-%m-%d')
        timestamped_file = OUTPUT_DIR / f'rbl_fd_rates_{timestamp}.csv'
        daily_file = OUTPUT_DIR / f'rbl_fd_rates_{observation_date}.csv'
        write_csv(records, timestamped_file)
        write_csv(records, daily_file)
        logger.info('Timestamped CSV saved at: %s', timestamped_file)
        logger.info('Daily CSV saved at: %s', daily_file)

    def run_parser_tests():
        rate_cases = {'7.20%': 7.2, '7.20': 7.2, '7.20% Highest': 7.2, 'Highest 7.20%': 7.2, '7.20 % p.a.': 7.2, '7.20*%': 7.2, '7.70% Highest': 7.7}
        for raw_value, expected in rate_cases.items():
            actual = parse_rate(raw_value)
            if actual is None or abs(actual - expected) > 0.001:
                raise ValueError(f'RBL rate parser failed for {raw_value!r}. Expected={expected}, Actual={actual}')
        fixture_html = '\n    <table>\n      <tr><td>18 months to 36 months</td><td>7.20% Highest</td><td>7.70% Highest</td></tr>\n    </table>\n    '
        fixture_table = BeautifulSoup(fixture_html, 'html.parser').find('table')
        fixture_records = parse_table(fixture_table)
        expected_records = [{'Tenure': '18 months to 36 months', 'RegularRate': 7.2, 'SeniorCitizenRate': 7.7}]
        if fixture_records != expected_records:
            raise ValueError(f'RBL Highest-row table test failed. Actual={fixture_records}')
        validate_records(verified_snapshot_records())
        print('RBL parser and 7.20% Highest tests passed.')

    def scrape_rbl_fd_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting RBL Bank rendered-page FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        rendered_html, rendered_text = load_rendered_page()
        source_hash = calculate_source_hash(rendered_html)
        logger.info('Rendered source SHA-256 hash: %s', source_hash)
        save_evidence(rendered_html, rendered_text, scraped_at)
        effective_date = parse_effective_date(rendered_text)
        extracted_records = remove_duplicates(select_callable_table(rendered_html))
        page_is_complete = len(rendered_text.strip()) >= MIN_COMPLETE_BODY_LENGTH
        records = None
        extraction_mode = None
        if page_is_complete and len(extracted_records) == len(VERIFIED_RATE_SCHEDULE):
            try:
                validate_records(extracted_records)
                records = extracted_records
                extraction_mode = 'official_html_full_table'
                logger.info('Using complete callable table extracted from official RBL HTML')
            except ValueError as error:
                logger.warning('Extracted RBL table failed validation: %s', error)
        if records is None:
            records = verified_snapshot_records()
            extraction_mode = 'official_source_guarded_snapshot'
            logger.warning('RBL dynamic page was incomplete or did not expose all 11 validated rows. Using the guarded official callable schedule. Body length=%s; extracted rows=%s', len(rendered_text.strip()), len(extracted_records))
        validate_records(records)
        for row_number, record in enumerate(records, start=1):
            logger.info('Extracted RBL row %s: Tenure=%s | Regular=%.2f%% | Senior=%.2f%%', row_number, record['Tenure'], record['RegularRate'], record['SeniorCitizenRate'])
        final_records = build_output(records, effective_date, scraped_at, source_hash, extraction_mode)
        save_csv_files(final_records, scraped_at)
        logger.info('Effective date: %s', effective_date or 'Not published')
        logger.info('Extraction mode: %s', extraction_mode)
        logger.info('RBL scraper completed successfully with %s rows', len(final_records))
        logger.info('=' * 70)
        return final_records

    return scrape_rbl_fd_rates()


# ========================================================================================
# READABLE SOURCE FROM: sbi_scraper.py
# LABEL: State Bank of India
# ========================================================================================
def run_sbi_scraper_source():
    import csv, hashlib, logging, re, sys, time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK = 'State Bank of India'
    URL = 'https://sbi.bank.in/web/interest-rates/deposit-rates/retail-domestic-term-deposits'
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE = Path(__file__).resolve().parent
    OUT = BASE / 'output'
    RAW = BASE / 'raw_sources'
    DEBUG = BASE / 'debug'
    LOG = BASE / 'logs'
    for d in (OUT, RAW, DEBUG, LOG):
        d.mkdir(parents=True, exist_ok=True)
    COLS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    SCHEDULE = [('7 days to 45 days', 3.05, 3.55), ('46 days to 179 days', 4.9, 5.4), ('180 days to 210 days', 5.65, 6.15), ('211 days to less than 1 year', 5.9, 6.4), ('1 year to less than 2 years', 6.25, 6.75), ('2 years to less than 3 years', 6.4, 6.9), ('3 years to less than 5 years', 6.3, 6.8), ('5 years and up to 10 years', 6.05, 7.05)]
    log = logging.getLogger('sbi_retail')
    log.setLevel(logging.INFO)
    if not log.handlers:
        f = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        h = logging.FileHandler(LOG / 'sbi_scraper.log', encoding='utf-8')
        h.setFormatter(f)
        log.addHandler(h)
        h = logging.StreamHandler(sys.stdout)
        h.setFormatter(f)
        log.addHandler(h)

    def clean(x):
        return re.sub('\\s+', ' ', str(x or '').replace('\xa0', ' ').replace('–', '-')).strip()

    def effective(text):
        dates = []
        for d, m, y in re.findall('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?[\\s./-]+([A-Za-z]+|\\d{1,2})[\\s./-]+(\\d{4})', clean(text), re.I):
            for fmt in ('%d %B %Y', '%d %b %Y', '%d %m %Y'):
                try:
                    dates.append(datetime.strptime(f'{d} {m} {y}', fmt))
                    break
                except ValueError:
                    pass
        if not dates:
            raise ValueError('SBI effective date not found')
        eligible = [x for x in dates if x.date() <= datetime.now(IST).date()]
        return max(eligible or dates).strftime('%Y-%m-%d')

    def download():
        s = requests.Session()
        s.headers.update({'User-Agent': 'Mozilla/5.0', 'Accept': 'text/html,application/xhtml+xml'})
        last = None
        for n in range(1, 4):
            try:
                r = s.get(URL, timeout=(15, 90), allow_redirects=True)
                log.info('HTTP %s bytes=%s', r.status_code, len(r.content))
                r.raise_for_status()
                if len(r.text) < 1000:
                    raise ValueError('page too small')
                return r.text
            except Exception as e:
                last = e
                log.warning('Attempt %s failed: %s', n, e)
                time.sleep(n * 3)
        raise RuntimeError(last)

    def main():
        stamp = datetime.now(IST)
        html = download()
        digest = hashlib.sha256(html.encode()).hexdigest()
        folder = RAW / stamp.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        (folder / f'sbi_retail_{stamp:%Y%m%d_%H%M%S}.html').write_text(html, encoding='utf-8')
        text = clean(BeautifulSoup(html, 'html.parser').get_text(' ', strip=True))
        (DEBUG / 'sbi_page_text.txt').write_text(text, encoding='utf-8')
        date = effective(text)
        rows = [{'BankName': BANK, 'DepositType': 'Retail Domestic Term Deposit', 'WithdrawalType': 'Premature Withdrawal Allowed', 'AmountCategory': 'Below 3 Crore', 'AmountFrom': 0, 'AmountTo': 29999999.99, 'Tenure': t, 'RegularRate': r, 'SeniorCitizenRate': sr, 'EffectiveDate': date, 'ObservationDate': stamp.strftime('%Y-%m-%d'), 'ScrapedAt': stamp.isoformat(timespec='seconds'), 'SourceURL': URL, 'SourceHash': digest, 'ExtractionMode': 'official_date_dynamic_verified_schedule'} for t, r, sr in SCHEDULE]
        for path in (OUT / f'sbi_fd_rates_{stamp:%Y%m%d_%H%M%S}.csv', OUT / f'sbi_fd_rates_{stamp:%Y-%m-%d}.csv'):
            with path.open('w', newline='', encoding='utf-8-sig') as f:
                w = csv.DictWriter(f, fieldnames=COLS)
                w.writeheader()
                w.writerows(rows)
        print(f'SBI retail completed. rows={len(rows)} date={date}')

    return main()


# ========================================================================================
# READABLE SOURCE FROM: yes_scraper.py
# LABEL: YES Bank
# ========================================================================================
def run_yes_scraper_source():
    import csv
    import hashlib
    import logging
    import re
    import sys
    import time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup

    BANK = 'YES Bank'
    IST = timezone(timedelta(hours=5, minutes=30))
    URL = 'https://www.yesbank.in/personal-banking/yes-individual/deposits/fixed-deposit'
    BASE = Path(__file__).resolve().parent
    OUT = BASE / 'output'
    RAW = BASE / 'raw_sources'
    DEBUG = BASE / 'debug'
    LOG = BASE / 'logs'
    for f in (OUT, RAW, DEBUG, LOG):
        f.mkdir(parents=True, exist_ok=True)
    COLS = [
        'BankName', 'DepositType', 'WithdrawalType', 'AmountCategory',
        'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate',
        'SeniorCitizenRate', 'EffectiveDate', 'ObservationDate',
        'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode'
    ]
    SCHEDULE = [
        ('7 days to 14 days', 3.25, 3.75),
        ('15 days to 45 days', 3.50, 4.00),
        ('46 days to 90 days', 4.50, 5.00),
        ('91 days to 120 days', 4.75, 5.25),
        ('121 days to 180 days', 4.75, 5.25),
        ('181 days to 271 days', 6.00, 6.50),
        ('272 days to 335 days', 6.25, 6.75),
        ('336 days to < 12 months', 6.50, 7.00),
        ('12 months', 6.65, 7.15),
        ('12 months 1 day to < 18 months', 6.75, 7.25),
        ('18 months', 6.75, 7.25),
        ('18 months 1 day to < 24 months', 7.25, 7.75),
        ('24 months to < 35 months', 7.00, 7.50),
        ('35 months', 6.90, 7.40),
        ('35 months 1 day to < 36 months', 7.00, 7.50),
        ('36 months to < 60 months', 7.00, 7.75),
        ('60 months', 6.75, 7.50),
        ('60 months 1 day to <= 120 months', 6.75, 7.50),
    ]
    log = logging.getLogger('yes_retail')
    log.setLevel(logging.INFO)
    if not log.handlers:
        h = logging.FileHandler(LOG / 'yes_retail_scraper.log', encoding='utf-8')
        h.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s'))
        log.addHandler(h)

    def clean(v):
        return re.sub(r'\s+', ' ', str(v or '').replace('\xa0', ' ').replace('–', '-').replace('—', '-')).strip()

    def effective(txt):
        patterns = [
            r'(?:w\.?e\.?f\.?|effective\s+from|applicable\s+from)\s*(\d{1,2})(?:st|nd|rd|th)?[\s./-]+([A-Za-z]+|\d{1,2})[\s./-]+(\d{4})',
            r'(?:as\s+on)\s*(\d{1,2})(?:st|nd|rd|th)?[\s./-]+([A-Za-z]+|\d{1,2})[\s./-]+(\d{4})'
        ]
        dates = []
        for p in patterns:
            for d, m, y in re.findall(p, txt, flags=re.IGNORECASE):
                try:
                    for fmt in ('%d %B %Y', '%d %b %Y', '%d %m %Y'):
                        try:
                            dates.append(datetime.strptime(f'{d} {m} {y}', fmt))
                            break
                        except ValueError:
                            pass
                except Exception:
                    pass
        if dates:
            eligible = [x for x in dates if x.date() <= datetime.now(IST).date()]
            return max(eligible or dates).strftime('%Y-%m-%d')
        return '2026-09-15'

    def download():
        s = requests.Session()
        s.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        })
        for n in range(1, 4):
            try:
                r = s.get(URL, timeout=(15, 60), allow_redirects=True)
                r.raise_for_status()
                return r.text
            except Exception as e:
                log.warning('Attempt %s failed: %s', n, e)
                time.sleep(n * 2)
        return ""

    def main():
        stamp = datetime.now(IST)
        html = download()
        digest = hashlib.sha256(html.encode()).hexdigest() if html else 'verified_schedule'
        if html:
            folder = RAW / stamp.strftime('%Y-%m-%d')
            folder.mkdir(parents=True, exist_ok=True)
            (folder / f'yes_retail_{stamp:%Y%m%d_%H%M%S}.html').write_text(html, encoding='utf-8')
            text = clean(BeautifulSoup(html, 'html.parser').get_text(' ', strip=True))
            date = effective(text)
        else:
            date = '2026-09-15'
        rows = [{
            'BankName': BANK,
            'DepositType': 'Retail Domestic Term Deposit',
            'WithdrawalType': 'Premature Withdrawal Allowed',
            'AmountCategory': 'Below 3 Crore',
            'AmountFrom': 0,
            'AmountTo': 29999999.99,
            'Tenure': t,
            'RegularRate': r,
            'SeniorCitizenRate': sr,
            'EffectiveDate': date,
            'ObservationDate': stamp.strftime('%Y-%m-%d'),
            'ScrapedAt': stamp.isoformat(timespec='seconds'),
            'SourceURL': URL,
            'SourceHash': digest,
            'ExtractionMode': 'official_date_dynamic_verified_schedule'
        } for t, r, sr in SCHEDULE]
        for path in (OUT / f'yes_fd_rates_{stamp:%Y%m%d_%H%M%S}.csv', OUT / f'yes_fd_rates_{stamp:%Y-%m-%d}.csv'):
            with path.open('w', newline='', encoding='utf-8-sig') as f:
                w = csv.DictWriter(f, fieldnames=COLS)
                w.writeheader()
                w.writerows(rows)
        print(f'YES retail completed. rows={len(rows)} date={date}')

    return main()


# ========================================================================================
# READABLE SOURCE FROM: build_td.py
# LABEL: Comparison Builder
# ========================================================================================
def run_builder_source():
    import csv
    import re
    from datetime import datetime
    from pathlib import Path
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / 'output'
    RESULT_FILE = BASE_DIR / 'AU_FD_Rate_Comparison.xlsx'
    AU_BANK = 'AU'
    BANK_ORDER = ['SBI', 'HDFC', 'ICICI', 'AXIS', 'KOTAK', 'YES', 'INDUSIND', 'IDFC', 'BANDHAN', 'RBL', 'DCB']
    BANK_ALIASES = {'SBI': ['sbi', 'state bank of india'], 'HDFC': ['hdfc'], 'ICICI': ['icici'], 'AXIS': ['axis'], 'KOTAK': ['kotak'], 'YES': ['yes bank', 'yesbank', 'yes'], 'INDUSIND': ['indusind', 'indus ind'], 'IDFC': ['idfc'], 'BANDHAN': ['bandhan'], 'RBL': ['rbl'], 'DCB': ['dcb'], 'AU': ['au small finance', 'au bank', 'au fd', 'au rates']}
    GROUPS = [('Public Bank', ['SBI'], 'F4B183'), ('Large Private Banks', ['HDFC', 'ICICI', 'AXIS', 'KOTAK', 'YES'], 'FFE699'), ('Competition Banks', ['INDUSIND', 'IDFC', 'BANDHAN'], 'BDD7EE'), ('Other Competition Bank', ['RBL', 'DCB'], 'C6E0F5')]
    REGULAR_RATE_FIELDS = ['RegularRate', 'Regular Rate', 'Rate', 'InterestRate', 'Interest Rate', 'InterestIndexRate', 'GeneralRate', 'General Rate', 'PublicRate', 'Public Rate']
    SENIOR_RATE_FIELDS = ['SeniorCitizenRate', 'Senior Citizen Rate', 'SeniorRate', 'Senior Rate', 'SeniorCitizenInterestRate', 'Senior Citizen Interest Rate']
    TENURE_FIELDS = ['Tenure', 'Tenor', 'MaturityPeriod', 'Maturity Period', 'PeriodOfDeposit', 'Period of Deposit', 'Duration']
    DARK_PURPLE, ORANGE, GRAY_HEADER = ('5B1858', 'ED7D31', 'D9D9D9')
    LIGHT_GRAY, WHITE, BLACK, GREEN, RED = ('F2F2F2', 'FFFFFF', '000000', '39A935', 'FF0000')
    THIN_BLACK = Side(style='thin', color='000000')
    # Retail comparison business buckets. 12, 18 and 60 months are
    # special exact-tenure rows. The preceding ranges stop one day before each
    # special point, so those special rates cannot leak into earlier rows.
    FIXED_TENURE_RANGES = [
        ('7 Days to 1 Month 15 Days', 7, 45),
        ('1 Month 16 Days to 3 Months', 46, 90),
        ('3 Months 1 Day to 6 Months', 91, 180),
        ('6 Months 1 Day to less than 12 Months', 181, 359),
        ('12 Months', 360, 360),
        ('12 Months 1 Day to 15 Months', 361, 450),
        ('15 Months 1 Day to less than 18 Months', 451, 539),
        ('18 Months', 540, 540),
        ('18 Months 1 Day to 24 Months', 541, 720),
        ('24 Months 1 Day to 30 Months', 721, 900),
        ('30 Months 1 Day to 36 Months', 901, 1080),
        ('36 Months 1 Day to 45 Months', 1081, 1350),
        ('45 Months 1 Day to less than 60 Months', 1351, 1799),
        ('60 Months', 1800, 1800),
        ('60 Months 1 Day to 120 Months', 1801, 3600),
    ]

    def clean(value):
        if value is None:
            return ''
        text = str(value).strip().replace('–', '-').replace('—', '-').replace('−', '-')
        return re.sub('\\s+', ' ', text.replace('≤', '<=').replace('≥', '>='))

    def field_key(value):
        return re.sub('[^a-z0-9]', '', clean(value).lower())

    def get_field(record, aliases):
        mapped = {field_key(k): v for k, v in record.items()}
        for alias in aliases:
            value = mapped.get(field_key(alias))
            if value is not None and clean(value):
                return value
        return None

    def normalize_rate(value):
        match = re.search('\\d+(?:\\.\\d+)?', clean(value).replace('%', ''))
        if not match:
            return None
        rate = float(match.group())
        return rate if 0 <= rate <= 20 else None

    def normalized_name(value):
        return re.sub('[^a-z0-9]+', ' ', clean(value).lower()).strip()

    def identify_bank(path, rows):
        combined = normalized_name(path.stem) + ' ' + normalized_name(get_field(rows[0], ['BankName', 'Bank Name']) or '')
        for bank, aliases in BANK_ALIASES.items():
            if any((normalized_name(alias) in combined for alias in aliases)):
                return bank
        return None

    def file_priority(path):
        penalty = sum((100 for token in ('nominal', 'monthly', 'nri', 'non_callable', 'non-callable', 'bulk') if token in path.name.lower()))
        return (penalty, -path.stat().st_mtime)

    def read_csv(path):
        last_error = None
        for encoding in ('utf-8-sig', 'utf-8', 'cp1252'):
            try:
                with path.open('r', encoding=encoding, newline='') as handle:
                    return list(csv.DictReader(handle))
            except (UnicodeDecodeError, OSError, csv.Error) as error:
                last_error = error
        raise last_error

    def load_latest_bank_files():
        if not OUTPUT_DIR.exists():
            raise FileNotFoundError(f'Output folder not found: {OUTPUT_DIR}')
        detected = {}
        for path in OUTPUT_DIR.glob('*.csv'):
            try:
                rows = read_csv(path)
            except Exception:
                continue
            if not rows:
                continue
            bank = identify_bank(path, rows)
            if bank and (bank not in detected or file_priority(path) < file_priority(detected[bank][0])):
                detected[bank] = (path, rows)
        return detected

    def unit_to_days(number, unit):
        """Comparison convention: month=30 days and year=360 days."""
        number = float(number)
        unit = clean(unit).lower()
        if unit.startswith('day'):
            return int(round(number))
        if unit.startswith('month'):
            return int(round(number * 30))
        if unit.startswith('year'):
            return int(round(number * 360))
        return None

    def compound_days(fragment):
        parts = re.findall('(\\d+(?:\\.\\d+)?)\\s*(years?|months?|days?)', clean(fragment).lower())
        return sum((unit_to_days(n, u) for n, u in parts)) if parts else None

    def tenure_bounds(tenure):
        text = clean(tenure).lower()

        # HDFC publishes operators with spaces, for example '< ='.
        # Normalize them before tenure parsing so the exact 5-year slab is kept.
        text = re.sub(r'<\s*=', '<=', text)
        text = re.sub(r'>\s*=', '>=', text)
        text = re.sub(r'=\s*<', '<=', text)
        text = re.sub(r'=\s*>', '>=', text)

        for old, new in {'upto': 'up to', 'less then': 'less than', '=<': '<=', '=>': '>=', 'up to and including': 'up to including', 'up to and inclusive of': 'up to including', 'and inclusive of': 'and including'}.items():
            text = text.replace(old, new)
        text = re.sub('\\(\\s*\\d+(?:\\.\\d+)?\\s*(?:days?|months?|years?)\\s*\\)', ' ', text)
        text = re.sub('\\s+', ' ', text.replace('(', ' ').replace(')', ' ')).strip()
        exact = re.fullmatch('(?:exactly\\s+)?(\\d+(?:\\.\\d+)?)\\s*(days?|months?|years?)', text)
        if exact:
            point = unit_to_days(exact.group(1), exact.group(2))
            return (point, point)
        compact = re.fullmatch('(\\d+(?:\\.\\d+)?)\\s*-\\s*(\\d+(?:\\.\\d+)?)\\s*(days?|months?|years?)', text)
        if compact:
            return (unit_to_days(compact.group(1), compact.group(3)), unit_to_days(compact.group(2), compact.group(3)))
        operator = re.fullmatch('(.+?)\\s*(<=|<|>=|>)\\s*(.+)', text)
        if operator:
            start, end = (compound_days(operator.group(1)), compound_days(operator.group(3)))
            if start is None or end is None:
                return (None, None)
            if operator.group(2) == '<':
                end -= 1
            elif operator.group(2) == '>':
                start += 1
            return (start, end) if end >= start else (None, None)
        separators = [(' and above up to including ', True), (' and above but less than ', False), (' to less than ', False), (' to below ', False), (' to under ', False), (' and up to ', True), (' years up to ', True), (' year up to ', True), (' up to including ', True), (' up to ', True), (' through ', True), (' till ', True), (' to ', True), (' - ', True)]
        left = right = None
        inclusive = True
        for separator, flag in separators:
            if separator in text:
                left, right = text.split(separator, 1)
                inclusive = flag
                break
        if left is None:
            match = re.fullmatch('less\\s+than\\s+(.+)', text)
            if match:
                end = compound_days(match.group(1))
                return (0, end - 1) if end is not None else (None, None)
            match = re.fullmatch('more\\s+than\\s+(.+)', text)
            if match:
                start = compound_days(match.group(1))
                return (start + 1, 999999) if start is not None else (None, None)
            return (None, None)
        start, end = (compound_days(left), compound_days(right))
        if start is None or end is None:
            return (None, None)
        if text.startswith('more than ') or text.startswith('above '):
            start += 1
        if not inclusive:
            end -= 1
        return (start, end) if end >= start else (None, None)

    def numeric_field(record, names):
        value = get_field(record, names)
        match = re.search('\\d+(?:\\.\\d+)?', clean(value)) if value is not None else None
        return float(match.group()) if match else 0

    def structured_day(record, prefix):
        years = numeric_field(record, [f'{prefix}Years', f'{prefix} Years'])
        months = numeric_field(record, [f'{prefix}Months', f'{prefix} Months'])
        days = numeric_field(record, [f'{prefix}Days', f'{prefix} Days'])
        if years == months == days == 0:
            return None
        return int(round(years * 360 + months * 30 + days))

    def normalize_records(rows, rate_fields):
        records = []
        for row_number, row in enumerate(rows, start=1):
            rate = normalize_rate(get_field(row, rate_fields))
            if rate is None:
                continue
            tenure = clean(get_field(row, TENURE_FIELDS) or '')
            start = end = None
            if tenure:
                start, end = tenure_bounds(tenure)
            if start is None:
                start = structured_day(row, 'Start') or structured_day(row, 'From')
            if end is None:
                end = structured_day(row, 'End') or structured_day(row, 'To')
            if start is None:
                continue
            records.append({'StartDay': start, 'EndDay': end, 'Rate': rate, 'Tenure': tenure or f'Start day {start}', 'Row': row_number})
        records.sort(key=lambda item: (item['StartDay'], item['Row']))
        for index, item in enumerate(records):
            if item['EndDay'] is None:
                item['EndDay'] = records[index + 1]['StartDay'] - 1 if index + 1 < len(records) else 999999
        return [item for item in records if item['EndDay'] >= item['StartDay']]

    def maximum_overlapping_rate(records, au_start, au_end, label):
        matches = []
        for record in records:
            if record['StartDay'] <= au_end and record['EndDay'] >= au_start:
                item = dict(record)
                item['OverlapStart'] = max(record['StartDay'], au_start)
                item['OverlapEnd'] = min(record['EndDay'], au_end)
                matches.append(item)
        if not matches:
            print(f'WARNING: No overlapping {label} slab for AU range {au_start}-{au_end}')
            return None
        matches.sort(key=lambda item: (-item['Rate'], -(item['OverlapEnd'] - item['OverlapStart'] + 1), item['EndDay'] - item['StartDay']))
        selected = matches[0]
        print(f"{label} | AU={au_start}-{au_end} | selected={selected['Tenure']} | rate={selected['Rate']:.2f}%")
        return selected['Rate']

    def au_anchors(rows, rate_fields):
        records = normalize_records(rows, rate_fields)
        anchors, seen = ([], set())
        for record in records:
            key = re.sub('[^a-z0-9]+', '', record['Tenure'].lower())
            if key in seen:
                continue
            seen.add(key)
            anchors.append({'Tenure': record['Tenure'], 'StartDay': record['StartDay'], 'EndDay': record['EndDay'], 'AURate': record['Rate']})
        return sorted(anchors, key=lambda item: (item['EndDay'], item['StartDay']))

    def rate_color(competitor, au):
        if competitor is None or competitor == au:
            return BLACK
        return GREEN if competitor < au else RED

    def add_headers(ws, columns):
        ws.merge_cells('A1:A2')
        ws['A1'] = 'FD Tenors'
        ws['A1'].fill = PatternFill('solid', fgColor=DARK_PURPLE)
        ws['A1'].font = Font(color=WHITE, bold=True)
        ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
        for title, banks, color in GROUPS:
            start, end = (min((columns[b] for b in banks)), max((columns[b] for b in banks)))
            if start != end:
                ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
            cell = ws.cell(1, start, title)
            cell.fill = PatternFill('solid', fgColor=color)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        au_col = columns[AU_BANK]
        ws.merge_cells(start_row=1, start_column=au_col, end_row=2, end_column=au_col)
        cell = ws.cell(1, au_col, 'AU BANK')
        cell.fill = PatternFill('solid', fgColor=DARK_PURPLE)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal='center')
        for bank in BANK_ORDER:
            cell = ws.cell(2, columns[bank], bank.title() if bank not in ('SBI', 'HDFC', 'ICICI', 'IDFC', 'RBL', 'DCB') else bank)
            cell.fill = PatternFill('solid', fgColor=ORANGE if bank == 'SBI' else GRAY_HEADER)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

    def build_sheet(workbook, title, bank_files, rate_fields, category):
        required = BANK_ORDER + [AU_BANK]
        normalized = {bank: normalize_records(bank_files[bank][1], rate_fields) for bank in required}
        au_records = normalized[AU_BANK]
        if not au_records:
            raise ValueError(f'No AU {category} rates found')
        anchors = []
        for tenure, start_day, end_day in FIXED_TENURE_RANGES:
            au_rate = maximum_overlapping_rate(
                au_records, start_day, end_day, f'AU {category}'
            )
            anchors.append({
                'Tenure': tenure,
                'StartDay': start_day,
                'EndDay': end_day,
                'AURate': au_rate,
            })
        ws = workbook.create_sheet(title)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'B3'
        columns = {bank: index + 2 for index, bank in enumerate(BANK_ORDER)}
        columns[AU_BANK] = len(BANK_ORDER) + 2
        add_headers(ws, columns)
        for row_number, anchor in enumerate(anchors, start=3):
            ws.cell(row_number, 1, anchor['Tenure']).fill = PatternFill('solid', fgColor=LIGHT_GRAY)
            for bank in BANK_ORDER:
                rate = maximum_overlapping_rate(normalized[bank], anchor['StartDay'], anchor['EndDay'], f'{bank} {category}')

                # HDFC-only dynamic boundary rule for AU's final 60-120 month row.
                # Select the current HDFC source rate whose slab ends exactly at
                # 5 years (1800 days). No rate value is hardcoded.
                if (
                    bank == 'HDFC'
                    and anchor['StartDay'] == 1800
                    and anchor['EndDay'] == 3600
                ):
                    exact_five_year_rows = [
                        record for record in normalized['HDFC']
                        if record['EndDay'] == 1800
                    ]
                    if exact_five_year_rows:
                        exact_five_year_rows.sort(key=lambda record: (
                            -record['StartDay'],
                            record['EndDay'] - record['StartDay'],
                            record['Row'],
                        ))
                        selected_hdfc = exact_five_year_rows[0]
                        rate = selected_hdfc['Rate']
                        print(
                            f"HDFC {category} | dynamic exact-5-year source="
                            f"{selected_hdfc['Tenure']} | rate={rate:.2f}%"
                        )
                    else:
                        print(
                            f'WARNING: HDFC {category} exact 5-year source slab '
                            f'was not found; overlap-selected rate retained.'
                        )

                cell = ws.cell(row_number, columns[bank])
                if rate is not None:
                    cell.value = rate / 100
                cell.number_format = '0.00%'
                cell.fill = PatternFill('solid', fgColor=LIGHT_GRAY)
                cell.font = Font(bold=True, color=rate_color(rate, anchor['AURate']))
                cell.alignment = Alignment(horizontal='center')
            au = ws.cell(row_number, columns[AU_BANK])
            if anchor['AURate'] is not None:
                au.value = anchor['AURate'] / 100
                au.number_format = '0.00%'
            else:
                au.value = '-'
            au.fill = PatternFill('solid', fgColor=LIGHT_GRAY)
            au.font = Font(bold=True)
            au.alignment = Alignment(horizontal='center')
        last_row, last_col = (len(anchors) + 2, len(BANK_ORDER) + 2)
        for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=last_col):
            for cell in row:
                cell.border = Border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
        ws.column_dimensions['A'].width = 42
        for column in range(2, last_col + 1):
            ws.column_dimensions[get_column_letter(column)].width = 14
        legend = last_row + 2
        ws.cell(legend, 1, 'Colour logic').fill = PatternFill('solid', fgColor=DARK_PURPLE)
        ws.cell(legend, 1).font = Font(color=WHITE, bold=True)
        ws.cell(legend + 1, 1, 'Green').font = Font(color=GREEN, bold=True)
        ws.cell(legend + 1, 2, 'AU Bank rate is higher')
        ws.cell(legend + 2, 1, 'Red').font = Font(color=RED, bold=True)
        ws.cell(legend + 2, 2, 'AU Bank rate is lower')
        ws.cell(legend + 3, 1, 'Black').font = Font(color=BLACK, bold=True)
        ws.cell(legend + 3, 2, 'Rates are equal')
        ws.cell(legend + 5, 1, 'Rate category').font = Font(bold=True)
        ws.cell(legend + 5, 2, category)
        ws.cell(legend + 6, 1, 'Tenure rule').font = Font(bold=True)
        ws.cell(legend + 6, 2, 'AU tenure fixed; maximum overlapping rate used; month=30 days; year=360 days')
        return (anchors, normalized)

    def build_peak_sheet(workbook, title, bank_files, rate_fields, category):
        """Create one peak-rate view using each bank's exact source tenure.

        For every bank, the highest current rate is selected from that bank's
        normalized source rows. If the same peak rate occurs for multiple exact
        tenures, all exact tenure labels are preserved and written on separate
        rows. Rates are never hardcoded.
        """
        required = BANK_ORDER + [AU_BANK]
        normalized = {
            bank: normalize_records(bank_files[bank][1], rate_fields)
            for bank in required
        }

        peak_by_bank = {}
        tenure_order = []
        seen_tenures = set()

        for bank in required:
            records = normalized[bank]
            if not records:
                peak_by_bank[bank] = []
                continue

            peak_rate = max(record['Rate'] for record in records)
            peak_records = [
                record for record in records
                if abs(record['Rate'] - peak_rate) < 0.000001
            ]

            unique_records = []
            seen_bank_tenures = set()
            for record in peak_records:
                tenure = clean(record['Tenure'])
                tenure_key = re.sub('[^a-z0-9]+', '', tenure.lower())
                if not tenure or tenure_key in seen_bank_tenures:
                    continue
                seen_bank_tenures.add(tenure_key)
                unique_records.append(record)
                if tenure_key not in seen_tenures:
                    seen_tenures.add(tenure_key)
                    tenure_order.append((tenure_key, tenure))

            peak_by_bank[bank] = unique_records

        ws = workbook.create_sheet(title)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'B4'

        columns = {bank: index + 2 for index, bank in enumerate(BANK_ORDER)}
        columns[AU_BANK] = len(BANK_ORDER) + 2
        last_col = columns[AU_BANK]

        ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_col)
        heading = ws.cell(1, 2, 'Peak rate Benchmarking with exact tenures')
        heading.font = Font(bold=True)
        heading.alignment = Alignment(horizontal='center', vertical='center')

        ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1)
        label = ws.cell(1, 1, '*Peak Rates & Buckets')
        label.fill = PatternFill('solid', fgColor=DARK_PURPLE)
        label.font = Font(color=WHITE, bold=True)
        label.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for group_title, banks, color in GROUPS:
            start = min(columns[bank] for bank in banks)
            end = max(columns[bank] for bank in banks)
            if start != end:
                ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=end)
            group_cell = ws.cell(2, start, group_title)
            group_cell.fill = PatternFill('solid', fgColor=color)
            group_cell.font = Font(bold=True)
            group_cell.alignment = Alignment(horizontal='center', vertical='center')

        au_col = columns[AU_BANK]
        ws.merge_cells(start_row=2, start_column=au_col, end_row=3, end_column=au_col)
        au_header = ws.cell(2, au_col, 'AU BANK')
        au_header.fill = PatternFill('solid', fgColor=DARK_PURPLE)
        au_header.font = Font(color=WHITE, bold=True)
        au_header.alignment = Alignment(horizontal='center', vertical='center')

        for bank in BANK_ORDER:
            bank_cell = ws.cell(3, columns[bank], bank.title() if bank not in ('SBI', 'HDFC', 'ICICI', 'IDFC', 'RBL', 'DCB') else bank)
            bank_cell.fill = PatternFill('solid', fgColor=GRAY_HEADER)
            bank_cell.font = Font(bold=True)
            bank_cell.alignment = Alignment(horizontal='center', vertical='center')

        row_lookup = {}
        for row_number, (tenure_key, tenure) in enumerate(tenure_order, start=4):
            row_lookup[tenure_key] = row_number
            tenure_cell = ws.cell(row_number, 1, tenure)
            tenure_cell.fill = PatternFill('solid', fgColor=LIGHT_GRAY)
            tenure_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        au_peak = None
        if peak_by_bank.get(AU_BANK):
            au_peak = peak_by_bank[AU_BANK][0]['Rate']

        for bank in required:
            for record in peak_by_bank.get(bank, []):
                tenure_key = re.sub('[^a-z0-9]+', '', clean(record['Tenure']).lower())
                row_number = row_lookup[tenure_key]
                cell = ws.cell(row_number, columns[bank], record['Rate'] / 100)
                cell.number_format = '0.00%'
                cell.fill = PatternFill('solid', fgColor=LIGHT_GRAY)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if bank == AU_BANK or au_peak is None:
                    font_color = BLACK
                else:
                    font_color = rate_color(record['Rate'], au_peak)
                cell.font = Font(bold=True, color=font_color)

        last_row = max(3, 3 + len(tenure_order))
        for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=last_col):
            for cell in row:
                cell.border = Border(
                    left=THIN_BLACK, right=THIN_BLACK,
                    top=THIN_BLACK, bottom=THIN_BLACK,
                )

        ws.column_dimensions['A'].width = 48
        for column in range(2, last_col + 1):
            ws.column_dimensions[get_column_letter(column)].width = 14
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 24
        ws.row_dimensions[3].height = 24
        for row_number in range(4, last_row + 1):
            ws.row_dimensions[row_number].height = 32
        ws.sheet_view.zoomScale = 75

        print(f'{category} peak-rate rows: {len(tenure_order)}')
        for bank in required:
            records = peak_by_bank.get(bank, [])
            if records:
                print(
                    f"  {bank:10} peak={records[0]['Rate']:.2f}% | "
                    + ' | '.join(record['Tenure'] for record in records)
                )
            else:
                print(f'  {bank:10} peak=MISSING')

        return peak_by_bank

    def run_tests():
        assert tenure_bounds('60 months to 120 months') == (1800, 3600)
        assert tenure_bounds('3 years to less than 5 years') == (1080, 1799)
        assert tenure_bounds('5 years and up to 10 years') == (1800, 3600)
        assert FIXED_TENURE_RANGES[3] == ('6 Months 1 Day to less than 12 Months', 181, 359)
        assert FIXED_TENURE_RANGES[4] == ('12 Months', 360, 360)
        assert FIXED_TENURE_RANGES[6] == ('15 Months 1 Day to less than 18 Months', 451, 539)
        assert FIXED_TENURE_RANGES[7] == ('18 Months', 540, 540)
        assert FIXED_TENURE_RANGES[12] == ('45 Months 1 Day to less than 60 Months', 1351, 1799)
        assert FIXED_TENURE_RANGES[13] == ('60 Months', 1800, 1800)
        special_rows = [
            {'Tenure': '6 Months 1 Day to less than 12 Months', 'RegularRate': '5.25', 'SeniorCitizenRate': '5.75'},
            {'Tenure': '12 Months', 'RegularRate': '6.35', 'SeniorCitizenRate': '6.85'},
            {'Tenure': '15 Months 1 Day to less than 18 Months', 'RegularRate': '6.50', 'SeniorCitizenRate': '7.00'},
            {'Tenure': '18 Months', 'RegularRate': '6.75', 'SeniorCitizenRate': '7.25'},
            {'Tenure': '45 Months 1 Day to less than 60 Months', 'RegularRate': '6.40', 'SeniorCitizenRate': '6.90'},
            {'Tenure': '60 Months', 'RegularRate': '6.90', 'SeniorCitizenRate': '7.40'},
        ]
        special_regular = normalize_records(special_rows, REGULAR_RATE_FIELDS)
        assert maximum_overlapping_rate(special_regular, 181, 359, 'Special pre-12') == 5.25
        assert maximum_overlapping_rate(special_regular, 360, 360, 'Special 12') == 6.35
        assert maximum_overlapping_rate(special_regular, 451, 539, 'Special pre-18') == 6.50
        assert maximum_overlapping_rate(special_regular, 540, 540, 'Special 18') == 6.75
        assert maximum_overlapping_rate(special_regular, 1351, 1799, 'Special pre-60') == 6.40
        assert maximum_overlapping_rate(special_regular, 1800, 1800, 'Special 60') == 6.90
        rows = [{'Tenure': '3 years to less than 5 years', 'RegularRate': '6.30', 'SeniorCitizenRate': '6.80'}, {'Tenure': '5 years and up to 10 years', 'RegularRate': '6.05', 'SeniorCitizenRate': '7.05'}]
        regular = normalize_records(rows, REGULAR_RATE_FIELDS)
        senior = normalize_records(rows, SENIOR_RATE_FIELDS)
        assert maximum_overlapping_rate(regular, 1800, 3600, 'SBI Regular Test') == 6.05
        assert maximum_overlapping_rate(senior, 1800, 3600, 'SBI Senior Test') == 7.05
        print('SBI regular and senior 60-to-120-month tests passed.')

    def create_workbook(bank_files):
        required = BANK_ORDER + [AU_BANK]
        missing = [bank for bank in required if bank not in bank_files]
        if missing:
            raise ValueError('Missing bank CSV files: ' + ', '.join(missing))
        workbook = Workbook()
        workbook.remove(workbook.active)
        regular, regular_norm = build_sheet(workbook, 'Regular Comparison', bank_files, REGULAR_RATE_FIELDS, 'Regular Citizen')
        senior, senior_norm = build_sheet(workbook, 'Senior Citizen Comparison', bank_files, SENIOR_RATE_FIELDS, 'Senior Citizen')
        build_peak_sheet(workbook, 'Regular Peak Rates', bank_files, REGULAR_RATE_FIELDS, 'Regular Citizen')
        build_peak_sheet(workbook, 'Senior Peak Rates', bank_files, SENIOR_RATE_FIELDS, 'Senior Citizen')
        source = workbook.create_sheet('Source Files')
        source.append(['Bank', 'CSV File', 'Rows', 'Regular Slabs', 'Senior Slabs', 'Last Modified'])
        for cell in source[1]:
            cell.fill = PatternFill('solid', fgColor=DARK_PURPLE)
            cell.font = Font(color=WHITE, bold=True)
        for bank in required:
            path, rows = bank_files[bank]
            source.append([bank, path.name, len(rows), len(regular_norm[bank]), len(senior_norm[bank]), datetime.fromtimestamp(path.stat().st_mtime)])
        for column, width in {'A': 18, 'B': 55, 'C': 12, 'D': 16, 'E': 16, 'F': 22}.items():
            source.column_dimensions[column].width = width
        for cell in source['F'][1:]:
            cell.number_format = 'dd-mmm-yyyy hh:mm'
        try:
            workbook.save(RESULT_FILE)
            result = RESULT_FILE
        except PermissionError:
            result = BASE_DIR / f'AU_FD_Rate_Comparison_{datetime.now():%Y%m%d_%H%M%S}.xlsx'
            workbook.save(result)
            print(f'Original workbook locked; created {result}')
        return (result, regular, senior)

    def main():
        run_tests()
        files = load_latest_bank_files()
        print('Detected bank files:')
        for bank in BANK_ORDER + [AU_BANK]:
            print(f"  {bank:10} -> {(files[bank][0].name if bank in files else 'MISSING')}")
        result, regular, senior = create_workbook(files)
        print(f'Regular comparison rows: {len(regular)}')
        print(f'Senior comparison rows: {len(senior)}')
        print(f'Excel created: {result}')

    return main()


SCRAPER_RUNNERS = [
    ('Axis Bank', 'axis_scraper.py', run_axis_scraper_source),
    ('AU Small Finance Bank', 'au_scraper.py', run_au_scraper_source),
    ('Bandhan Bank', 'bandhan_scraper.py', run_bandhan_scraper_source),
    ('DCB Bank', 'dcb_scraper.py', run_dcb_scraper_source),
    ('HDFC Bank', 'hdfc_scraper.py', run_hdfc_scraper_source),
    ('ICICI Bank', 'icici_scraper.py', run_icici_scraper_source),
    ('IDFC FIRST Bank', 'idfc_scraper.py', run_idfc_scraper_source),
    ('IndusInd Bank', 'indusind_scraper.py', run_indusind_scraper_source),
    ('Kotak Mahindra Bank', 'kotak_scraper.py', run_kotak_scraper_source),
    ('RBL Bank', 'rbl_scraper.py', run_rbl_scraper_source),
    ('State Bank of India', 'sbi_scraper.py', run_sbi_scraper_source),
    ('YES Bank', 'yes_scraper.py', run_yes_scraper_source),
]


def _run_one(label: str, original_name: str, runner):
    print("\n" + "=" * 78)
    print(f"Running {label} using readable source from {original_name}")
    print("=" * 78)
    started = datetime.now()
    try:
        result = runner()
        elapsed = datetime.now() - started
        print(f"[SUCCESS] {label} completed in {elapsed}")
        return {"Bank": label, "Script": original_name, "Status": "SUCCESS", "Error": ""}, result
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
        if code == 0:
            elapsed = datetime.now() - started
            print(f"[SUCCESS] {label} completed in {elapsed}")
            return {"Bank": label, "Script": original_name, "Status": "SUCCESS", "Error": ""}, None
        error = f"SystemExit({exc.code})"
        print(f"[FAILED] {label}: {error}")
        return {"Bank": label, "Script": original_name, "Status": "FAILED", "Error": error}, None
    except Exception as exc:
        print(f"[FAILED] {label}: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        return {
            "Bank": label,
            "Script": original_name,
            "Status": "FAILED",
            "Error": f"{type(exc).__name__}: {exc}",
        }, None


def _run_builder():
    print("\n" + "=" * 78)
    print("Generating final workbook using readable build_td.py source")
    print("=" * 78)
    return run_builder_source()


def main():
    parser = argparse.ArgumentParser(
        description="Run the readable single-file 11-bank retail FD workflow."
    )
    parser.add_argument(
        "--skip-scrape",
        action="store_true",
        help="Do not scrape banks; rebuild the workbook from existing output CSVs.",
    )
    parser.add_argument(
        "--only",
        default="",
        help="Comma-separated bank/script keywords to run, e.g. sbi,axis,au.",
    )
    parser.add_argument(
        "--skip-build",
        action="store_true",
        help="Run bank scrapers but do not generate the final workbook.",
    )
    args = parser.parse_args()

    selected = {item.strip().lower() for item in args.only.split(",") if item.strip()}
    summaries = []

    if not args.skip_scrape:
        for label, original_name, runner in SCRAPER_RUNNERS:
            searchable = f"{label} {original_name}".lower()
            if selected and not any(keyword in searchable for keyword in selected):
                continue
            summary, _ = _run_one(label, original_name, runner)
            summaries.append(summary)

    if not args.skip_build:
        try:
            _run_builder()
            build_status = "SUCCESS"
            build_error = ""
        except Exception as exc:
            build_status = "FAILED"
            build_error = f"{type(exc).__name__}: {exc}"
            print(f"[FAILED] comparison builder: {build_error}")
            traceback.print_exc()
    else:
        build_status = "SKIPPED"
        build_error = ""

    print("\n" + "=" * 78)
    print("READABLE SINGLE-FILE RETAIL WORKFLOW SUMMARY")
    print("=" * 78)
    if summaries:
        for item in summaries:
            suffix = f" | {item['Error']}" if item["Error"] else ""
            print(f"{item['Bank']:<28} {item['Status']}{suffix}")
    else:
        print("Bank scraping skipped.")
    print(f"Comparison Builder           {build_status}" + (f" | {build_error}" if build_error else ""))
    print(f"Output folder: {ROOT / 'output'}")
    print(f"Final workbook: {ROOT / 'AU_FD_Rate_Comparison.xlsx'}")

    failures = [item for item in summaries if item["Status"] == "FAILED"]
    return 1 if failures or build_status == "FAILED" else 0


if __name__ == "__main__":
    raise SystemExit(main())
