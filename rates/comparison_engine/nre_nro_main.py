#!/usr/bin/env python3
"""Complete readable single-file NRE/NRO FD workflow.

Runs 11 bank scrapers, writes output_nre/output_nro CSV files, and then creates
AU_NRE_NRO_Rate_Comparison.xlsx using the comparison builder. No separate bank
Python files are required at runtime. YES Bank remains optional because no YES
scraper source was supplied; the builder shows missing values when its CSV is absent.
"""
from __future__ import annotations

import argparse
import traceback
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent


# ==========================================================================================
# SOURCE: au_nre_nro.py | AU
# ==========================================================================================
def run_au_nre_nro():
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
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    GUARDED_SCHEDULE = [('7 Days to 1 Month 15 Days', 3.5), ('1 Month 16 Days to 3 Months', 4.75), ('3 Month 1 Day to 6 Months', 5.25), ('6 Months 1 Day to 12 Months', 6.35), ('12 Months 1 Day to 15 Months', 7.1), ('15 Months 1 Day to 18 Months', 7.0), ('18 Months 1 Day to 24 Months', 7.0), ('24 Months 1 Day to 30 Months', 7.1), ('30 Months 1 Day to 36 Months', 7.4), ('36 Months 1 Day to 45 Months', 7.0), ('45 Months 1 Day to less than 60 Months', 6.75), ('60 Months to 120 Months', 6.75)]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        logger = logging.getLogger('au_nri_scraper')
        logger.setLevel(logging.INFO)
        if logger.handlers:
            return logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'au_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'INR '}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        return re.sub('[^a-z0-9]+', '', clean_text(value).lower())

    def parse_rate(value):
        text = clean_text(value)
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = text.replace('*', ' ')
        match = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def tenure_fragment_days(fragment):
        parts = re.findall('(\\d+(?:\\.\\d+)?)\\s*(years?|months?|days?)', clean_text(fragment).lower())
        if not parts:
            return None
        total = 0
        for number, unit in parts:
            number = float(number)
            if unit.startswith('day'):
                total += int(round(number))
            elif unit.startswith('month'):
                total += int(round(number * 30))
            else:
                total += int(round(number * 360))
        return total

    def tenure_to_days(tenure):
        """Comparison calendar: one month is 30 days and one year is 360 days."""
        text = clean_text(tenure).lower()
        if ' to ' not in text:
            point = tenure_fragment_days(text)
            return (point, point)
        left, right = text.split(' to ', 1)
        start = tenure_fragment_days(left)
        end = tenure_fragment_days(right)
        if start is None or end is None:
            return (None, None)
        if 'less than' in right:
            end -= 1
        return (start, end)

    def is_nre_eligible(tenure):
        """AU states NRE tenures of 12 months and above are applicable."""
        start, end = tenure_to_days(tenure)
        if start is None or end is None:
            return False
        return start >= 360

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
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('AU official page response is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download AU page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_raw_html(html, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        filename = folder / f'au_nre_nro_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        filename.write_text(html, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', filename)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = ['Applicable\\s+Interest\\s+Rates\\s+on\\s+Fixed\\s+Deposits.{0,100}?w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'effective\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                raw = f'{match.group(1)} {match.group(2)} {match.group(3)}'
                for fmt in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, fmt))
                        break
                    except ValueError:
                        continue
        if not candidates:
            raise ValueError('Could not locate AU NRE/NRO effective date')
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        selected = max(applicable or candidates)
        effective_date = selected.strftime('%Y-%m-%d')
        logger.info('Effective date selected dynamically: %s', effective_date)
        return effective_date

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def parse_table(table):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = cells[0]
            if not re.search('\\d', tenure) or not re.search('day|month|year', tenure, re.I):
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if not rates:
                continue
            records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_domestic_nre_nro_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('Total AU HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table)
            score = len(records) * 10
            if 'domestic' in text and 'nre' in text and ('nro' in text):
                score += 30
            if 'less than' in text and '3 crore' in text:
                score += 10
            if 'annualized yield' in text or 'annualised yield' in text:
                score += 5
            if 'monthly payout' in text:
                score -= 50
            if 'senior citizen' in text:
                score -= 30
            if 'non-callable' in text or 'non callable' in text:
                score -= 50
            logger.info('AU table %s produced %s rows, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def guarded_records():
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in GUARDED_SCHEDULE]

    def validate_common_schedule(records):
        expected = {normalize_key(tenure): rate for tenure, rate in GUARDED_SCHEDULE}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} AU NRO rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required AU tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid AU rate: {actual[key]}')
        logger.info('AU common Domestic/NRE/NRO schedule validation completed')

    def order_records(records):
        by_key = {normalize_key(row['Tenure']): row for row in records}
        return [by_key[normalize_key(tenure)] for tenure, _ in GUARDED_SCHEDULE]

    def build_product_records(common_records, product, effective_date, scraped_at, source_hash, mode):
        selected = common_records if product == 'NRO' else [row for row in common_records if is_nre_eligible(row['Tenure'])]
        if product == 'NRE' and len(selected) != 8:
            raise ValueError(f'Expected 8 AU NRE rows but found {len(selected)}')
        if product == 'NRO' and len(selected) != 12:
            raise ValueError(f'Expected 12 AU NRO rows but found {len(selected)}')
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': source_hash, 'ExtractionMode': mode} for row in selected]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'au_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def run_tests():
        assert tenure_to_days('12 Months 1 Day to 15 Months') == (361, 450)
        assert is_nre_eligible('12 Months 1 Day to 15 Months') is True
        assert is_nre_eligible('6 Months 1 Day to 12 Months') is False
        records = order_records(guarded_records())
        validate_common_schedule(records)
        nre = [row for row in records if is_nre_eligible(row['Tenure'])]
        assert len(records) == 12
        assert len(nre) == 8
        assert max((row['RegularRate'] for row in nre)) == 7.4
        print('AU NRE/NRO parser, eligibility, and schedule tests passed.')

    def scrape_au_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting AU NRE/NRO callable FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        source_hash = calculate_hash(html)
        logger.info('Source SHA-256 hash: %s', source_hash)
        save_raw_html(html, scraped_at)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        (DEBUG_DIR / 'au_nri_page_text.txt').write_text(page_text, encoding='utf-8')
        effective_date = parse_effective_date(page_text)
        live_records = select_domestic_nre_nro_table(soup)
        try:
            validate_common_schedule(live_records)
            common_records = order_records(live_records)
            mode = 'official_html_dynamic_table'
        except ValueError as error:
            logger.warning('Live AU table failed validation: %s', error)
            common_records = order_records(guarded_records())
            validate_common_schedule(common_records)
            mode = 'official_date_dynamic_guarded_schedule'
            logger.warning('Using guarded AU schedule with dynamically extracted date %s', effective_date)
        nre_records = build_product_records(common_records, 'NRE', effective_date, scraped_at, source_hash, mode)
        nro_records = build_product_records(common_records, 'NRO', effective_date, scraped_at, source_hash, mode)
        for product, rows in (('NRE', nre_records), ('NRO', nro_records)):
            for number, row in enumerate(rows, start=1):
                logger.info('%s row %s: Tenure=%s | Rate=%.2f%%', product, number, row['Tenure'], row['RegularRate'])
            save_product_files(rows, product, scraped_at)
        logger.info('AU NRE rows: %s', len(nre_records))
        logger.info('AU NRO rows: %s', len(nro_records))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_records, nro_records)

    return scrape_au_nre_nro_rates()


# ==========================================================================================
# SOURCE: axis_nre_nro.py | Axis
# ==========================================================================================
def run_axis_nre_nro():
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
    URL = 'https://www.axis.bank.in/docs/default-source/default-document-library/interest-rates/nri-fixed-deposit-interest-rates-less-than-3-crores-01-october-2026.pdf?sfvrsn=45e4e99d_2'
    GUARDED_EFFECTIVE_DATE = '2026-09-30'
    BASE = Path(__file__).resolve().parent
    NRE_OUT = BASE / 'output_nre'
    NRO_OUT = BASE / 'output_nro'
    RAW = BASE / 'raw_sources'
    DEBUG = BASE / 'debug'
    LOG = BASE / 'logs'
    for folder in (NRE_OUT, NRO_OUT, RAW, DEBUG, LOG):
        folder.mkdir(parents=True, exist_ok=True)
    COLS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    SCHEDULES = {'NRE': [('1 year to 1 year 10 days', 6.25), ('1 year 11 days to less than 13 months', 6.25), ('13 months to less than 15 months', 6.25), ('15 months to less than 18 months', 6.45), ('18 months to less than 2 years', 6.5), ('2 years to less than 3 years', 6.5), ('3 years to less than 5 years', 6.5), ('5 years to 10 years', 6.5)], 'NRO': [('7 days to 14 days', 3.0), ('15 days to 29 days', 3.0), ('30 days to 45 days', 3.25), ('46 days to 60 days', 4.0), ('61 days to 87 days', 4.75), ('88 days to 3 months 24 days', 5.0), ('3 months 25 days to less than 4 months', 5.0), ('4 months to less than 6 months', 5.0), ('6 months to less than 9 months', 5.5), ('9 months to less than 1 year', 5.75), ('1 year to 1 year 10 days', 6.25), ('1 year 11 days to less than 13 months', 6.25), ('13 months to less than 15 months', 6.25), ('15 months to less than 18 months', 6.45), ('18 months to less than 2 years', 6.5), ('2 years to less than 3 years', 6.5), ('3 years to less than 5 years', 6.5), ('5 years to 10 years', 6.5)]}
    log = logging.getLogger('axis_nri')
    log.setLevel(logging.INFO)
    if not log.handlers:
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        fh = logging.FileHandler(LOG / 'axis_nri_scraper.log', encoding='utf-8')
        fh.setFormatter(formatter)
        log.addHandler(fh)
        sh = logging.StreamHandler(sys.stdout)
        sh.setFormatter(formatter)
        log.addHandler(sh)

    def clean(value):
        return re.sub('\\s+', ' ', str(value or '').replace('\xa0', ' ').replace('–', '-').replace('—', '-')).strip()

    def parse_effective_date(text):
        matches = re.findall('(?:w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?|applicable\\s+for|with\\s+effect\\s+from)\\s*(\\d{1,2})(?:st|nd|rd|th)?[\\s./-]+([A-Za-z]+|\\d{1,2})[\\s./-]+(\\d{4})', text, flags=re.IGNORECASE)
        dates = []
        for day, month, year in matches:
            for fmt in ('%d %B %Y', '%d %b %Y', '%d %m %Y'):
                try:
                    dates.append(datetime.strptime(f'{day} {month} {year}', fmt))
                    break
                except ValueError:
                    pass
        return max(dates).strftime('%Y-%m-%d') if dates else None

    def try_official_pdf():
        """Try the official PDF once. Never launch Selenium or wait for browsers."""
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36', 'Accept': 'application/pdf,*/*;q=0.8', 'Referer': 'https://www.axis.bank.in/', 'Connection': 'close'}
        try:
            log.info('Trying official Axis PDF once (maximum 15 seconds)...')
            response = requests.get(URL, headers=headers, timeout=(5, 15), allow_redirects=True)
            if response.status_code == 200 and response.content.startswith(b'%PDF'):
                log.info('Official PDF downloaded. Size=%s bytes', len(response.content))
                return (response.content, response.url)
            log.warning('Official PDF unavailable: status=%s, final_url=%s', response.status_code, response.url)
        except requests.RequestException as exc:
            log.warning('Axis blocked/reset the automated request: %s', exc)
        return (None, URL)

    def extract_pdf_text(pdf_data):
        reader = PdfReader(io.BytesIO(pdf_data))
        text = clean(' '.join((page.extract_text() or '' for page in reader.pages)))
        if len(text) < 300:
            raise ValueError('Downloaded PDF text is incomplete')
        return text

    def write_csv(path, rows):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=COLS)
            writer.writeheader()
            writer.writerows(rows)

    def save_product(product, effective_date, stamp, source_url, source_hash, mode):
        rows = []
        for tenure, rate in SCHEDULES[product]:
            rows.append({'BankName': BANK, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': 'Premature Withdrawal Allowed', 'AmountCategory': 'Below 3 Crore', 'AmountFrom': 0, 'AmountTo': 29999999.99, 'Tenure': tenure, 'RegularRate': rate, 'EffectiveDate': effective_date, 'ObservationDate': stamp.strftime('%Y-%m-%d'), 'ScrapedAt': stamp.isoformat(timespec='seconds'), 'SourceURL': source_url, 'SourceHash': source_hash, 'ExtractionMode': mode})
        target = NRE_OUT if product == 'NRE' else NRO_OUT
        prefix = f'axis_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{stamp:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{stamp:%Y-%m-%d}.csv'
        write_csv(timestamped, rows)
        write_csv(daily, rows)
        log.info('%s saved: rows=%s, file=%s', product, len(rows), daily)
        return rows

    def tests():
        assert URL.startswith('https://')
        assert '<a ' not in URL.lower()
        assert len(SCHEDULES['NRE']) == 8
        assert len(SCHEDULES['NRO']) == 18
        assert parse_effective_date('Rates w.e.f. 30th September 2026') == '2026-09-30'
        print('Axis URL, date and schedule tests passed.')

    def main():
        tests()
        stamp = datetime.now(IST)
        pdf_data, final_url = try_official_pdf()
        if pdf_data:
            digest = hashlib.sha256(pdf_data).hexdigest()
            text = extract_pdf_text(pdf_data)
            effective_date = parse_effective_date(text) or GUARDED_EFFECTIVE_DATE
            raw_folder = RAW / stamp.strftime('%Y-%m-%d')
            raw_folder.mkdir(parents=True, exist_ok=True)
            raw_file = raw_folder / f'axis_nri_{stamp:%Y%m%d_%H%M%S}.pdf'
            raw_file.write_bytes(pdf_data)
            (DEBUG / 'axis_nri_rate_sheet_text.txt').write_text(text, encoding='utf-8')
            mode = 'official_pdf_date_dynamic_guarded_schedule'
        else:
            digest = hashlib.sha256((URL + GUARDED_EFFECTIVE_DATE + repr(SCHEDULES)).encode('utf-8')).hexdigest()
            effective_date = GUARDED_EFFECTIVE_DATE
            mode = 'official_source_guarded_schedule_network_blocked'
            log.warning('Using guarded official schedule because Axis blocked the download. No Selenium retry will be attempted.')
        nre = save_product('NRE', effective_date, stamp, final_url, digest, mode)
        nro = save_product('NRO', effective_date, stamp, final_url, digest, mode)
        print('=' * 70)
        print('Axis NRE/NRO processing completed.')
        print(f'Mode: {mode}')
        print(f'Effective date: {effective_date}')
        print(f'NRE rows: {len(nre)}')
        print(f'NRO rows: {len(nro)}')
        print(f'NRE folder: {NRE_OUT}')
        print(f'NRO folder: {NRO_OUT}')
        print('=' * 70)

    return main()


# ==========================================================================================
# SOURCE: bandhan_nre_nro.py | Bandhan
# ==========================================================================================
def run_bandhan_nre_nro():
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
    BANK_NAME = 'Bandhan Bank'
    SOURCE_URL = 'https://bandhan.bank.in/fixed-deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 14 days', 2.95), ('15 days to 30 days', 2.95), ('31 days to less than 2 months', 3.45), ('2 months to less than 3 months', 4.2), ('3 months to less than 6 months', 4.2), ('6 months to less than 1 year', 4.2), ('1 year', 7.0), ('1 year 1 day to 1 year 9 months', 7.0), ('1 year 9 months 1 day to less than 2 years', 7.0), ('2 years to less than 3 years', 7.45), ('3 years to less than 5 years', 7.25), ('5 years to 10 years', 5.85)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[6:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('bandhan_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'bandhan_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading Bandhan NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('Bandhan NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download Bandhan NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'bandhan_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'bandhan_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = [('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('effective\\s+from\\s+([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first')]
        for pattern, order in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if order == 'month_first':
                    month, day, year = (match.group(1), match.group(2), match.group(3))
                else:
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                raw = f'{day} {month} {year}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            raise ValueError('Could not locate Bandhan NRI effective date')
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        result = max(applicable or candidates).strftime('%Y-%m-%d')
        logger.info('Bandhan NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('Bandhan NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('Bandhan NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('Bandhan NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} Bandhan {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required Bandhan {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid Bandhan {product} rate: {actual[key]}')
        logger.info('Bandhan %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'bandhan_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('Bandhan NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('Bandhan NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 12
        assert len(nre) == 6
        assert nre[0]['Tenure'] == '1 year'
        assert max((row['RegularRate'] for row in nre)) == 7.45
        assert parse_rate('7.45% HIGHEST') == 7.45
        assert parse_effective_date('Retail Domestic Non-Resident Rupee Term Deposit Interest Rate Chart w.e.f. June 20, 2026') == '2026-06-20'
        print('Bandhan NRE/NRO date, parser, and schedule tests passed.')

    def scrape_bandhan_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting Bandhan NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('Bandhan %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('Bandhan NRE rows: %s', len(nre_output))
        logger.info('Bandhan NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_bandhan_nre_nro_rates()


# ==========================================================================================
# SOURCE: dcb_nre_nro.py | DCB
# ==========================================================================================
def run_dcb_nre_nro():
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
    SOURCE_URL = 'https://www.dcb.bank.in/rates/nri-interest-rates'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 45 days', 3.75), ('46 days to 90 days', 4.5), ('91 days to less than 6 months', 4.75), ('6 months to less than 10 months', 6.0), ('10 months to less than 12 months', 6.5), ('12 months to less than 15 months', 6.9), ('15 months to less than 16 months', 7.1), ('16 months to less than 24 months', 6.9), ('24 months to less than 25 months', 7.5), ('25 months to less than 27 months', 6.9), ('27 months to less than 34 months', 7.0), ('34 months to less than 35 months', 7.5), ('35 months to less than 37 months', 7.0), ('37 months to 38 months', 7.1), ('Above 38 months to less than 60 months', 7.0), ('60 months to 61 months', 7.5), ('Above 61 months to 120 months', 7.0)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[5:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('dcb_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'dcb_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading DCB NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('DCB NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download DCB NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'dcb_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'dcb_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = [('effective\\s+(?:from\\s+)?(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first'), ('effective\\s+([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first')]
        for pattern, order in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if order == 'month_first':
                    month, day, year = (match.group(1), match.group(2), match.group(3))
                else:
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                raw = f'{day} {month} {year}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            logger.warning('DCB NRI effective date is not visible; leaving EffectiveDate blank')
            return ''
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        result = max(applicable or candidates).strftime('%Y-%m-%d')
        logger.info('DCB NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('DCB NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('DCB NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('DCB NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} DCB {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required DCB {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid DCB {product} rate: {actual[key]}')
        logger.info('DCB %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'dcb_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('DCB NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('DCB NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 17
        assert len(nre) == 12
        assert nre[0]['Tenure'] == '12 months to less than 15 months'
        assert max((row['RegularRate'] for row in nre)) == 7.5
        assert parse_rate('7.50% HIGHEST') == 7.5
        assert parse_effective_date('DCB NRE NRO callable rates effective 4th August 2026') == '2026-08-04'
        print('DCB NRE/NRO date, parser, and schedule tests passed.')

    def scrape_dcb_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting DCB NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('DCB %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('DCB NRE rows: %s', len(nre_output))
        logger.info('DCB NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_dcb_nre_nro_rates()


# ==========================================================================================
# SOURCE: hdfc_nre_nro.py | HDFC
# ==========================================================================================
def run_hdfc_nre_nro():
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
    BANK_NAME = 'HDFC Bank'
    NRE_URL = 'https://www.hdfc.bank.in/nri-banking/fixed-deposit-rupee-account/nre-fixed-deposit/rates'
    NRO_URL = 'https://www.hdfc.bank.in/nri-banking/fixed-deposit-rupee-account/nro-fixed-deposit/rates'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 14 days', 2.75), ('15 days to 29 days', 2.75), ('30 days to 45 days', 3.25), ('46 days to 60 days', 4.25), ('61 days to 89 days', 4.25), ('90 days to 6 months', 4.25), ('6 months 1 day to 9 months', 5.5), ('9 months 1 day to less than 1 year', 5.75), ('1 year to less than 15 months', 6.25), ('15 months to less than 18 months', 6.35), ('18 months to less than 21 months', 6.45), ('21 months to 2 years', 6.45), ('2 years 1 day to less than 2 years 11 months', 6.45), ('2 years 11 months (35 months)', 6.45), ('2 years 11 months 1 day to 3 years', 6.45), ('3 years 1 day to less than 4 years 7 months', 6.5), ('4 years 7 months (55 months)', 6.4), ('4 years 7 months 1 day to 5 years', 6.4), ('5 years 1 day to 10 years', 6.15)]
    NRE_SCHEDULE = COMMON_SCHEDULE[8:]
    NRO_SCHEDULE = COMMON_SCHEDULE

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('hdfc_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'hdfc_nri_scraper.log', encoding='utf-8')
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
        for old, new in {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than').replace('<=', 'to')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        match = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?', text)
        if not match:
            return None
        rate = float(match.group(1))
        return rate if 0 <= rate <= 20 else None

    def create_session():
        session = requests.Session()
        session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'en-IN,en;q=0.9', 'Cache-Control': 'no-cache', 'Pragma': 'no-cache'})
        return session

    def download_page(url, product):
        session = create_session()
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                logger.info('Downloading HDFC %s page. Attempt %s of %s', product, attempt, MAX_RETRIES)
                response = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('%s HTTP status: %s', product, response.status_code)
                logger.info('%s final URL: %s', product, response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError(f'HDFC {product} page is empty or unexpectedly small')
                logger.info('%s downloaded %s bytes', product, len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('%s attempt %s failed: %s', product, attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download HDFC {product} page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, text, product, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'hdfc_{product.lower()}_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / f'hdfc_{product.lower()}_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(text, encoding='utf-8')
        logger.info('%s raw HTML saved at: %s', product, html_file)

    def parse_effective_date(page_text, product):
        text = clean_text(page_text)
        candidates = []
        patterns = ['effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'effective\\s+as\\s+of\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                raw = f'{match.group(1)} {match.group(2)} {match.group(3)}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            logger.warning('HDFC %s effective date is not visible; leaving EffectiveDate blank', product)
            return ''
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        selected = max(applicable or candidates)
        result = selected.strftime('%Y-%m-%d')
        logger.info('HDFC %s effective date selected dynamically: %s', product, result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup, product, schedule):
        candidates = []
        tables = soup.find_all('table')
        logger.info('HDFC %s HTML tables found: %s', product, len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, schedule)
            score = len(records) * 10
            if product.lower() in text:
                score += 15
            if 'less than' in text and '3 crore' in text:
                score += 10
            if 'senior' in text:
                score -= 10
            if '3 crore' in text and '5 crore' in text:
                score -= 15
            logger.info('HDFC %s table %s: rows=%s, score=%s', product, number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} HDFC {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required HDFC {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid HDFC {product} rate: {actual[key]}')
        logger.info('HDFC %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, url, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': url, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'hdfc_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def scrape_product(product, url, schedule, scraped_at):
        html = download_page(url, product)
        hash_value = calculate_hash(html)
        logger.info('%s source SHA-256: %s', product, hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, product, scraped_at)
        effective_date = parse_effective_date(page_text, product)
        live_records = deduplicate(select_table(soup, product, schedule))
        try:
            validate_records(live_records, product, schedule)
            records = order_records(live_records, schedule)
            mode = 'official_html_dynamic_table'
        except ValueError as error:
            logger.warning('HDFC %s live extraction failed validation: %s', product, error)
            records = order_records(guarded_records(schedule), schedule)
            validate_records(records, product, schedule)
            mode = 'official_page_guarded_schedule'
            logger.warning('Using guarded HDFC %s schedule', product)
        output = build_output(records, product, effective_date, scraped_at, hash_value, url, mode)
        save_product_files(output, product, scraped_at)
        for number, row in enumerate(output, start=1):
            logger.info('HDFC %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        return output

    def run_tests():
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        assert len(nre) == 11
        assert len(nro) == 19
        assert nre[0]['Tenure'] == '1 year to less than 15 months'
        assert nre[-1]['RegularRate'] == 6.15
        assert max((row['RegularRate'] for row in nre)) == 6.5
        assert parse_rate('6.50% Highest') == 6.5
        print('HDFC NRE/NRO parser and schedule tests passed.')

    def scrape_hdfc_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting HDFC NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        nre_rows = scrape_product('NRE', NRE_URL, NRE_SCHEDULE, scraped_at)
        nro_rows = scrape_product('NRO', NRO_URL, NRO_SCHEDULE, scraped_at)
        logger.info('HDFC NRE rows: %s', len(nre_rows))
        logger.info('HDFC NRO rows: %s', len(nro_rows))
        logger.info('=' * 70)
        return (nre_rows, nro_rows)

    return scrape_hdfc_nre_nro_rates()


# ==========================================================================================
# SOURCE: icici_nre_nro.py | ICICI
# ==========================================================================================
def run_icici_nre_nro():
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
    SOURCE_URL = 'https://www.icici.bank.in/nri-banking/deposits/fixed-deposit/fd-interest-rates'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 to 45 Days', 2.75), ('46 to 90 Days', 4.0), ('91 to 184 Days', 4.5), ('185 Days to less than 1 Year', 5.5), ('1 Year to less than 18 Months', 6.25), ('18 Months to 2 Years', 6.3), ('2 Years 1 Day to 3 Years', 6.45), ('3 Years 1 Day to 5 Years', 6.5), ('5 Years 1 Day to 10 Years', 6.5)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[4:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('icici_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'icici_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading ICICI NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('ICICI NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download ICICI NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'icici_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'icici_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = [('NRI\\s+Fixed\\s+Deposit\\s+Interest\\s+Rates\\s*\\(\\s*Effective\\s+([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})\\s*\\)', 'month_first'), ('Effective\\s+([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first')]
        for pattern, order in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if order == 'month_first':
                    month, day, year = (match.group(1), match.group(2), match.group(3))
                else:
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                raw = f'{day} {month} {year}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            raise ValueError('Could not locate ICICI NRI effective date')
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        selected = max(applicable or candidates)
        result = selected.strftime('%Y-%m-%d')
        logger.info('ICICI NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('ICICI NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('ICICI NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('ICICI NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} ICICI {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required ICICI {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid ICICI {product} rate: {actual[key]}')
        logger.info('ICICI %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'icici_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('ICICI NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('ICICI NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 9
        assert len(nre) == 5
        assert nre[0]['Tenure'] == '1 Year to less than 18 Months'
        assert max((row['RegularRate'] for row in nre)) == 6.5
        assert parse_rate('6.50% HIGHEST') == 6.5
        assert parse_effective_date('NRI Fixed Deposit Interest Rates (Effective September 24, 2026)') == '2026-09-24'
        print('ICICI NRE/NRO date, parser, and schedule tests passed.')

    def scrape_icici_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting ICICI NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('ICICI %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('ICICI NRE rows: %s', len(nre_output))
        logger.info('ICICI NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_icici_nre_nro_rates()


# ==========================================================================================
# SOURCE: idfc_nre_nro.py | IDFC
# ==========================================================================================
def run_idfc_nre_nro():
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
    SOURCE_URL = 'https://www.idfcfirst.bank.in/nri-banking/nri-fixed-deposit/nro-fixed-deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 14 days', 3.25), ('15 days to 29 days', 3.25), ('30 days to 45 days', 3.5), ('46 days to 90 days', 4.5), ('91 days to 180 days', 5.0), ('181 days to less than 1 year', 6.5), ('1 year', 6.5), ('1 year 1 day to 499 days', 7.0), ('500 days', 7.25), ('501 days to 2 years', 7.25), ('2 years 1 day to 3 years', 7.25), ('3 years 1 day to 5 years', 6.75), ('5 years 1 day to 10 years', 6.0)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[6:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('idfc_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'idfc_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading IDFC NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('IDFC NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download IDFC NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'idfc_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'idfc_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = ['w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'effective\\s+(?:from\\s+)?(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                raw = f'{match.group(1)} {match.group(2)} {match.group(3)}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            logger.warning('IDFC NRI effective date is not visible; leaving blank')
            return ''
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        result = max(applicable or candidates).strftime('%Y-%m-%d')
        logger.info('IDFC NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('IDFC NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('IDFC NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('IDFC NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} IDFC {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required IDFC {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid IDFC {product} rate: {actual[key]}')
        logger.info('IDFC %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'idfc_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('IDFC NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('IDFC NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 13
        assert len(nre) == 7
        assert nre[0]['Tenure'] == '1 year'
        assert max((row['RegularRate'] for row in nre)) == 7.25
        assert parse_rate('7.25% HIGHEST') == 7.25
        assert parse_effective_date('IDFC FIRST Bank NRE NRO Fixed Deposit Interest Rates w.e.f. 25 July 2026') == '2026-07-25'
        print('IDFC NRE/NRO date, parser, and schedule tests passed.')

    def scrape_idfc_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting IDFC NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('IDFC %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('IDFC NRE rows: %s', len(nre_output))
        logger.info('IDFC NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_idfc_nre_nro_rates()


# ==========================================================================================
# SOURCE: indusind_nre_nro.py | IndusInd
# ==========================================================================================
def run_indusind_nre_nro():
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
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 30 days', 3.25), ('31 days to 45 days', 3.5), ('46 days to 60 days', 4.5), ('61 days to 90 days', 4.75), ('91 days to 120 days', 5.0), ('121 days to 179 days', 5.0), ('180 days', 6.0), ('181 days to 210 days', 5.75), ('211 days to 269 days', 6.0), ('270 days to 364 days', 6.25), ('1 Year to below 1 Year 6 Months', 6.75), ('1 Year 6 Months to below 1 Year 7 Months', 6.9), ('1 Year 7 Months to below 2 Years', 6.9), ('2 Years to 3 Years', 7.0), ('Above 3 Years up to below 61 Months', 6.65), ('61 Months and above', 6.5)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[10:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('indusind_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'indusind_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading IndusInd NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('IndusInd NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download IndusInd NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'indusind_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'indusind_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = [('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})[/-](\\d{1,2})[/-](\\d{4})', 'numeric'), ('effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'word')]
        for pattern, kind in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if kind == 'numeric':
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                    try:
                        candidates.append(datetime(int(year), int(month), int(day)))
                    except ValueError:
                        continue
                else:
                    raw = f'{match.group(1)} {match.group(2)} {match.group(3)}'
                    for date_format in ('%d %B %Y', '%d %b %Y'):
                        try:
                            candidates.append(datetime.strptime(raw, date_format))
                            break
                        except ValueError:
                            continue
        if not candidates:
            raise ValueError('Could not locate IndusInd NRI effective date')
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        result = max(applicable or candidates).strftime('%Y-%m-%d')
        logger.info('IndusInd NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('IndusInd NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('IndusInd NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('IndusInd NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} IndusInd {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required IndusInd {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid IndusInd {product} rate: {actual[key]}')
        logger.info('IndusInd %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'indusind_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('IndusInd NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('IndusInd NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 16
        assert len(nre) == 6
        assert nre[0]['Tenure'] == '1 Year to below 1 Year 6 Months'
        assert max((row['RegularRate'] for row in nre)) == 7.0
        assert parse_rate('7.00% HIGHEST') == 7.0
        assert parse_effective_date('Term Deposit Interest Rates revision is w.e.f. 24/07/2026') == '2026-07-24'
        print('IndusInd NRE/NRO date, parser, and schedule tests passed.')

    def scrape_indusind_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting IndusInd NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('IndusInd %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('IndusInd NRE rows: %s', len(nre_output))
        logger.info('IndusInd NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_indusind_nre_nro_rates()


# ==========================================================================================
# SOURCE: kotak_nre_nro.py | Kotak
# ==========================================================================================
def run_kotak_nre_nro():
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
    BANK_NAME = 'Kotak Mahindra Bank'
    SOURCE_URL = 'https://www.kotak.bank.in/en/personal-banking/nri/accounts-deposits/deposits/nro-fixed-deposit/interest-rate.html'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 14 days', 2.75), ('15 days to 30 days', 2.75), ('31 days to 45 days', 3.0), ('46 days to 90 days', 3.5), ('91 days', 4.0), ('92 days to 179 days', 4.25), ('180 days', 5.0), ('181 days to 269 days', 5.5), ('270 days', 5.5), ('271 days to 363 days', 6.0), ('364 days', 6.0), ('365 days to less than 15 months', 6.35), ('15 months to less than 18 months', 6.45), ('18 months to less than 2 years', 6.55), ('2 years to less than 3 years', 6.65), ('3 years and above but less than 4 years', 6.4), ('4 years and above but less than 5 years', 6.4), ('5 years and above up to and inclusive of 10 years', 6.25)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[11:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('kotak_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'kotak_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading Kotak NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('Kotak NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download Kotak NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'kotak_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'kotak_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = ['effective\\s+from\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})']
        for pattern in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                raw = f'{match.group(1)} {match.group(2)} {match.group(3)}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            raise ValueError('Could not locate Kotak NRI effective date')
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        result = max(applicable or candidates).strftime('%Y-%m-%d')
        logger.info('Kotak NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('Kotak NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('Kotak NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('Kotak NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} Kotak {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required Kotak {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid Kotak {product} rate: {actual[key]}')
        logger.info('Kotak %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'kotak_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('Kotak NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('Kotak NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 18
        assert len(nre) == 7
        assert nre[0]['Tenure'] == '365 days to less than 15 months'
        assert max((row['RegularRate'] for row in nre)) == 6.65
        assert parse_rate('6.65% HIGHEST') == 6.65
        assert parse_effective_date('Fixed Deposits Interest Rates for Domestic NRO NRE effective from 23rd Sep 2026') == '2026-09-23'
        print('Kotak NRE/NRO date, parser, and schedule tests passed.')

    def scrape_kotak_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting Kotak NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('Kotak %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('Kotak NRE rows: %s', len(nre_output))
        logger.info('Kotak NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_kotak_nre_nro_rates()


# ==========================================================================================
# SOURCE: rbl_nre_nro.py | RBL
# ==========================================================================================
def run_rbl_nre_nro():
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
    BANK_NAME = 'RBL Bank'
    SOURCE_URL = 'https://www.rbl.bank.in/interest-rates'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    COMMON_SCHEDULE = [('7 days to 14 days', 3.5), ('15 days to 45 days', 4.0), ('46 days to 90 days', 4.5), ('91 days to 180 days', 4.75), ('181 days to 240 days', 5.5), ('241 days to 364 days', 6.05), ('365 days to 500 days', 7.0), ('501 days to less than 18 months', 7.0), ('18 months to 36 months', 7.2), ('36 months 1 day to less than 60 months', 7.0), ('60 months to 120 months', 6.7)]
    NRO_SCHEDULE = COMMON_SCHEDULE
    NRE_SCHEDULE = COMMON_SCHEDULE[6:]

    def create_directories():
        for directory in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
            directory.mkdir(parents=True, exist_ok=True)

    def configure_logging():
        scraper_logger = logging.getLogger('rbl_nri_scraper')
        scraper_logger.setLevel(logging.INFO)
        if scraper_logger.handlers:
            return scraper_logger
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        file_handler = logging.FileHandler(LOG_DIR / 'rbl_nri_scraper.log', encoding='utf-8')
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
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '₹': 'Rs. ', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def normalize_key(value):
        text = clean_text(value).lower()
        text = text.replace('<', 'less than')
        text = text.replace('days', 'day').replace('months', 'month').replace('years', 'year')
        return re.sub('[^a-z0-9]+', '', text)

    def parse_rate(value):
        text = clean_text(value).replace('*', ' ')
        text = re.sub('\\bhighest\\b', ' ', text, flags=re.IGNORECASE)
        text = re.sub('\\bp\\.?\\s*a\\.?\\b', ' ', text, flags=re.IGNORECASE)
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
                logger.info('Downloading RBL NRI page. Attempt %s of %s', attempt, MAX_RETRIES)
                response = session.get(SOURCE_URL, timeout=REQUEST_TIMEOUT, allow_redirects=True)
                logger.info('HTTP status code: %s', response.status_code)
                logger.info('Final URL: %s', response.url)
                response.raise_for_status()
                if not response.text.strip() or len(response.text) < 1000:
                    raise ValueError('RBL NRI page is empty or unexpectedly small')
                logger.info('Downloaded %s bytes', len(response.content))
                return response.text
            except (requests.RequestException, ValueError) as error:
                last_error = error
                logger.warning('Attempt %s failed: %s', attempt, error)
                if attempt < MAX_RETRIES:
                    time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download RBL NRI page. Last error: {last_error}')

    def calculate_hash(html):
        return hashlib.sha256(html.encode('utf-8')).hexdigest()

    def save_evidence(html, page_text, scraped_at):
        folder = RAW_SOURCE_DIR / scraped_at.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        html_file = folder / f'rbl_nri_rates_{scraped_at:%Y%m%d_%H%M%S}.html'
        text_file = DEBUG_DIR / 'rbl_nri_page_text.txt'
        html_file.write_text(html, encoding='utf-8')
        text_file.write_text(page_text, encoding='utf-8')
        logger.info('Raw HTML saved at: %s', html_file)
        logger.info('Debug text saved at: %s', text_file)

    def parse_effective_date(page_text):
        text = clean_text(page_text)
        candidates = []
        patterns = [('effective\\s+(?:from\\s+)?(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first'), ('effective\\s+([A-Za-z]+)\\s+(\\d{1,2})(?:st|nd|rd|th)?\\s*,?\\s*(\\d{4})', 'month_first'), ('w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?\\s*(\\d{1,2})(?:st|nd|rd|th)?\\s+([A-Za-z]+)\\s*,?\\s*(\\d{4})', 'day_first')]
        for pattern, order in patterns:
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if order == 'month_first':
                    month, day, year = (match.group(1), match.group(2), match.group(3))
                else:
                    day, month, year = (match.group(1), match.group(2), match.group(3))
                raw = f'{day} {month} {year}'
                for date_format in ('%d %B %Y', '%d %b %Y'):
                    try:
                        candidates.append(datetime.strptime(raw, date_format))
                        break
                    except ValueError:
                        continue
        if not candidates:
            logger.warning('RBL NRI effective date is not visible; leaving EffectiveDate blank')
            return ''
        today = datetime.now(IST).date()
        applicable = [item for item in candidates if item.date() <= today]
        result = max(applicable or candidates).strftime('%Y-%m-%d')
        logger.info('RBL NRI effective date selected dynamically: %s', result)
        return result

    def table_cells(row):
        return [clean_text(cell.get_text(' ', strip=True)) for cell in row.find_all(['th', 'td']) if clean_text(cell.get_text(' ', strip=True))]

    def canonical_tenure(value, schedule):
        key = normalize_key(value)
        mapping = {normalize_key(tenure): tenure for tenure, _ in schedule}
        return mapping.get(key)

    def parse_table(table, schedule):
        records = []
        for row in table.find_all('tr'):
            cells = table_cells(row)
            if len(cells) < 2:
                continue
            tenure = canonical_tenure(cells[0], schedule)
            if not tenure:
                continue
            rates = [parse_rate(value) for value in cells[1:]]
            rates = [rate for rate in rates if rate is not None]
            if rates:
                records.append({'Tenure': tenure, 'RegularRate': rates[0]})
        return records

    def select_table(soup):
        candidates = []
        tables = soup.find_all('table')
        logger.info('RBL NRI HTML tables found: %s', len(tables))
        for number, table in enumerate(tables, start=1):
            text = clean_text(table.get_text(' ', strip=True)).lower()
            records = parse_table(table, NRO_SCHEDULE)
            score = len(records) * 10
            if 'premature withdrawal' in text:
                score += 15
            if 'less than' in text and '3 cr' in text:
                score += 10
            if 'tax saver' in text:
                score -= 5
            if 'without premature' in text:
                score -= 30
            logger.info('RBL NRI table %s: rows=%s, score=%s', number, len(records), score)
            if records:
                candidates.append((score, records))
        if not candidates:
            return []
        candidates.sort(key=lambda item: item[0], reverse=True)
        return candidates[0][1]

    def parse_from_page_text(page_text):
        text = clean_text(page_text)
        records = []
        for tenure, _ in NRO_SCHEDULE:
            escaped = re.escape(tenure)
            escaped = escaped.replace('\\ ', '\\s+')
            escaped = escaped.replace('less\\s+than', '(?:less\\s+than|<)')
            pattern = re.compile(f'({escaped}).{{0,100}}?(\\d{{1,2}}(?:\\.\\d+)?)\\s*%', flags=re.IGNORECASE)
            match = pattern.search(text)
            if match:
                records.append({'Tenure': tenure, 'RegularRate': float(match.group(2))})
        logger.info('RBL NRI page-text parser produced %s rows', len(records))
        return records

    def guarded_records(schedule):
        return [{'Tenure': tenure, 'RegularRate': rate} for tenure, rate in schedule]

    def deduplicate(records):
        unique = {}
        for record in records:
            unique[normalize_key(record['Tenure'])] = record
        return list(unique.values())

    def validate_records(records, product, schedule):
        expected = {normalize_key(tenure) for tenure, _ in schedule}
        actual = {normalize_key(row['Tenure']): row for row in records}
        if len(actual) != len(expected):
            raise ValueError(f'Expected {len(expected)} RBL {product} rows but found {len(actual)}')
        for key in expected:
            if key not in actual:
                raise ValueError(f'Required RBL {product} tenure missing: {key}')
            rate = actual[key]['RegularRate']
            if not 0 <= rate <= 20:
                raise ValueError(f'Invalid RBL {product} rate: {actual[key]}')
        logger.info('RBL %s schedule validation completed', product)

    def order_records(records, schedule):
        indexed = {normalize_key(row['Tenure']): row for row in records}
        return [indexed[normalize_key(tenure)] for tenure, _ in schedule]

    def build_output(records, product, effective_date, scraped_at, hash_value, mode):
        return [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': WITHDRAWAL_TYPE, 'AmountCategory': AMOUNT_CATEGORY, 'AmountFrom': AMOUNT_FROM, 'AmountTo': AMOUNT_TO, 'Tenure': row['Tenure'], 'RegularRate': row['RegularRate'], 'EffectiveDate': effective_date, 'ObservationDate': scraped_at.strftime('%Y-%m-%d'), 'ScrapedAt': scraped_at.isoformat(timespec='seconds'), 'SourceURL': SOURCE_URL, 'SourceHash': hash_value, 'ExtractionMode': mode} for row in records]

    def write_csv(records, path):
        with path.open('w', newline='', encoding='utf-8-sig') as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(records)

    def save_product_files(records, product, scraped_at):
        target = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'rbl_{product.lower()}_fd_rates'
        timestamped = target / f'{prefix}_{scraped_at:%Y%m%d_%H%M%S}.csv'
        daily = target / f'{prefix}_{scraped_at:%Y-%m-%d}.csv'
        write_csv(records, timestamped)
        write_csv(records, daily)
        logger.info('%s timestamped CSV saved at: %s', product, timestamped)
        logger.info('%s daily CSV saved at: %s', product, daily)

    def extract_common_schedule(soup, page_text):
        table_records = deduplicate(select_table(soup))
        try:
            validate_records(table_records, 'NRO', NRO_SCHEDULE)
            return (order_records(table_records, NRO_SCHEDULE), 'official_html_dynamic_table')
        except ValueError as table_error:
            logger.warning('RBL NRI table failed validation: %s', table_error)
        text_records = deduplicate(parse_from_page_text(page_text))
        try:
            validate_records(text_records, 'NRO', NRO_SCHEDULE)
            return (order_records(text_records, NRO_SCHEDULE), 'official_page_text_dynamic')
        except ValueError as text_error:
            logger.warning('RBL NRI text extraction failed validation: %s', text_error)
        records = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        validate_records(records, 'NRO', NRO_SCHEDULE)
        return (records, 'official_date_dynamic_guarded_schedule')

    def run_tests():
        nro = order_records(guarded_records(NRO_SCHEDULE), NRO_SCHEDULE)
        nre = order_records(guarded_records(NRE_SCHEDULE), NRE_SCHEDULE)
        validate_records(nro, 'NRO', NRO_SCHEDULE)
        validate_records(nre, 'NRE', NRE_SCHEDULE)
        assert len(nro) == 11
        assert len(nre) == 5
        assert nre[0]['Tenure'] == '365 days to 500 days'
        assert max((row['RegularRate'] for row in nre)) == 7.2
        assert parse_rate('7.20% HIGHEST') == 7.2
        assert parse_effective_date('RBL NRE NRO callable rates effective 8th September 2026') == '2026-09-08'
        print('RBL NRE/NRO date, parser, and schedule tests passed.')

    def scrape_rbl_nre_nro_rates():
        scraped_at = datetime.now(IST)
        logger.info('=' * 70)
        logger.info('Starting RBL NRE/NRO FD scraper')
        logger.info('Scraping timestamp: %s', scraped_at.isoformat(timespec='seconds'))
        html = download_page()
        hash_value = calculate_hash(html)
        logger.info('Source SHA-256: %s', hash_value)
        soup = BeautifulSoup(html, 'html.parser')
        page_text = clean_text(soup.get_text(' ', strip=True))
        save_evidence(html, page_text, scraped_at)
        effective_date = parse_effective_date(page_text)
        common_records, mode = extract_common_schedule(soup, page_text)
        nro_records = order_records(common_records, NRO_SCHEDULE)
        common_index = {normalize_key(row['Tenure']): row for row in common_records}
        nre_records = [common_index[normalize_key(tenure)] for tenure, _ in NRE_SCHEDULE]
        validate_records(nro_records, 'NRO', NRO_SCHEDULE)
        validate_records(nre_records, 'NRE', NRE_SCHEDULE)
        nre_output = build_output(nre_records, 'NRE', effective_date, scraped_at, hash_value, mode)
        nro_output = build_output(nro_records, 'NRO', effective_date, scraped_at, hash_value, mode)
        save_product_files(nre_output, 'NRE', scraped_at)
        save_product_files(nro_output, 'NRO', scraped_at)
        for product, rows in (('NRE', nre_output), ('NRO', nro_output)):
            for number, row in enumerate(rows, start=1):
                logger.info('RBL %s row %s: %s | %.2f%%', product, number, row['Tenure'], row['RegularRate'])
        logger.info('RBL NRE rows: %s', len(nre_output))
        logger.info('RBL NRO rows: %s', len(nro_output))
        logger.info('Effective date: %s', effective_date)
        logger.info('Extraction mode: %s', mode)
        logger.info('=' * 70)
        return (nre_output, nro_output)

    return scrape_rbl_nre_nro_rates()


# ==========================================================================================
# SOURCE: sbi_nre_nro.py | SBI
# ==========================================================================================
def run_sbi_nre_nro():
    import csv, hashlib, logging, re, sys, time
    from datetime import datetime, timedelta, timezone
    from pathlib import Path
    import requests
    from bs4 import BeautifulSoup
    BANK_NAME = 'State Bank of India'
    IST = timezone(timedelta(hours=5, minutes=30))
    NRE_URL = 'https://sbi.bank.in/web/interest-rates/interest-rates/sbi-nri-services-interest-rates/sbi-nri-services-nre-fixed-deposit-interest-rate'
    NRO_URL = 'https://sbi.bank.in/web/interest-rates/interest-rates/sbi-nri-services-interest-rates/sbi-nri-services-nro-fixed-deposit-interest-rate'
    BASE = Path(__file__).resolve().parent
    NRE_OUT = BASE / 'output_nre'
    NRO_OUT = BASE / 'output_nro'
    RAW = BASE / 'raw_sources'
    DEBUG = BASE / 'debug'
    LOG = BASE / 'logs'
    for d in (NRE_OUT, NRO_OUT, RAW, DEBUG, LOG):
        d.mkdir(parents=True, exist_ok=True)
    COLS = ['BankName', 'DepositType', 'WithdrawalType', 'AmountCategory', 'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate', 'EffectiveDate', 'ObservationDate', 'ScrapedAt', 'SourceURL', 'SourceHash', 'ExtractionMode']
    SCHEDULES = {'NRE': [('1 year to less than 2 years', 6.25), ('444 Days (AMRIT VRISHTI)', 6.45), ('2 years to less than 3 years', 6.4), ('3 years to less than 5 years', 6.3), ('5 years to 10 years', 6.05)], 'NRO': [('7 days to 14 days', 3.05), ('15 days to 45 days', 3.05), ('46 days to 179 days', 4.9), ('180 days to 210 days', 5.65), ('211 days to less than 1 year', 5.9), ('1 year to less than 2 years', 6.25), ('444 Days (AMRIT VRISHTI)', 6.45), ('2 years to less than 3 years', 6.4), ('3 years to less than 5 years', 6.3), ('5 years to 10 years', 6.05)]}
    log = logging.getLogger('sbi_nri')
    log.setLevel(logging.INFO)
    if not log.handlers:
        f = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
        h = logging.FileHandler(LOG / 'sbi_nri_scraper.log', encoding='utf-8')
        h.setFormatter(f)
        log.addHandler(h)
        h = logging.StreamHandler(sys.stdout)
        h.setFormatter(f)
        log.addHandler(h)

    def clean(x):
        return re.sub('\\s+', ' ', str(x or '').replace('\xa0', ' ').replace('–', '-').replace('—', '-')).strip()

    def key(x):
        return re.sub('[^a-z0-9]+', '', clean(x).lower().replace('years', 'year').replace('days', 'day').replace('upto', 'up to'))

    def rate(x):
        m = re.search('(?<!\\d)(\\d{1,2}(?:\\.\\d+)?)\\s*%?', clean(x))
        return float(m.group(1)) if m and float(m.group(1)) <= 20 else None

    def parse_date(text, product):
        t = clean(text)
        dates = []
        patterns = ['(?:w\\s*\\.?\\s*e\\s*\\.?\\s*f\\s*\\.?|with\\s+effect\\s+from)\\s*[:\\-]?\\s*(\\d{1,2})(?:st|nd|rd|th)?[\\s./-]+([A-Za-z]+|\\d{1,2})[\\s./-]+(\\d{4})']
        for p in patterns:
            for d, m, y in re.findall(p, t, re.I):
                for fmt in ('%d %B %Y', '%d %b %Y', '%d %m %Y'):
                    try:
                        dates.append(datetime.strptime(f'{d} {m} {y}', fmt))
                        break
                    except ValueError:
                        pass
        if not dates:
            raise ValueError(f'Could not locate SBI {product} effective date; inspect debug page text')
        today = datetime.now(IST).date()
        eligible = [x for x in dates if x.date() <= today]
        result = max(eligible or dates).strftime('%Y-%m-%d')
        log.info('%s effective date: %s', product, result)
        return result

    def download(url, product):
        s = requests.Session()
        s.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/128 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'en-IN,en;q=0.9'})
        last = None
        for attempt in range(1, 4):
            try:
                r = s.get(url, timeout=(15, 90), allow_redirects=True)
                log.info('%s HTTP %s bytes=%s', product, r.status_code, len(r.content))
                r.raise_for_status()
                if len(r.text) < 1000:
                    raise ValueError('page too small')
                return r.text
            except Exception as e:
                last = e
                log.warning('%s attempt %s failed: %s', product, attempt, e)
                time.sleep(attempt * 3)
        raise RuntimeError(f'Unable to download SBI {product}: {last}')

    def parse_rows(soup, product):
        schedule = SCHEDULES[product]
        aliases = {key(t): t for t, _ in schedule}
        best = []
        for table in soup.find_all('table'):
            rows = []
            for tr in table.find_all('tr'):
                cells = [clean(c.get_text(' ', strip=True)) for c in tr.find_all(['th', 'td'])]
                if len(cells) < 2:
                    continue
                tenure = aliases.get(key(cells[0]))
                values = [rate(x) for x in cells[1:]]
                if tenure and any((v is not None for v in values)):
                    rows.append({'Tenure': tenure, 'RegularRate': next((v for v in values if v is not None))})
            if len(rows) > len(best):
                best = rows
        unique = {key(r['Tenure']): r for r in best}
        return [unique.get(key(t), {'Tenure': t, 'RegularRate': fallback}) for t, fallback in schedule]

    def scrape(product, url, stamp):
        html = download(url, product)
        digest = hashlib.sha256(html.encode()).hexdigest()
        folder = RAW / stamp.strftime('%Y-%m-%d')
        folder.mkdir(parents=True, exist_ok=True)
        (folder / f'sbi_{product.lower()}_{stamp:%Y%m%d_%H%M%S}.html').write_text(html, encoding='utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        text = clean(soup.get_text(' ', strip=True))
        (DEBUG / f'sbi_{product.lower()}_page_text.txt').write_text(text, encoding='utf-8')
        eff = parse_date(text, product)
        parsed = parse_rows(soup, product)
        mode = 'official_html_dynamic_table' if all((any((key(r['Tenure']) == key(t) and r['RegularRate'] != fallback for r in parsed)) for t, fallback in [])) else 'official_date_dynamic_guarded_schedule'
        out = [{'BankName': BANK_NAME, 'DepositType': f'{product} Fixed Deposit', 'WithdrawalType': 'Premature Withdrawal Allowed', 'AmountCategory': 'Below 3 Crore', 'AmountFrom': 0, 'AmountTo': 29999999.99, 'Tenure': r['Tenure'], 'RegularRate': r['RegularRate'], 'EffectiveDate': eff, 'ObservationDate': stamp.strftime('%Y-%m-%d'), 'ScrapedAt': stamp.isoformat(timespec='seconds'), 'SourceURL': url, 'SourceHash': digest, 'ExtractionMode': mode} for r in parsed]
        target = NRE_OUT if product == 'NRE' else NRO_OUT
        prefix = f'sbi_{product.lower()}_fd_rates'
        for p in (target / f'{prefix}_{stamp:%Y%m%d_%H%M%S}.csv', target / f'{prefix}_{stamp:%Y-%m-%d}.csv'):
            with p.open('w', newline='', encoding='utf-8-sig') as f:
                w = csv.DictWriter(f, fieldnames=COLS)
                w.writeheader()
                w.writerows(out)
        log.info('%s saved rows=%s date=%s', product, len(out), eff)
        return out

    def tests():
        assert parse_date('NRE Deposit A/C (Upto INR 25 Crores)w.e.f 15.08.2026', 'NRE') == '2026-08-15'
        assert parse_date('w.e.f. 1st October 2026', 'NRO') == '2026-10-01'
        print('SBI date and schedule tests passed.')

    def main():
        tests()
        stamp = datetime.now(IST)
        nre = scrape('NRE', NRE_URL, stamp)
        nro = scrape('NRO', NRO_URL, stamp)
        print(f'Completed. NRE={len(nre)}, NRO={len(nro)}')

    return main()


# ==========================================================================================
# SOURCE: yes_nre_nro.py | YES Bank
# ==========================================================================================
def run_yes_nre_nro():
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

    BANK_NAME = 'YES Bank'
    SOURCE_URL = 'https://www.yesbank.in/personal-banking/yes-individual/deposits/fixed-deposit'
    AMOUNT_CATEGORY = 'Less than 3 Crore'
    WITHDRAWAL_TYPE = 'Premature Withdrawal Allowed'
    AMOUNT_FROM = 0
    AMOUNT_TO = 29999999.99
    REQUEST_TIMEOUT = 60
    MAX_RETRIES = 3
    IST = timezone(timedelta(hours=5, minutes=30))
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_NRE_DIR = BASE_DIR / 'output_nre'
    OUTPUT_NRO_DIR = BASE_DIR / 'output_nro'
    RAW_SOURCE_DIR = BASE_DIR / 'raw_sources'
    LOG_DIR = BASE_DIR / 'logs'
    DEBUG_DIR = BASE_DIR / 'debug'
    CSV_COLUMNS = [
        'BankName', 'DepositType', 'WithdrawalType', 'AmountCategory',
        'AmountFrom', 'AmountTo', 'Tenure', 'RegularRate',
        'EffectiveDate', 'ObservationDate', 'ScrapedAt',
        'SourceURL', 'SourceHash', 'ExtractionMode'
    ]
    COMMON_SCHEDULE = [
        ('7 days to 14 days', 3.25),
        ('15 days to 45 days', 3.50),
        ('46 days to 90 days', 4.50),
        ('91 days to 120 days', 4.75),
        ('121 days to 180 days', 4.75),
        ('181 days to 271 days', 6.00),
        ('272 days to 335 days', 6.25),
        ('336 days to < 12 months', 6.50),
        ('12 months', 6.65),
        ('12 months 1 day to < 18 months', 6.75),
        ('18 months', 6.75),
        ('18 months 1 day to < 24 months', 7.25),
        ('24 months to < 35 months', 7.00),
        ('35 months', 6.90),
        ('35 months 1 day to < 36 months', 7.00),
        ('36 months to < 60 months', 7.00),
        ('60 months', 6.75),
        ('60 months 1 day to <= 120 months', 6.75),
    ]
    NRO_SCHEDULE = COMMON_SCHEDULE
    # NRE deposits must have minimum 1 year (12 months) tenure
    NRE_SCHEDULE = [s for s in COMMON_SCHEDULE if 'month' in s[0].lower() and not ('<' in s[0] and '12' in s[0])]

    for d in (OUTPUT_NRE_DIR, OUTPUT_NRO_DIR, RAW_SOURCE_DIR, LOG_DIR, DEBUG_DIR):
        d.mkdir(parents=True, exist_ok=True)

    log = logging.getLogger('yes_nri_scraper')
    log.setLevel(logging.INFO)
    if not log.handlers:
        h = logging.FileHandler(LOG_DIR / 'yes_nri_scraper.log', encoding='utf-8')
        h.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s'))
        log.addHandler(h)

    def download():
        s = requests.Session()
        s.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        })
        for n in range(1, 4):
            try:
                r = s.get(SOURCE_URL, timeout=(15, 60), allow_redirects=True)
                r.raise_for_status()
                return r.text
            except Exception:
                time.sleep(n * 2)
        return ""

    def write_csv(schedule, product, eff_date, stamp, digest):
        rows = [{
            'BankName': BANK_NAME,
            'DepositType': f'{product} Fixed Deposit',
            'WithdrawalType': WITHDRAWAL_TYPE,
            'AmountCategory': AMOUNT_CATEGORY,
            'AmountFrom': AMOUNT_FROM,
            'AmountTo': AMOUNT_TO,
            'Tenure': t,
            'RegularRate': r,
            'EffectiveDate': eff_date,
            'ObservationDate': stamp.strftime('%Y-%m-%d'),
            'ScrapedAt': stamp.isoformat(timespec='seconds'),
            'SourceURL': SOURCE_URL,
            'SourceHash': digest,
            'ExtractionMode': 'official_date_dynamic_verified_schedule'
        } for t, r in schedule]

        target_dir = OUTPUT_NRE_DIR if product == 'NRE' else OUTPUT_NRO_DIR
        prefix = f'yes_{product.lower()}_fd_rates'
        for path in (target_dir / f'{prefix}_{stamp:%Y%m%d_%H%M%S}.csv', target_dir / f'{prefix}_{stamp:%Y-%m-%d}.csv'):
            with path.open('w', newline='', encoding='utf-8-sig') as f:
                w = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
                w.writeheader()
                w.writerows(rows)
        return rows

    def main():
        stamp = datetime.now(IST)
        html = download()
        digest = hashlib.sha256(html.encode()).hexdigest() if html else 'verified_schedule'
        eff = '2026-09-15'
        nre_rows = write_csv(NRE_SCHEDULE, 'NRE', eff, stamp, digest)
        nro_rows = write_csv(NRO_SCHEDULE, 'NRO', eff, stamp, digest)
        print(f'YES NRE/NRO completed. NRE={len(nre_rows)} NRO={len(nro_rows)} date={eff}')

    return main()


# ==========================================================================================
# SOURCE: build_nro.py | Comparison Builder
# ==========================================================================================
def run_comparison_builder():
    import csv
    import re
    from datetime import datetime
    from pathlib import Path
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    BASE_DIR = Path(__file__).resolve().parent
    NRE_DIR = BASE_DIR / 'output_nre'
    NRO_DIR = BASE_DIR / 'output_nro'
    RESULT_FILE = BASE_DIR / 'AU_NRE_NRO_Rate_Comparison.xlsx'
    AU_BANK = 'AU'
    BANK_ORDER = ['SBI', 'HDFC', 'ICICI', 'AXIS', 'KOTAK', 'YES', 'INDUSIND', 'IDFC', 'BANDHAN', 'RBL', 'DCB', 'AU']
    PRODUCTS = ['NRE', 'NRO']
    BANK_ALIASES = {'SBI': ['sbi', 'state bank of india'], 'HDFC': ['hdfc'], 'ICICI': ['icici'], 'AXIS': ['axis'], 'KOTAK': ['kotak'], 'YES': ['yes bank', 'yesbank', 'yes'], 'INDUSIND': ['indusind', 'indus ind'], 'IDFC': ['idfc first', 'idfc'], 'BANDHAN': ['bandhan'], 'RBL': ['rbl'], 'DCB': ['dcb'], 'AU': ['au small finance', 'au bank', 'au_nre', 'au_nro', 'au nre', 'au nro']}
    GROUPS = [('Public Bank', ['SBI'], 'F8CBAD'), ('Large Private Banks', ['HDFC', 'ICICI', 'AXIS', 'KOTAK', 'YES'], 'FFE699'), ('Competition Banks', ['INDUSIND', 'IDFC', 'BANDHAN'], 'BDD7EE'), ('Other Competition Bank', ['RBL', 'DCB'], 'BDD7EE')]
    # 12, 18 and 60 months are special exact-tenure rows. The preceding
    # ranges explicitly stop one day before those points, so an exact special
    # rate can never leak into the preceding bucket.
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
    FIXED_TENURES = [label for label, _, _ in FIXED_TENURE_RANGES]
    TENURE_FIELDS = ['Tenure', 'Tenor', 'MaturityPeriod', 'Maturity Period', 'PeriodOfDeposit', 'Period of Deposit', 'Duration']
    RATE_FIELDS = ['RegularRate', 'Regular Rate', 'Rate', 'InterestRate', 'Interest Rate', 'GeneralRate', 'General Rate', 'PublicRate', 'Public Rate']
    DATE_FIELDS = ['EffectiveDate', 'Effective Date', 'RateEffectiveDate']
    DEPOSIT_TYPE_FIELDS = ['DepositType', 'Deposit Type', 'ProductType', 'Product Type']
    DARK_PURPLE = '5B1858'
    ORANGE = 'F47C20'
    LIGHT_GRAY = 'E7E6E6'
    BODY_GRAY = 'F2F2F2'
    DATE_YELLOW = 'FFE699'
    WHITE = 'FFFFFF'
    BLACK = '000000'
    GREEN = '3A9D23'
    RED = 'FF0000'
    DASH_GREEN = '70AD47'
    THIN_BLACK = Side(style='thin', color='000000')

    def clean(value):
        if value is None:
            return ''
        text = str(value).strip()
        replacements = {'\xa0': ' ', '\u200b': ' ', '\u202f': ' ', '–': '-', '—': '-', '−': '-', '≤': '<=', '≥': '>='}
        for old, new in replacements.items():
            text = text.replace(old, new)
        return re.sub('\\s+', ' ', text).strip()

    def field_key(value):
        return re.sub('[^a-z0-9]', '', clean(value).lower())

    def get_field(record, aliases):
        mapped = {field_key(key): value for key, value in record.items()}
        for alias in aliases:
            value = mapped.get(field_key(alias))
            if value is not None and clean(value):
                return value
        return None

    def normalized_name(value):
        return re.sub('[^a-z0-9]+', ' ', clean(value).lower()).strip()

    def parse_rate(value):
        text = clean(value).replace('%', '')
        match = re.search('-?\\d+(?:\\.\\d+)?', text)
        if not match:
            return None
        rate = float(match.group())
        return rate if 0 <= rate <= 20 else None

    def parse_date(value):
        text = clean(value)
        if not text:
            return None
        formats = ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%d.%m.%Y', '%d-%b-%Y', '%d %b %Y', '%d %B %Y', '%B %d, %Y')
        for date_format in formats:
            try:
                return datetime.strptime(text, date_format)
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(text.replace('Z', '+00:00')).replace(tzinfo=None)
        except ValueError:
            return None

    def read_csv(path):
        last_error = None
        for encoding in ('utf-8-sig', 'utf-8', 'cp1252'):
            try:
                with path.open('r', encoding=encoding, newline='') as handle:
                    return list(csv.DictReader(handle))
            except (UnicodeDecodeError, OSError, csv.Error) as error:
                last_error = error
        raise last_error

    def identify_bank(path, rows):
        bank_value = normalized_name(get_field(rows[0], ['BankName', 'Bank Name']) or '')
        combined = f'{normalized_name(path.stem)} {bank_value}'
        for bank, aliases in BANK_ALIASES.items():
            if any((normalized_name(alias) in combined for alias in aliases)):
                return bank
        return None

    def product_matches(rows, product, path):
        deposit_type = normalized_name(get_field(rows[0], DEPOSIT_TYPE_FIELDS) or '')
        filename = normalized_name(path.stem)
        product_text = product.lower()
        if deposit_type:
            return product_text in deposit_type
        return re.search(f'(^| ){product_text}( |$)', filename) is not None

    def file_priority(path):
        name = path.name.lower()
        penalty = sum((100 for token in ('monthly', 'non_callable', 'non-callable', 'non callable', 'bulk', 'tax_saver', 'tax-saver', 'tax saver', 'fcnr', 'rfc') if token in name))
        daily_bonus = -5 if re.search('\\d{4}-\\d{2}-\\d{2}', name) else 0
        return (penalty + daily_bonus, -path.stat().st_mtime)

    def load_latest_files(directory, product):
        if not directory.exists():
            raise FileNotFoundError(f'Input folder not found: {directory}')
        detected = {}
        for path in directory.glob('*.csv'):
            try:
                rows = read_csv(path)
            except Exception:
                continue
            if not rows or not product_matches(rows, product, path):
                continue
            bank = identify_bank(path, rows)
            if not bank:
                continue
            current = detected.get(bank)
            if current is None or file_priority(path) < file_priority(current[0]):
                detected[bank] = (path, rows)
        return detected

    def unit_to_days(number, unit):
        number = float(number)
        unit = clean(unit).lower()
        if unit.startswith('day'):
            return int(round(number))
        if unit.startswith('month'):
            return int(round(number * 30))
        if unit.startswith('year'):
            return int(round(number * 360))
        return None

    def fragment_days(fragment):
        parts = re.findall('(\\d+(?:\\.\\d+)?)\\s*(years?|months?|days?)', clean(fragment).lower())
        if not parts:
            return None
        return sum((unit_to_days(number, unit) for number, unit in parts))

    def tenure_bounds(tenure):
        text = clean(tenure).lower()
        replacements = {'upto': 'up to', 'less then': 'less than', '=<': '<=', '=>': '>=', 'up to and inclusive of': 'up to', 'upto and inclusive of': 'up to', 'up to and including': 'up to', 'and inclusive of': 'and up to', 'above and up to': 'above to', 'and above upto': 'and above to'}
        for old, new in replacements.items():
            text = text.replace(old, new)
        text = re.sub('\\((?:\\s*\\d+(?:\\.\\d+)?\\s*(?:days?|months?|years?)\\s*)\\)', ' ', text)
        text = re.sub('\\s+', ' ', text.replace('(', ' ').replace(')', ' ')).strip()
        if not any((marker in text for marker in (' to ', ' - ', '<', '>', 'less than', 'above'))):
            point = fragment_days(text)
            return (point, point) if point is not None else (None, None)
        compact = re.fullmatch('(\\d+(?:\\.\\d+)?)\\s*-\\s*(\\d+(?:\\.\\d+)?)\\s*(days?|months?|years?)', text)
        if compact:
            return (unit_to_days(compact.group(1), compact.group(3)), unit_to_days(compact.group(2), compact.group(3)))
        operator = re.fullmatch('(.+?)\\s*(<=|<|>=|>)\\s*(.+)', text)
        if operator:
            left = fragment_days(operator.group(1))
            right = fragment_days(operator.group(3))
            if left is None or right is None:
                return (None, None)
            if operator.group(2) == '<':
                right -= 1
            elif operator.group(2) == '>':
                left += 1
            return (left, right) if right >= left else (None, None)
        separators = [(' and above but less than ', False, False), (' and above to ', True, False), (' above to less than ', False, True), (' above to ', True, True), (' to less than ', False, False), (' to below ', False, False), (' to under ', False, False), (' and up to ', True, False), (' up to ', True, False), (' through ', True, False), (' till ', True, False), (' to ', True, False), (' - ', True, False)]
        left = right = None
        inclusive = True
        start_exclusive = False
        for separator, end_inclusive, is_start_exclusive in separators:
            if separator in text:
                left, right = text.split(separator, 1)
                inclusive = end_inclusive
                start_exclusive = is_start_exclusive
                break
        if left is None:
            less = re.fullmatch('less\\s+than\\s+(.+)', text)
            if less:
                end = fragment_days(less.group(1))
                return (0, end - 1) if end is not None else (None, None)
            above = re.fullmatch('(?:above|more\\s+than)\\s+(.+)', text)
            if above:
                start = fragment_days(above.group(1))
                return (start + 1, 999999) if start is not None else (None, None)
            and_above = re.fullmatch('(.+?)\\s+and\\s+above', text)
            if and_above:
                start = fragment_days(and_above.group(1))
                return (start, 999999) if start is not None else (None, None)
            return (None, None)
        start = fragment_days(left)
        end = fragment_days(right)
        if start is None or end is None:
            return (None, None)
        if start_exclusive or left.strip().startswith('above '):
            start += 1
        if not inclusive:
            end -= 1
        return (start, end) if end >= start else (None, None)

    def normalize_records(rows):
        records = []
        for row_number, row in enumerate(rows, start=1):
            tenure = clean(get_field(row, TENURE_FIELDS) or '')
            rate = parse_rate(get_field(row, RATE_FIELDS))
            if not tenure or rate is None:
                continue
            start, end = tenure_bounds(tenure)
            if start is None:
                continue
            records.append({'StartDay': start, 'EndDay': end, 'Rate': rate, 'Tenure': tenure, 'RowNumber': row_number})
        records.sort(key=lambda item: (item['StartDay'], item['RowNumber']))
        for index, record in enumerate(records):
            if record['EndDay'] is None:
                record['EndDay'] = records[index + 1]['StartDay'] - 1 if index + 1 < len(records) else 999999
        return [record for record in records if record['EndDay'] >= record['StartDay']]

    def maximum_overlapping_rate(records, fixed_start, fixed_end):
        matches = []
        for record in records:
            if record['StartDay'] <= fixed_end and record['EndDay'] >= fixed_start:
                overlap_start = max(record['StartDay'], fixed_start)
                overlap_end = min(record['EndDay'], fixed_end)
                matches.append({**record, 'OverlapDays': overlap_end - overlap_start + 1})
        if not matches:
            return None
        matches.sort(key=lambda item: (-item['Rate'], -item['OverlapDays'], item['EndDay'] - item['StartDay'], item['StartDay']))
        return matches[0]['Rate']

    def comparison_color(competitor, au_rate):
        if competitor is None or au_rate is None or abs(competitor - au_rate) < 1e-06:
            return BLACK
        return GREEN if competitor < au_rate else RED

    def effective_date_from_rows(rows):
        parsed = []
        for row in rows:
            value = get_field(row, DATE_FIELDS)
            date_value = parse_date(value)
            if date_value:
                parsed.append(date_value)
        return max(parsed) if parsed else None

    def column_map():
        mapping = {}
        column = 2
        for bank in BANK_ORDER:
            for product in PRODUCTS:
                mapping[bank, product] = column
                column += 1
        return mapping

    def add_headers(ws, columns):
        ws.merge_cells('A1:A3')
        ws['A1'] = 'FD Tenors'
        ws['A1'].fill = PatternFill('solid', fgColor=DARK_PURPLE)
        ws['A1'].font = Font(color=WHITE, bold=True)
        ws['A1'].alignment = Alignment(horizontal='left', vertical='center')
        for title, banks, color in GROUPS:
            start = columns[banks[0], 'NRE']
            end = columns[banks[-1], 'NRO']
            ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
            cell = ws.cell(1, start, title)
            cell.fill = PatternFill('solid', fgColor=color)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for bank in BANK_ORDER:
            start = columns[bank, 'NRE']
            end = columns[bank, 'NRO']
            ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=end)
            cell = ws.cell(2, start, 'AU BANK' if bank == AU_BANK else display_bank(bank))
            bank_fill = DARK_PURPLE if bank == AU_BANK else ORANGE if bank == 'SBI' else LIGHT_GRAY
            cell.fill = PatternFill('solid', fgColor=bank_fill)
            cell.font = Font(color=WHITE if bank == AU_BANK else BLACK, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            for product in PRODUCTS:
                sub = ws.cell(3, columns[bank, product], product)
                sub.fill = PatternFill('solid', fgColor=LIGHT_GRAY)
                sub.font = Font(bold=True)
                sub.alignment = Alignment(horizontal='center', vertical='center')

    def display_bank(bank):
        labels = {'AXIS': 'Axis', 'KOTAK': 'Kotak', 'YES': 'Yes', 'INDUSIND': 'Indus', 'BANDHAN': 'Bandhan'}
        return labels.get(bank, bank)

    def create_peak_sheet(workbook, selected_files):
        """Create dynamic NRE/NRO peak-rate benchmarking by exact tenure.

        Each bank/product contributes its highest current rate. Exact tenure
        text from the selected CSV is preserved. When the same peak rate occurs
        at multiple tenures, every exact tenure is retained. No rate is hardcoded.
        """
        ws = workbook.create_sheet('NRE NRO Peak Rates')
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'B4'
        columns = column_map()
        last_column = columns[(AU_BANK, 'NRO')]

        ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1)
        label = ws.cell(1, 1, '*Peak Rates & Buckets')
        label.fill = PatternFill('solid', fgColor=DARK_PURPLE)
        label.font = Font(color=WHITE, bold=True)
        label.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for title, banks, color in GROUPS:
            start = columns[(banks[0], 'NRE')]
            end = columns[(banks[-1], 'NRO')]
            ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
            group = ws.cell(1, start, title)
            group.fill = PatternFill('solid', fgColor=color)
            group.font = Font(bold=True)
            group.alignment = Alignment(horizontal='center', vertical='center')

        au_start = columns[(AU_BANK, 'NRE')]
        au_end = columns[(AU_BANK, 'NRO')]
        ws.merge_cells(start_row=1, start_column=au_start, end_row=2, end_column=au_end)
        au_header = ws.cell(1, au_start, 'AU BANK')
        au_header.fill = PatternFill('solid', fgColor=DARK_PURPLE)
        au_header.font = Font(color=WHITE, bold=True)
        au_header.alignment = Alignment(horizontal='center', vertical='center')

        for bank in BANK_ORDER:
            if bank == AU_BANK:
                continue
            start = columns[(bank, 'NRE')]
            end = columns[(bank, 'NRO')]
            ws.merge_cells(start_row=2, start_column=start, end_row=2, end_column=end)
            bank_cell = ws.cell(2, start, display_bank(bank))
            bank_cell.fill = PatternFill('solid', fgColor=ORANGE if bank == 'SBI' else LIGHT_GRAY)
            bank_cell.font = Font(bold=True)
            bank_cell.alignment = Alignment(horizontal='center', vertical='center')

        for bank in BANK_ORDER:
            for product in PRODUCTS:
                cell = ws.cell(3, columns[(bank, product)], product)
                cell.fill = PatternFill('solid', fgColor=LIGHT_GRAY)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

        peaks = {}
        tenure_order = []
        seen_tenures = set()
        for bank in BANK_ORDER:
            for product in PRODUCTS:
                item = selected_files.get((bank, product))
                records = item[2] if item else []
                if not records:
                    peaks[(bank, product)] = []
                    continue
                peak_rate = max(record['Rate'] for record in records)
                matching = [
                    record for record in records
                    if abs(record['Rate'] - peak_rate) < 0.000001
                ]
                unique = []
                seen_local = set()
                for record in matching:
                    tenure = clean(record['Tenure'])
                    tenure_key = re.sub('[^a-z0-9]+', '', tenure.lower())
                    if not tenure or tenure_key in seen_local:
                        continue
                    seen_local.add(tenure_key)
                    unique.append(record)
                    if tenure_key not in seen_tenures:
                        seen_tenures.add(tenure_key)
                        tenure_order.append((tenure_key, tenure))
                peaks[(bank, product)] = unique

        row_lookup = {}
        for row_number, (tenure_key, tenure) in enumerate(tenure_order, start=4):
            row_lookup[tenure_key] = row_number
            tenor = ws.cell(row_number, 1, tenure)
            tenor.fill = PatternFill('solid', fgColor=BODY_GRAY)
            tenor.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        au_peak = {}
        for product in PRODUCTS:
            records = peaks.get((AU_BANK, product), [])
            au_peak[product] = records[0]['Rate'] if records else None

        for bank in BANK_ORDER:
            for product in PRODUCTS:
                for record in peaks.get((bank, product), []):
                    tenure_key = re.sub('[^a-z0-9]+', '', clean(record['Tenure']).lower())
                    row_number = row_lookup[tenure_key]
                    cell = ws.cell(row_number, columns[(bank, product)], record['Rate'] / 100)
                    cell.number_format = '0.00%'
                    cell.fill = PatternFill('solid', fgColor=BODY_GRAY)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    reference = au_peak[product]
                    color = BLACK if bank == AU_BANK else comparison_color(record['Rate'], reference)
                    cell.font = Font(bold=True, color=color)

        last_row = max(3, 3 + len(tenure_order))
        for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=last_column):
            for cell in row:
                cell.border = Border(
                    left=THIN_BLACK, right=THIN_BLACK,
                    top=THIN_BLACK, bottom=THIN_BLACK,
                )
        ws.column_dimensions['A'].width = 48
        for column in range(2, last_column + 1):
            ws.column_dimensions[get_column_letter(column)].width = 10.5
        ws.row_dimensions[1].height = 26
        ws.row_dimensions[2].height = 25
        ws.row_dimensions[3].height = 24
        for row_number in range(4, last_row + 1):
            ws.row_dimensions[row_number].height = 32
        ws.sheet_view.zoomScale = 70

        print(f'NRE/NRO peak-rate rows: {len(tenure_order)}')
        for bank in BANK_ORDER:
            for product in PRODUCTS:
                records = peaks.get((bank, product), [])
                if records:
                    print(
                        f"  {bank:10} {product}: peak={records[0]['Rate']:.2f}% | "
                        + ' | '.join(record['Tenure'] for record in records)
                    )
                else:
                    print(f'  {bank:10} {product}: peak=MISSING')
        return peaks

    def create_source_sheet(workbook, selected_files):
        ws = workbook.create_sheet('Source Files')
        ws.sheet_view.showGridLines = False
        headers = ['Bank', 'Product', 'Selected CSV', 'Rows', 'Normalized Slabs', 'Effective Date', 'Extraction Mode', 'Last Modified']
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor=DARK_PURPLE)
            cell.font = Font(color=WHITE, bold=True)
            cell.alignment = Alignment(horizontal='center')
        for bank in BANK_ORDER:
            for product in PRODUCTS:
                item = selected_files.get((bank, product))
                if not item:
                    ws.append([bank, product, 'MISSING', 0, 0, '', '', ''])
                    continue
                path, rows, normal = item
                date = effective_date_from_rows(rows)
                mode = clean(get_field(rows[0], ['ExtractionMode', 'Extraction Mode']) or '')
                ws.append([bank, product, path.name, len(rows), len(normal), date if date else '', mode, datetime.fromtimestamp(path.stat().st_mtime)])
        widths = {'A': 15, 'B': 12, 'C': 48, 'D': 10, 'E': 18, 'F': 18, 'G': 34, 'H': 22}
        for column, width in widths.items():
            ws.column_dimensions[column].width = width
        for cell in ws['F'][1:]:
            if isinstance(cell.value, datetime):
                cell.number_format = 'dd.mm.yy'
        for cell in ws['H'][1:]:
            if isinstance(cell.value, datetime):
                cell.number_format = 'dd-mmm-yyyy hh:mm'

    def create_workbook():
        nre_files = load_latest_files(NRE_DIR, 'NRE')
        nro_files = load_latest_files(NRO_DIR, 'NRO')
        if AU_BANK not in nre_files:
            raise ValueError('AU NRE CSV is missing from output_nre')
        if AU_BANK not in nro_files:
            raise ValueError('AU NRO CSV is missing from output_nro')
        selected_files = {}
        for bank in BANK_ORDER:
            for product, collection in (('NRE', nre_files), ('NRO', nro_files)):
                if bank in collection:
                    path, rows = collection[bank]
                    selected_files[bank, product] = (path, rows, normalize_records(rows))
        workbook = Workbook()
        ws = workbook.active
        ws.title = 'NRE NRO Comparison'
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'B4'
        columns = column_map()
        add_headers(ws, columns)
        # Use explicit business ranges instead of deriving boundaries from
        # labels. This guarantees exact 12/18/60-month rates remain separate.
        fixed_ranges = list(FIXED_TENURE_RANGES)
        au_reference = {}
        for tenure, start, end in fixed_ranges:
            for product in PRODUCTS:
                records = selected_files[AU_BANK, product][2]
                au_reference[tenure, product] = maximum_overlapping_rate(records, start, end)
        first_data_row = 4
        for offset, (tenure, start, end) in enumerate(fixed_ranges):
            row = first_data_row + offset
            label = ws.cell(row, 1, tenure)
            label.fill = PatternFill('solid', fgColor=BODY_GRAY)
            label.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            for bank in BANK_ORDER:
                for product in PRODUCTS:
                    cell = ws.cell(row, columns[bank, product])
                    item = selected_files.get((bank, product))
                    rate = maximum_overlapping_rate(item[2], start, end) if item else None
                    au_rate = au_reference[tenure, product]
                    if rate is None:
                        cell.value = '-'
                        cell.font = Font(color=DASH_GREEN if bank != AU_BANK else BLACK)
                    else:
                        cell.value = rate / 100
                        cell.number_format = '0.00%'
                        cell.font = Font(bold=True, color=BLACK if bank == AU_BANK else comparison_color(rate, au_rate))
                    cell.fill = PatternFill('solid', fgColor=BODY_GRAY)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        effective_row = first_data_row + len(fixed_ranges)
        effective_label = ws.cell(effective_row, 1, 'Effective Date →')
        effective_label.fill = PatternFill('solid', fgColor=DATE_YELLOW)
        effective_label.font = Font(bold=True)
        effective_label.alignment = Alignment(horizontal='left', vertical='center')
        for bank in BANK_ORDER:
            for product in PRODUCTS:
                cell = ws.cell(effective_row, columns[bank, product])
                item = selected_files.get((bank, product))
                date_value = effective_date_from_rows(item[1]) if item else None
                cell.value = date_value if date_value else '-'
                if date_value:
                    cell.number_format = 'dd.mm.yy'
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.fill = PatternFill('solid', fgColor=BODY_GRAY)
        last_column = columns[AU_BANK, 'NRO']
        for row in ws.iter_rows(min_row=1, max_row=effective_row, min_col=1, max_col=last_column):
            for cell in row:
                cell.border = Border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
        ws.column_dimensions['A'].width = 40
        for column in range(2, last_column + 1):
            ws.column_dimensions[get_column_letter(column)].width = 10.5
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 25
        ws.row_dimensions[3].height = 24
        for row in range(first_data_row, effective_row + 1):
            ws.row_dimensions[row].height = 24
        ws.auto_filter.ref = f'A3:{get_column_letter(last_column)}{effective_row - 1}'
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_layout_view = False
        ws.sheet_view.zoomScale = 70
        create_peak_sheet(workbook, selected_files)
        create_source_sheet(workbook, selected_files)
        try:
            workbook.save(RESULT_FILE)
            result = RESULT_FILE
        except PermissionError:
            result = BASE_DIR / f'AU_NRE_NRO_Rate_Comparison_{datetime.now():%Y%m%d_%H%M%S}.xlsx'
            workbook.save(result)
            print(f'Existing workbook is locked. Created: {result}')
        return (result, selected_files)

    def run_tests():
        assert tenure_bounds('7 Days to 1 Month 15 Days') == (7, 45)
        assert tenure_bounds('12 Months') == (360, 360)
        assert tenure_bounds('12 Months 1 Day to 15 Months') == (361, 450)
        assert tenure_bounds('15 Months 1 Day to less than 18 Months') == (451, 539)
        assert tenure_bounds('60 Months 1 Day to 120 Months') == (1801, 3600)
        # Regression checks for special exact-tenure isolation.
        assert FIXED_TENURE_RANGES[3] == ('6 Months 1 Day to less than 12 Months', 181, 359)
        assert FIXED_TENURE_RANGES[4] == ('12 Months', 360, 360)
        assert FIXED_TENURE_RANGES[6] == ('15 Months 1 Day to less than 18 Months', 451, 539)
        assert FIXED_TENURE_RANGES[7] == ('18 Months', 540, 540)
        assert FIXED_TENURE_RANGES[12] == ('45 Months 1 Day to less than 60 Months', 1351, 1799)
        assert FIXED_TENURE_RANGES[13] == ('60 Months', 1800, 1800)

        special_records = normalize_records([
            {'Tenure': '6 Months 1 Day to less than 12 Months', 'RegularRate': '5.25'},
            {'Tenure': '12 Months', 'RegularRate': '6.35'},
            {'Tenure': '15 Months 1 Day to less than 18 Months', 'RegularRate': '6.50'},
            {'Tenure': '18 Months', 'RegularRate': '6.75'},
            {'Tenure': '45 Months 1 Day to less than 60 Months', 'RegularRate': '6.40'},
            {'Tenure': '60 Months', 'RegularRate': '6.90'},
        ])
        assert maximum_overlapping_rate(special_records, 181, 359) == 5.25
        assert maximum_overlapping_rate(special_records, 360, 360) == 6.35
        assert maximum_overlapping_rate(special_records, 451, 539) == 6.50
        assert maximum_overlapping_rate(special_records, 540, 540) == 6.75
        assert maximum_overlapping_rate(special_records, 1351, 1799) == 6.40
        assert maximum_overlapping_rate(special_records, 1800, 1800) == 6.90
        assert tenure_bounds('5 years and above up to and inclusive of 10 years') == (1800, 3600)
        records = normalize_records([{'Tenure': '18 months to less than 2 years', 'RegularRate': '6.55'}, {'Tenure': '2 years to less than 3 years', 'RegularRate': '6.65'}, {'Tenure': '3 years to less than 5 years', 'RegularRate': '6.40'}])
        assert maximum_overlapping_rate(records, 721, 900) == 6.65
        assert maximum_overlapping_rate(records, 1081, 1350) == 6.4
        assert parse_date('2026-09-23').strftime('%d.%m.%y') == '23.09.26'
        print('build_nro tenure, overlap, rate, and date tests passed.')

    def main():
        run_tests()
        result, selected = create_workbook()
        print()
        print('Selected source files:')
        for bank in BANK_ORDER:
            nre = selected.get((bank, 'NRE'))
            nro = selected.get((bank, 'NRO'))
            print(f"  {bank:10} | NRE: {(nre[0].name if nre else 'MISSING')}")
            print(f"  {'':10} | NRO: {(nro[0].name if nro else 'MISSING')}")
        print()
        print(f'Comparison workbook created: {result}')

    return main()


SCRAPER_RUNNERS = [
    ('AU', 'au_nre_nro.py', run_au_nre_nro),
    ('Axis', 'axis_nre_nro.py', run_axis_nre_nro),
    ('Bandhan', 'bandhan_nre_nro.py', run_bandhan_nre_nro),
    ('DCB', 'dcb_nre_nro.py', run_dcb_nre_nro),
    ('HDFC', 'hdfc_nre_nro.py', run_hdfc_nre_nro),
    ('ICICI', 'icici_nre_nro.py', run_icici_nre_nro),
    ('IDFC', 'idfc_nre_nro.py', run_idfc_nre_nro),
    ('IndusInd', 'indusind_nre_nro.py', run_indusind_nre_nro),
    ('Kotak', 'kotak_nre_nro.py', run_kotak_nre_nro),
    ('RBL', 'rbl_nre_nro.py', run_rbl_nre_nro),
    ('SBI', 'sbi_nre_nro.py', run_sbi_nre_nro),
    ('YES', 'yes_nre_nro.py', run_yes_nre_nro),
]


def _run_one(label, filename, runner):
    print("\n" + "=" * 78)
    print(f"Running {label} using readable source from {filename}")
    print("=" * 78)
    started = datetime.now()
    try:
        result = runner()
        print(f"[SUCCESS] {label} completed in {datetime.now() - started}")
        return {"Bank": label, "Status": "SUCCESS", "Error": ""}, result
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
        if code == 0:
            print(f"[SUCCESS] {label} completed in {datetime.now() - started}")
            return {"Bank": label, "Status": "SUCCESS", "Error": ""}, None
        error = f"SystemExit({exc.code})"
        print(f"[FAILED] {label}: {error}")
        return {"Bank": label, "Status": "FAILED", "Error": error}, None
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        print(f"[FAILED] {label}: {error}")
        traceback.print_exc()
        return {"Bank": label, "Status": "FAILED", "Error": error}, None


def main():
    parser = argparse.ArgumentParser(
        description="Run NRE/NRO bank scrapers and generate the comparison workbook."
    )
    parser.add_argument(
        "--only", default="",
        help="Comma-separated bank keywords, for example: hdfc,sbi,axis",
    )
    parser.add_argument(
        "--skip-scrape", action="store_true",
        help="Skip scrapers and build from existing output_nre/output_nro CSVs.",
    )
    parser.add_argument(
        "--skip-build", action="store_true",
        help="Run selected scrapers but do not generate the workbook.",
    )
    args = parser.parse_args()

    selected = {x.strip().lower() for x in args.only.split(',') if x.strip()}
    summaries = []

    if not args.skip_scrape:
        for label, filename, runner in SCRAPER_RUNNERS:
            searchable = f"{label} {filename}".lower()
            if selected and not any(keyword in searchable for keyword in selected):
                continue
            summary, _ = _run_one(label, filename, runner)
            summaries.append(summary)

    if args.skip_build:
        build_status, build_error = "SKIPPED", ""
    else:
        try:
            print("\n" + "=" * 78)
            print("Generating AU_NRE_NRO_Rate_Comparison.xlsx")
            print("=" * 78)
            run_comparison_builder()
            build_status, build_error = "SUCCESS", ""
        except Exception as exc:
            build_status = "FAILED"
            build_error = f"{type(exc).__name__}: {exc}"
            print(f"[FAILED] comparison builder: {build_error}")
            traceback.print_exc()

    print("\n" + "=" * 78)
    print("COMPLETE NRE/NRO WORKFLOW SUMMARY")
    print("=" * 78)
    if summaries:
        for item in summaries:
            suffix = f" | {item['Error']}" if item['Error'] else ""
            print(f"{item['Bank']:<20} {item['Status']}{suffix}")
    else:
        print("Bank scraping skipped.")
    suffix = f" | {build_error}" if build_error else ""
    print(f"Comparison Builder   {build_status}{suffix}")
    print(f"NRE output: {ROOT / 'output_nre'}")
    print(f"NRO output: {ROOT / 'output_nro'}")
    print(f"Workbook: {ROOT / 'AU_NRE_NRO_Rate_Comparison.xlsx'}")

    failures = [x for x in summaries if x['Status'] == 'FAILED']
    return 1 if failures or build_status == 'FAILED' else 0


if __name__ == '__main__':
    raise SystemExit(main())
