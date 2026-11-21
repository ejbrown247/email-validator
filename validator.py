"""Free syntax and DNS mail-routing checks; does NOT establish mailbox existence."""
import csv
from pathlib import Path
from email_validator import validate_email, EmailSyntaxError, EmailUndeliverableError
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill


def check_address(address):
    address = str(address or '').strip()
    if not address:
        return ('', 'MISSING', 'No address provided')
    try:
        result = validate_email(address, check_deliverability=True)
        return (result.normalized, 'DOMAIN ACCEPTS MAIL', 'Syntax and DNS mail-routing passed; mailbox NOT verified')
    except EmailSyntaxError as exc:
        return ('', 'INVALID SYNTAX', str(exc))
    except EmailUndeliverableError as exc:
        message = str(exc)
        if any(s in message.lower() for s in ('timeout', 'could not', 'failed', 'error', 'unavailable', 'no nameserver')):
            return ('', 'INCONCLUSIVE', message)
        return ('', 'DOMAIN CANNOT RECEIVE MAIL', message)
    except Exception as exc:
        return ('', 'INCONCLUSIVE', f'{type(exc).__name__}: {exc}')


def load_data(path):
    path = Path(path)
    if path.suffix.lower() == '.csv':
        with path.open('r', encoding='utf-8-sig', newline='') as handle:
            data = list(csv.reader(handle))
    elif path.suffix.lower() == '.xlsx':
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            data = [list(row) for row in book.active.iter_rows(values_only=True)]
        finally:
            book.close()
    else:
        raise ValueError('Choose an .xlsx or .csv file')
    if not data:
        raise ValueError('The file is empty')
    headers = [str(x).strip() if x is not None else '' for x in data[0]]
    if not any(headers):
        raise ValueError('The first row must contain column headings')
    rows = [(list(row) + [''] * len(headers))[:len(headers)] for row in data[1:]]
    return headers, rows


def email_columns(headers):
    return [i for i, h in enumerate(headers) if 'email' in h.casefold() and not any(
        word in h.casefold() for word in ('ai 1', 'ai 2', 'raw', 'original'))]


def validate_file(source, destination, email_index, progress=None):
    headers, rows = load_data(source)
    if not 0 <= email_index < len(headers):
        raise ValueError('Select an email column')
    book = Workbook()
    sheet = book.active
    sheet.title = 'Validation'
    sheet.append(headers + ['Normalized Email', 'Status', 'Explanation'])
    cache = {}
    totals = {}
    for i, row in enumerate(rows, start=1):
        value = row[email_index]
        key = str(value or '').strip()
        if key not in cache:
            cache[key] = check_address(key)
        normalized, status, explanation = cache[key]
        sheet.append(row + [normalized, status, explanation])
        totals[status] = totals.get(status, 0) + 1
        if progress:
            progress(i, len(rows))
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(color='FFFFFF', bold=True)
        cell.fill = PatternFill('solid', fgColor='24466B')
    for col in sheet.columns:
        letter = col[0].column_letter
        sheet.column_dimensions[letter].width = min(65, max(18, max(len(str(c.value or '')) for c in col[:150]) + 2))
    book.save(destination)
    return totals
