"""Email syntax, DNS and optional best-effort SMTP recipient checks.

An SMTP 250 response does not prove mailbox existence. No DATA is sent.
"""
import csv
import secrets
import smtplib
import socket
import time
from pathlib import Path
from email_validator import validate_email, EmailSyntaxError, EmailUndeliverableError
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
import dns.resolver

TIMEOUT = 6

def dns_check(address):
    address = str(address or '').strip()
    if not address:
        return '', 'MISSING', 'No address provided', []
    try:
        result = validate_email(address, check_deliverability=False)
    except EmailSyntaxError as exc:
        return '', 'INVALID SYNTAX', str(exc), []
    normalized = result.normalized
    domain = result.domain
    try:
        mx = dns.resolver.resolve(domain, 'MX', lifetime=TIMEOUT)
        records = sorted((int(r.preference), str(r.exchange).rstrip('.')) for r in mx)
        if any(host == '' for _, host in records):
            return normalized, 'NO MAIL ROUTING', 'Domain has a null MX record', []
        hosts = [host for _, host in records]
        return normalized, 'DOMAIN ACCEPTS MAIL', 'MX records exist; mailbox NOT verified', hosts
    except dns.resolver.NXDOMAIN:
        return normalized, 'NO MAIL ROUTING', 'Domain does not exist (NXDOMAIN)', []
    except dns.resolver.NoAnswer:
        # SMTP permits A/AAAA fallback when no MX is published.
        try:
            dns.resolver.resolve(domain, 'A', lifetime=TIMEOUT)
            return normalized, 'DOMAIN ACCEPTS MAIL', 'No MX; A-record fallback exists; mailbox NOT verified', [domain]
        except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN):
            try:
                dns.resolver.resolve(domain, 'AAAA', lifetime=TIMEOUT)
                return normalized, 'DOMAIN ACCEPTS MAIL', 'No MX; AAAA fallback exists; mailbox NOT verified', [domain]
            except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN):
                return normalized, 'NO MAIL ROUTING', 'No MX or A/AAAA mail destination', []
            except Exception as exc:
                return normalized, 'INCONCLUSIVE', f'DNS lookup inconclusive: {exc}', []
        except Exception as exc:
            return normalized, 'INCONCLUSIVE', f'DNS lookup inconclusive: {exc}', []
    except Exception as exc:
        return normalized, 'INCONCLUSIVE', f'DNS lookup inconclusive: {exc}', []

def smtp_rcpt(host, recipient, timeout=TIMEOUT):
    """Return (stage, SMTP code, reply). Never send DATA.

    Stage is essential: a 550 at MAIL FROM does not reject the mailbox.
    """
    with smtplib.SMTP(timeout=timeout, local_hostname='localhost') as smtp:
        smtp.connect(host, 25)
        code, reply = smtp.ehlo_or_helo_if_needed() or (250, b'EHLO OK')
        if code not in (250, 251):
            return 'EHLO', code, _reply_text(reply)
        code, reply = smtp.mail('')
        if code not in (250, 251):
            return 'MAIL FROM', code, _reply_text(reply)
        code, reply = smtp.rcpt(recipient)
        return 'RCPT TO', code, _reply_text(reply)


def _reply_text(value):
    return value.decode('utf-8', 'replace') if isinstance(value, bytes) else str(value)


def _probe_result(result):
    """Support older two-field probes without mistaking their stage for RCPT."""
    if len(result) == 3:
        stage, code, reply = result
        return str(stage).upper(), int(code), _reply_text(reply)
    if len(result) == 2:
        code, reply = result
        reply = _reply_text(reply)
        if 'MAIL FROM refused' in reply.upper() or 'MAIL FROM REFUSED' in reply.upper():
            return 'MAIL FROM', int(code), reply
        # Legacy mock probes represent RCPT responses, but real probes should
        # return an explicit stage to make a rejection defensible.
        return 'RCPT TO', int(code), reply
    raise ValueError('SMTP probe must return (stage, code, reply)')


def _recipient_missing(code, reply):
    """Conservatively recognize explicit unknown-recipient responses only."""
    if code not in (550, 551, 553):
        return False
    msg = reply.casefold()
    policy_markers = ('5.7.', 'spamhaus', 'blocked', 'blacklist', 'policy',
                      'access denied', 'service unavailable', 'rate limit',
                      'authentication', 'reputation', 'relay denied')
    if any(marker in msg for marker in policy_markers):
        return False
    return any(marker in msg for marker in (
        '5.1.1', 'user unknown', 'unknown user', 'no such user',
        'mailbox does not exist', 'recipient does not exist',
        'unknown recipient', 'recipient address rejected: user unknown',
        'no such mailbox', 'account that you tried to reach does not exist'))


def smtp_check(address, hosts, probe=smtp_rcpt):
    """Best effort. Acceptance is not proof of inbox existence."""
    if not hosts:
        return 'NOT TESTED', 'No mail host available'
    domain = address.rsplit('@', 1)[-1]
    host = hosts[0]
    try:
        stage, code, reply = _probe_result(probe(host, address))
    except (OSError, smtplib.SMTPException, TimeoutError) as exc:
        return 'INCONCLUSIVE', f'SMTP connection/probe failed: {type(exc).__name__}: {exc}'
    if stage != 'RCPT TO':
        return 'INCONCLUSIVE', f'SMTP {stage} rejected/failed ({code}): {reply}; recipient was not checked'
    if _recipient_missing(code, reply):
        return 'MAILBOX REJECTED', f'SMTP RCPT TO {code}: {reply}'
    if code not in (250, 251):
        return 'INCONCLUSIVE', f'SMTP RCPT TO {code}: {reply}; no definitive mailbox verdict'

    random_recipient = f'nonexistent-{secrets.token_hex(12)}@{domain}'
    try:
        random_stage, random_code, random_reply = _probe_result(probe(host, random_recipient))
    except (OSError, smtplib.SMTPException, TimeoutError) as exc:
        return 'ACCEPTED (UNCONFIRMED)', f'Target accepted (SMTP RCPT TO {code}); catch-all test inconclusive: {exc}'
    if random_stage != 'RCPT TO':
        return 'ACCEPTED (UNCONFIRMED)', (f'Target accepted (SMTP RCPT TO {code}); '
                f'catch-all probe blocked at {random_stage} ({random_code}); mailbox NOT proven')
    if random_code in (250, 251):
        return 'CATCH-ALL / UNKNOWN', (f'Server accepted target and random recipient '
                f'(SMTP {code}/{random_code}); mailbox NOT verified')
    if _recipient_missing(random_code, random_reply):
        return 'ACCEPTED (UNCONFIRMED)', (f'Target accepted (SMTP {code}); random recipient '
                f'explicitly rejected (SMTP {random_code}); mailbox NOT proven')
    return 'ACCEPTED (UNCONFIRMED)', (f'Target accepted (SMTP {code}); catch-all check '
            f'inconclusive (SMTP {random_code}: {random_reply}); mailbox NOT proven')

def check_address(address, use_smtp=False, probe=smtp_rcpt):
    normalized, status, explanation, hosts = dns_check(address)
    if status != 'DOMAIN ACCEPTS MAIL' or not use_smtp:
        return normalized, status, explanation
    smtp_status, smtp_explanation = smtp_check(normalized, hosts, probe=probe)
    return normalized, smtp_status, explanation + '; ' + smtp_explanation

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

def validate_file(source, destination, email_index, progress=None, use_smtp=False):
    headers, rows = load_data(source)
    if not 0 <= email_index < len(headers):
        raise ValueError('Select an email column')
    book = Workbook()
    sheet = book.active
    sheet.title = 'Validation'
    sheet.append(headers + ['Normalized Email', 'Status', 'Explanation'])
    cache, totals = {}, {}
    for i, row in enumerate(rows, start=1):
        key = str(row[email_index] or '').strip()
        if key not in cache:
            cache[key] = check_address(key, use_smtp=use_smtp)
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
