import unittest
from unittest.mock import patch
from tempfile import TemporaryDirectory
from pathlib import Path
from openpyxl import Workbook, load_workbook
from validator import check_address, smtp_check, load_data, email_columns, validate_file

class ValidatorTests(unittest.TestCase):
    def test_syntax(self):
        self.assertEqual(check_address('ejbrown##@@!!@gfart.com')[1], 'INVALID SYNTAX')
        self.assertEqual(check_address('not-an-email')[1], 'INVALID SYNTAX')
        self.assertEqual(check_address('')[1], 'MISSING')

    def test_smtp_rejected(self):
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'], lambda h,a:('RCPT TO',550,'5.1.1 user unknown'))[0], 'MAILBOX REJECTED')

    def test_smtp_catchall(self):
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'], lambda h,a:(250,'ok'))[0], 'CATCH-ALL / UNKNOWN')

    def test_smtp_target_accepted_random_rejected(self):
        def probe(host, address):
            return ('RCPT TO', 250, 'ok') if address == 'person@example.com' else ('RCPT TO', 550, '5.1.1 unknown user')
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'], probe)[0], 'ACCEPTED (UNCONFIRMED)')

    def test_smtp_timeout(self):
        def probe(host, address):
            raise TimeoutError('blocked')
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'], probe)[0], 'INCONCLUSIVE')

    def test_hotmail_sender_block_is_inconclusive(self):
        def probe(host, address):
            return ('MAIL FROM', 550, '5.7.1 Service unavailable, Client host blocked using Spamhaus')
        status, reason = smtp_check('lizbrandt@hotmail.com', ['mx.hotmail.com'], probe)
        self.assertEqual(status, 'INCONCLUSIVE')
        self.assertIn('MAIL FROM', reason)

    def test_recipient_policy_rejection_is_inconclusive(self):
        def probe(host, address):
            return ('RCPT TO', 550, '5.7.1 blocked by policy')
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'], probe)[0], 'INCONCLUSIVE')

    def test_unknown_550_is_inconclusive(self):
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'],
                         lambda h, a: ('RCPT TO', 550, 'Rejected'))[0], 'INCONCLUSIVE')

    def test_bell_disconnection_is_inconclusive(self):
        def probe(host, address):
            import smtplib
            raise smtplib.SMTPServerDisconnected('Connection unexpectedly closed')
        self.assertEqual(smtp_check('person@example.com', ['mx.example.com'], probe)[0], 'INCONCLUSIVE')

    def test_gmail_known_cases(self):
        def accepted(host, address):
            if address == 'amymullen00@gmail.com':
                return ('RCPT TO', 250, 'OK')
            return ('RCPT TO', 550, '5.1.1 The email account does not exist')
        self.assertEqual(smtp_check('amymullen00@gmail.com', ['gmail.mx'], accepted)[0],
                         'ACCEPTED (UNCONFIRMED)')
        self.assertEqual(smtp_check('amymullen66@gmail.com', ['gmail.mx'], accepted)[0],
                         'MAILBOX REJECTED')

    def test_excel(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / 'input.xlsx'
            output = Path(tmp) / 'output.xlsx'
            book = Workbook()
            book.active.append(['Name', 'Email Address'])
            book.active.append(['Jane', 'not-an-email'])
            book.save(source)
            headers, rows = load_data(source)
            self.assertEqual(email_columns(headers), [1])
            self.assertEqual(len(rows), 1)
            totals = validate_file(source, output, 1)
            self.assertEqual(totals, {'INVALID SYNTAX': 1})
            self.assertEqual(load_workbook(output).active['D2'].value, 'INVALID SYNTAX')

if __name__ == '__main__':
    unittest.main()
