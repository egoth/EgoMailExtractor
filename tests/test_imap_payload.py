import email
import unittest

from egomail.mail import imap_fetch_payload, message_body, decode_mime


class ImapPayloadTests(unittest.TestCase):
    def test_fetch_tuple_payload_used_for_headers(self):
        chunks = [
            (b'3163 (UID 3163 BODY[HEADER.FIELDS (FROM SUBJECT DATE MESSAGE-ID)] {149}',
             b'From: Airbnb <automated@airbnb.com>\r\nSubject: Test Airbnb\r\nDate: Wed, 19 Aug 2026 10:00:00 +0200\r\nMessage-ID: <x@test>\r\n\r\n'),
            b')'
        ]
        raw = imap_fetch_payload(chunks)
        msg = email.message_from_bytes(raw)
        self.assertEqual(decode_mime(msg.get('Subject')), 'Test Airbnb')
        self.assertIn('Airbnb', decode_mime(msg.get('From')))

    def test_fetch_tuple_payload_used_for_full_message(self):
        chunks = [
            (b'3163 (UID 3163 BODY[] {200}',
             b'From: Airbnb <automated@airbnb.com>\r\nSubject: Preview\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nCorpo della mail\r\nCheck-in\r\n16 ago\r\n'),
            b')'
        ]
        raw = imap_fetch_payload(chunks)
        msg = email.message_from_bytes(raw)
        self.assertIn('Corpo della mail', message_body(msg))
        self.assertIn('16 ago', message_body(msg))


if __name__ == '__main__':
    unittest.main()


class ProgressiveSearchTests(unittest.TestCase):
    class FakeConn:
        def select(self, folder, readonly=True):
            return "OK", [b"1"]

        def response(self, key):
            return "UIDVALIDITY", [b"123"]

        def status(self, folder, query):
            return "OK", [b"INBOX (UIDVALIDITY 123)"]

        def uid(self, command, *args):
            if command == "search":
                return "OK", [b"10"]
            if command == "fetch":
                uid = args[0]
                spec = args[1]
                if "HEADER.FIELDS" in spec:
                    return "OK", [(
                        b"10 (UID 10 BODY[HEADER.FIELDS] {120}",
                        b"From: Airbnb <automated@airbnb.com>\r\nSubject: Promemoria test\r\nDate: Wed, 19 Aug 2026 10:00:00 +0200\r\n\r\n",
                    ), b")"]
                return "OK", [(
                    b"10 (UID 10 BODY[] {180}",
                    b"From: Airbnb <automated@airbnb.com>\r\nSubject: Promemoria test\r\nDate: Wed, 19 Aug 2026 10:00:00 +0200\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nCodice di conferma\r\nABC12345\r\n",
                ), b")"]
            raise AssertionError((command, args))

        def logout(self):
            return "BYE", []

    def test_progressive_search_can_return_full_message_for_body_matching(self):
        import threading
        from egomail.mail import ImapService, MailMessage

        service = ImapService.__new__(ImapService)
        service._connect = lambda: self.FakeConn()
        found = []
        service.search_progressive(
            "INBOX", "", "", threading.Event(), found.append,
            include_body=True,
        )
        self.assertEqual(len(found), 1)
        self.assertIsInstance(found[0], MailMessage)
        self.assertIn("ABC12345", found[0].body)
