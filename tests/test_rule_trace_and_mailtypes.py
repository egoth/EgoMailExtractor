import unittest

from egomail.extraction import extract_rule_trace, humanize_extraction_rule
from egomail.mail import MailMessage
from egomail.profile import delete_mail_type, rename_mail_type


class RuleTraceTests(unittest.TestCase):
    def setUp(self):
        self.mail = MailMessage(
            uid=1,
            sender="Airbnb <automated@airbnb.com>",
            subject="Promemoria di prenotazione",
            date="Sun, 9 Aug 2026 10:00:00 +0200",
            raw_date="Sun, 9 Aug 2026 10:00:00 +0200",
            body="Check-in\ndom 16 ago\n\nCodice di conferma\nHMWPYHB3RP\n",
        )

    def test_after_trigger_trace_spans(self):
        rule = {
            "field": "id_soggiorno",
            "source": "body",
            "strategy": "after_trigger",
            "trigger": "Codice di conferma",
            "nonempty_offset": 1,
            "transform": "text",
        }
        trace = extract_rule_trace(self.mail, rule)
        self.assertEqual(trace.value, "HMWPYHB3RP")
        self.assertTrue(trace.matched)
        self.assertEqual(trace.source_text[slice(*trace.trigger_span)], "Codice di conferma")
        self.assertEqual(trace.source_text[slice(*trace.value_span)], "HMWPYHB3RP")

    def test_regex_trace_uses_group_as_value_span(self):
        rule = {
            "field": "checkin",
            "source": "body",
            "strategy": "regex",
            "pattern": r"Check-in\s+([^\n]+)",
            "group": 1,
            "transform": "text",
        }
        trace = extract_rule_trace(self.mail, rule)
        self.assertEqual(trace.value, "dom 16 ago")
        self.assertIn("Check-in", trace.source_text[slice(*trace.trigger_span)])
        self.assertEqual(trace.source_text[slice(*trace.value_span)], "dom 16 ago")

    def test_human_translation_mentions_field_and_excel_type(self):
        rule = {
            "field": "importo_ricevuto",
            "source": "body",
            "strategy": "after_trigger",
            "trigger": "Totale pagato:",
            "nonempty_offset": 1,
            "transform": "money",
        }
        text = humanize_extraction_rule(rule, field_type="number", currency=True, reconcile_key=False)
        self.assertIn("Totale pagato:", text)
        self.assertIn("importo_ricevuto", text)
        self.assertIn("NUMERO", text)
        self.assertIn("valuta €", text)


class MailTypeManagementTests(unittest.TestCase):
    def test_rename_and_delete_mail_type(self):
        profile = {"mail_types": [{"name": "vecchio", "match_rules": [], "rules": [{"name": "r"}]}]}
        new_name = rename_mail_type(profile, "vecchio", "Nuovo Tipo")
        self.assertEqual(new_name, "nuovo_tipo")
        self.assertEqual(profile["mail_types"][0]["name"], "nuovo_tipo")
        self.assertTrue(delete_mail_type(profile, "nuovo_tipo"))
        self.assertEqual(profile["mail_types"], [])

    def test_rename_rejects_duplicate(self):
        profile = {"mail_types": [{"name": "a"}, {"name": "b"}]}
        with self.assertRaises(ValueError):
            rename_mail_type(profile, "a", "b")


if __name__ == "__main__":
    unittest.main()
