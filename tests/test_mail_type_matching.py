import unittest

from egomail.extraction import extract_message_all, matches_mail_type, matching_mail_types
from egomail.mail import MailMessage


class MailTypeMatchingTests(unittest.TestCase):
    def setUp(self):
        self.mail = MailMessage(
            uid=1,
            sender="Airbnb <automated@airbnb.com>",
            subject="Promemoria di prenotazione: Melanie arriverà presto!",
            date="Wed, 19 Aug 2026 10:00:00 +0200",
            raw_date="Wed, 19 Aug 2026 10:00:00 +0200",
            body="Codice di conferma\nHMWPYHB3RP\nOspiti\n2 adulti, 3 bambini",
        )

    def test_all_compiled_triggers_are_and(self):
        mt = {
            "name": "promemoria",
            "match_rules": [
                {"source": "subject", "mode": "contains", "trigger": "Promemoria di prenotazione"},
                {"source": "sender", "mode": "contains", "trigger": "airbnb.com"},
                {"source": "body", "mode": "regex", "trigger": r"Codice di conferma\s+[A-Z0-9]+"},
            ],
        }
        self.assertTrue(matches_mail_type(self.mail, mt))
        mt["match_rules"][2]["trigger"] = "TESTO CHE NON ESISTE"
        self.assertFalse(matches_mail_type(self.mail, mt))

    def test_empty_trigger_is_ignored_but_empty_type_never_matches(self):
        mt = {
            "name": "tipo",
            "match_rules": [
                {"source": "subject", "mode": "contains", "trigger": "Promemoria"},
                {"source": "body", "mode": "contains", "trigger": ""},
            ],
        }
        self.assertTrue(matches_mail_type(self.mail, mt))
        mt["match_rules"][0]["trigger"] = ""
        self.assertFalse(matches_mail_type(self.mail, mt))

    def test_multiple_rules_on_same_source_must_all_match(self):
        mt = {
            "name": "tipo",
            "match_rules": [
                {"source": "subject", "mode": "contains", "trigger": "Promemoria"},
                {"source": "subject", "mode": "contains", "trigger": "Melanie"},
            ],
        }
        self.assertTrue(matches_mail_type(self.mail, mt))
        mt["match_rules"][1]["trigger"] = "Nathan"
        self.assertFalse(matches_mail_type(self.mail, mt))

    def test_one_mail_can_match_more_than_one_type_and_both_extract(self):
        profile = {
            "fields": [{"name": "a", "type": "text"}, {"name": "b", "type": "text"}],
            "mail_types": [
                {
                    "name": "tipo_a",
                    "match_rules": [{"source": "sender", "mode": "contains", "trigger": "Airbnb"}],
                    "rules": [{"field": "a", "source": "subject", "strategy": "constant", "value": "A", "transform": "text"}],
                },
                {
                    "name": "tipo_b",
                    "match_rules": [{"source": "subject", "mode": "contains", "trigger": "Melanie"}],
                    "rules": [{"field": "b", "source": "subject", "strategy": "constant", "value": "B", "transform": "text"}],
                },
            ],
        }
        self.assertEqual([x["name"] for x in matching_mail_types(self.mail, profile)], ["tipo_a", "tipo_b"])
        results = extract_message_all(self.mail, profile)
        self.assertEqual(results, [("tipo_a", {"a": "A"}), ("tipo_b", {"b": "B"})])


if __name__ == "__main__":
    unittest.main()
