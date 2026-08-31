import tempfile
import unittest
from pathlib import Path

from egomail.config import ConfigManager
from egomail.extraction import extract_rule, matches_mail_type
from egomail.mail import MailMessage
from egomail.regex_utils import expand_regex, set_regex_abbreviations


class RegexAbbreviationTests(unittest.TestCase):
    def tearDown(self):
        set_regex_abbreviations([])

    def test_expansion_is_applied_to_selection_and_extraction_regex(self):
        set_regex_abbreviations([
            {"token": "--ALFANUMERICO--", "replacement": r"[A-Z0-9]+"},
            {"token": "--NUMERO--", "replacement": r"\d+"},
        ])
        mail = MailMessage(
            uid=1, sender="Airbnb <automated@airbnb.com>",
            subject="Codice HMWPYHB3RP", date="", raw_date="",
            body="Ospiti 5",
        )
        mt = {
            "match_rules": [{
                "name": "m", "source": "subject", "mode": "regex",
                "trigger": r"Codice\s+(--ALFANUMERICO--)",
            }]
        }
        self.assertTrue(matches_mail_type(mail, mt))
        rule = {
            "source": "body", "strategy": "regex",
            "pattern": r"Ospiti\s+(--NUMERO--)", "group": 1, "transform": "int",
        }
        self.assertEqual(extract_rule(mail, rule), 5)

    def test_nested_expansion(self):
        set_regex_abbreviations([
            {"token": "--CHAR--", "replacement": r"[A-Z0-9]"},
            {"token": "--CODICE--", "replacement": r"--CHAR--+"},
        ])
        self.assertEqual(expand_regex("(--CODICE--)"), r"([A-Z0-9]+)")

    def test_preferences_persist(self):
        with tempfile.TemporaryDirectory() as td:
            cfg = ConfigManager(Path(td))
            cfg.set_regex_abbreviations([
                {"token": "--X--", "replacement": r"\d+"},
            ])
            cfg2 = ConfigManager(Path(td))
            self.assertEqual(cfg2.regex_abbreviations, [{"token": "--X--", "replacement": r"\d+"}])


if __name__ == "__main__":
    unittest.main()
