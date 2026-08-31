import tempfile
import unittest
from pathlib import Path

from egomail.exporter import apply_detailed_extractions, apply_existing_value_policy
from egomail.extraction import extract_mail_detailed
from egomail.mail import MailMessage
from egomail.profile import normalize_profile


class UpdatePolicyTests(unittest.TestCase):
    def test_replace_keep_sum(self):
        number = {"name": "totale", "type": "number"}
        text = {"name": "nome", "type": "text"}
        self.assertEqual(apply_existing_value_policy(10, 3, "replace", number)[0], 3)
        self.assertEqual(apply_existing_value_policy(10, 3, "keep", number)[0], 10)
        self.assertEqual(apply_existing_value_policy(10, 3, "sum", number)[0], 13)
        self.assertEqual(apply_existing_value_policy("A", "B", "keep", text)[0], "A")
        currency = {"name": "importo", "type": "number", "format": "currency_eur"}
        self.assertEqual(apply_existing_value_policy(581.54, 69.32, "sum", currency)[0], 650.86)

    def test_sum_normalized_only_for_numeric_field(self):
        profile = normalize_profile({
            "fields": [{"name": "nome", "type": "text"}, {"name": "totale", "type": "number"}],
            "mail_types": [{
                "name": "x", "match_rules": [{"name": "m", "source": "subject", "mode": "contains", "trigger": "x"}],
                "rules": [
                    {"name": "r1", "field": "nome", "source": "subject", "strategy": "constant", "value": "B", "existing_value_policy": "sum"},
                    {"name": "r2", "field": "totale", "source": "subject", "strategy": "constant", "value": "2", "transform": "money", "existing_value_policy": "sum"},
                ],
            }],
        })
        rules = profile["mail_types"][0]["rules"]
        self.assertEqual(rules[0]["existing_value_policy"], "replace")
        self.assertEqual(rules[1]["existing_value_policy"], "sum")

    def test_detailed_log_contains_uid_rules_row_and_field(self):
        profile = normalize_profile({
            "name": "Test",
            "fields": [{"name": "id", "type": "text"}, {"name": "totale", "type": "number"}],
            "reconcile_keys": ["id"],
            "mail_types": [{
                "name": "pagamento",
                "match_rules": [{"name": "seleziona_pagamento", "source": "subject", "mode": "contains", "trigger": "Pagamento"}],
                "rules": [
                    {"name": "id_costante", "field": "id", "source": "body", "strategy": "constant", "value": "ABC", "existing_value_policy": "replace"},
                    {"name": "estrai_totale", "field": "totale", "source": "body", "strategy": "regex", "pattern": r"Totale\s+(\d+)", "group": 1, "transform": "int", "existing_value_policy": "sum"},
                ],
            }],
        })
        mail = MailMessage(uid=7, sender="a@example.com", subject="Pagamento ricevuto", date="Wed, 19 Aug 2026 10:00:00 +0200", raw_date="Wed, 19 Aug 2026 10:00:00 +0200", message_id="<m7>", body="Totale 5")
        event = extract_mail_detailed(mail, profile)
        records, stats, log = apply_detailed_extractions(
            [{"id": "ABC", "totale": 10}], [event], profile, mode="test", excel_path="x.xlsx"
        )
        self.assertEqual(records[0]["totale"], 15)
        self.assertIn("UID: 7", log)
        self.assertIn("Message-ID: <m7>", log)
        self.assertIn("seleziona_pagamento", log)
        self.assertIn("estrai_totale", log)
        self.assertIn("campo=totale", log)
        self.assertIn("Riga Excel: 2", log)
        self.assertIn("azione=SOMMA", log)
        self.assertEqual(stats["field_updates"], 2)


if __name__ == "__main__":
    unittest.main()

class ChronologicalOrderTests(unittest.TestCase):
    def test_mail_sort_key_oldest_first(self):
        from egomail.ui import MainWindow
        older = MailMessage(uid=20, sender="", subject="", date="Mon, 17 Aug 2026 12:00:00 +0200", raw_date="Mon, 17 Aug 2026 12:00:00 +0200")
        newer = MailMessage(uid=10, sender="", subject="", date="Tue, 18 Aug 2026 12:00:00 +0200", raw_date="Tue, 18 Aug 2026 12:00:00 +0200")
        ordered = sorted([newer, older], key=MainWindow._mail_chronological_key)
        self.assertEqual([m.uid for m in ordered], [20, 10])
