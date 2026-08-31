import unittest

from egomail.profile import copy_extraction_rules, normalize_profile


class CopyExtractionRulesTests(unittest.TestCase):
    def setUp(self):
        self.source = normalize_profile({
            "name": "Sorgente",
            "fields": [
                {"name": "codice", "type": "text"},
                {"name": "importo", "type": "number", "format": "currency_eur"},
            ],
            "reconcile_keys": ["codice"],
            "mail_types": [{
                "name": "tipo_a",
                "match_rules": [],
                "rules": [
                    {"name": "estrai_codice", "field": "codice", "source": "body", "strategy": "regex", "pattern": "(X+)", "group": 1},
                    {"name": "estrai_importo", "field": "importo", "source": "body", "strategy": "regex", "pattern": "([0-9]+)", "group": 1, "existing_value_policy": "sum"},
                ],
            }],
        })
        self.destination = normalize_profile({
            "name": "Destinazione",
            "fields": [],
            "reconcile_keys": [],
            "mail_types": [{"name": "tipo_b", "match_rules": [], "rules": []}],
        })

    def test_copy_multiple_rules_and_fields(self):
        result = copy_extraction_rules(self.source, "tipo_a", [0, 1], self.destination, "tipo_b")
        rules = self.destination["mail_types"][0]["rules"]
        self.assertEqual(2, result["copied_count"])
        self.assertEqual(["estrai_codice", "estrai_importo"], [r["name"] for r in rules])
        self.assertEqual(["codice", "importo"], [f["name"] for f in self.destination["fields"]])
        self.assertIn("codice", self.destination["reconcile_keys"])
        self.assertEqual("sum", rules[1]["existing_value_policy"])
        self.assertEqual("currency_eur", self.destination["fields"][1]["format"])

    def test_duplicate_rule_name_is_renamed(self):
        self.destination["fields"].append({"name": "codice", "type": "text"})
        self.destination["mail_types"][0]["rules"].append({
            "name": "estrai_codice", "field": "codice", "source": "body", "strategy": "regex", "pattern": "old"
        })
        result = copy_extraction_rules(self.source, "tipo_a", [0], self.destination, "tipo_b")
        self.assertEqual("estrai_codice_copia", result["copied_names"][0])
        self.assertEqual("estrai_codice_copia", self.destination["mail_types"][0]["rules"][-1]["name"])

    def test_copy_inside_same_mail_type(self):
        result = copy_extraction_rules(self.source, "tipo_a", [0], self.source, "tipo_a")
        self.assertEqual(3, len(self.source["mail_types"][0]["rules"]))
        self.assertEqual("estrai_codice_copia", result["copied_names"][0])

    def test_invalid_destination_type(self):
        with self.assertRaises(ValueError):
            copy_extraction_rules(self.source, "tipo_a", [0], self.destination, "manca")


if __name__ == "__main__":
    unittest.main()
