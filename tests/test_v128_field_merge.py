import unittest

from egomail.profile import merge_field_into, normalize_profile


class FieldMergeV128Tests(unittest.TestCase):
    def test_normalize_consolidates_duplicate_field_definitions(self):
        profile = {
            "fields": [
                {"name": "numero_prenotazione", "type": "text"},
                {"name": "numero_prenotazione", "type": "number", "final_formula": "sum"},
            ],
            "mail_types": [],
            "reconcile_keys": [],
        }
        normalize_profile(profile)
        self.assertEqual(1, len(profile["fields"]))
        self.assertEqual("numero_prenotazione", profile["fields"][0]["name"])
        self.assertEqual("number", profile["fields"][0]["type"])
        self.assertEqual("sum", profile["fields"][0]["final_formula"])

    def test_merge_repoints_every_field_reference(self):
        profile = {
            "fields": [
                {"name": "campo_prenotazione", "type": "text"},
                {"name": "numero_prenotazione", "type": "text"},
                {"name": "data_inizio", "type": "date"},
            ],
            "mail_types": [
                {
                    "name": "prenotazione",
                    "match_rules": [],
                    "rules": [
                        {"name": "r1", "field": "campo_prenotazione", "source": "body", "strategy": "regex", "pattern": "(X)", "group": 1},
                        {"name": "r2", "field": "numero_prenotazione", "source": "body", "strategy": "regex", "pattern": "(Y)", "group": 1},
                    ],
                }
            ],
            "reconcile_keys": ["campo_prenotazione"],
            "computed_fields": [
                {"field": "data_inizio", "kind": "days_between", "start_field": "campo_prenotazione", "end_field": "numero_prenotazione"}
            ],
            "summary_formulas": [
                {"name": "tot", "formula": "=conta_campo_prenotazione+conta_numero_prenotazione"}
            ],
        }
        merge_field_into(profile, "campo_prenotazione", "numero_prenotazione", source_index=0)

        names = [f["name"] for f in profile["fields"]]
        self.assertEqual(1, names.count("numero_prenotazione"))
        self.assertNotIn("campo_prenotazione", names)
        rules = profile["mail_types"][0]["rules"]
        self.assertTrue(all(r["field"] == "numero_prenotazione" for r in rules))
        self.assertEqual(["numero_prenotazione"], profile["reconcile_keys"])
        comp = profile["computed_fields"][0]
        self.assertEqual("numero_prenotazione", comp["start_field"])
        self.assertEqual("numero_prenotazione", comp["end_field"])
        self.assertEqual(
            "=conta_numero_prenotazione+conta_numero_prenotazione",
            profile["summary_formulas"][0]["formula"],
        )

    def test_merge_already_duplicated_name_by_index(self):
        profile = {
            "fields": [
                {"name": "numero_prenotazione", "type": "text"},
                {"name": "numero_prenotazione", "type": "number"},
            ],
            "mail_types": [],
            "reconcile_keys": [],
        }
        merge_field_into(profile, "numero_prenotazione", "numero_prenotazione", source_index=1)
        self.assertEqual(1, len(profile["fields"]))
        self.assertEqual("numero_prenotazione", profile["fields"][0]["name"])
        self.assertEqual("number", profile["fields"][0]["type"])


if __name__ == "__main__":
    unittest.main()
