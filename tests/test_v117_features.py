import unittest

from egomail.profile import field_generation_sources, normalize_profile


class FieldGenerationSourcesTests(unittest.TestCase):
    def _profile(self):
        return normalize_profile({
            "name": "Lineage",
            "fields": [
                {"name": "id", "type": "text"},
                {"name": "start", "type": "date"},
                {"name": "end", "type": "date"},
                {"name": "durata", "type": "number", "final_formula": "sum"},
                {"name": "cancellata_count", "type": "number", "mail_type_counter": "cancellata"},
                {"name": "solo_excel", "type": "number", "final_formula": "sum"},
            ],
            "computed_fields": [{
                "field": "durata",
                "kind": "days_between",
                "start_field": "start",
                "end_field": "end",
                "priority": "rule_wins",
            }],
            "mail_types": [
                {
                    "name": "confermata",
                    "match_rules": [],
                    "rules": [{
                        "name": "durata_da_body",
                        "field": "durata",
                        "source": "body",
                        "strategy": "regex",
                        "pattern": r"Notti=(\d+)",
                        "group": 1,
                    }],
                },
                {
                    "name": "cancellata",
                    "match_rules": [],
                    "rules": [{
                        "name": "durata_zero",
                        "field": "durata",
                        "source": "body",
                        "strategy": "constant",
                        "value": "0",
                    }],
                },
            ],
        })

    def test_lists_extraction_and_general_computed_generators(self):
        items = field_generation_sources(self._profile(), "durata")
        kinds = [x["kind"] for x in items]
        self.assertEqual(kinds.count("extraction_rule"), 2)
        self.assertEqual(kinds.count("computed_field"), 1)
        names = {x["name"] for x in items}
        self.assertIn("durata_da_body", names)
        self.assertIn("durata_zero", names)
        computed = next(x for x in items if x["kind"] == "computed_field")
        self.assertIn("start", computed["detail"])
        self.assertIn("end", computed["detail"])

    def test_lists_mail_type_counter_as_general_generator(self):
        items = field_generation_sources(self._profile(), "cancellata_count")
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual(item["kind"], "mail_type_counter")
        self.assertEqual(item["mail_type"], "cancellata")
        self.assertIn("Incrementa di 1", item["detail"])

    def test_final_formula_alone_is_not_a_row_generator(self):
        items = field_generation_sources(self._profile(), "solo_excel")
        self.assertEqual(items, [])

    def test_unknown_or_empty_field_returns_no_generators(self):
        profile = self._profile()
        self.assertEqual(field_generation_sources(profile, "manca"), [])
        self.assertEqual(field_generation_sources(profile, ""), [])


if __name__ == "__main__":
    unittest.main()
