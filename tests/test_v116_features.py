import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from egomail.debug_bundle import build_debug_bundle
from egomail.exporter import apply_detailed_extractions
from egomail.extraction import extract_mail_detailed
from egomail.mail import MailMessage
from egomail.profile import normalize_profile


class ComputedFieldPriorityTests(unittest.TestCase):
    def _profile(self, priority="rule_wins"):
        return normalize_profile({
            "name": "Computed",
            "fields": [
                {"name": "id", "type": "text"},
                {"name": "check_in", "type": "date"},
                {"name": "check_out", "type": "date"},
                {"name": "durata", "type": "number"},
            ],
            "reconcile_keys": ["id"],
            "computed_fields": [{
                "field": "durata",
                "kind": "days_between",
                "start_field": "check_in",
                "end_field": "check_out",
                "priority": priority,
            }],
            "mail_types": [{
                "name": "cancellata",
                "match_rules": [{
                    "name": "m", "source": "subject", "mode": "contains", "trigger": "Cancellata"
                }],
                "rules": [
                    {"name": "id", "field": "id", "source": "body", "strategy": "regex", "pattern": r"ID=(\w+)", "group": 1},
                    {"name": "zero", "field": "durata", "source": "body", "strategy": "constant", "value": "0", "transform": "int", "existing_value_policy": "replace"},
                ],
            }],
        })

    def test_explicit_rule_wins_over_computed_field(self):
        profile = self._profile("rule_wins")
        existing = [{"id": "ABC", "check_in": "2026-06-21", "check_out": "2026-06-27", "durata": 6}]
        mail = MailMessage(uid=1, sender="x", subject="Cancellata", date="", message_id="<1>", body="ID=ABC")
        event = extract_mail_detailed(mail, profile)
        records, _, log = apply_detailed_extractions(existing, [event], profile)
        self.assertEqual(records[0]["durata"], 0)
        self.assertIn("valore dopo=0", log)
        self.assertIn("NON RICALCOLATO (priorità al valore della regola corrente)", log)

    def test_always_priority_recalculates_after_explicit_rule(self):
        profile = self._profile("always")
        existing = [{"id": "ABC", "check_in": "2026-06-21", "check_out": "2026-06-27", "durata": 6}]
        mail = MailMessage(uid=1, sender="x", subject="Cancellata", date="", message_id="<1>", body="ID=ABC")
        event = extract_mail_detailed(mail, profile)
        records, _, _ = apply_detailed_extractions(existing, [event], profile)
        self.assertEqual(records[0]["durata"], 6)

    def test_legacy_computed_field_defaults_to_rule_wins(self):
        profile = normalize_profile({
            "fields": [
                {"name": "a", "type": "date"},
                {"name": "b", "type": "date"},
                {"name": "d", "type": "number"},
            ],
            "computed_fields": [{"field": "d", "kind": "days_between", "start_field": "a", "end_field": "b"}],
        })
        self.assertEqual(profile["computed_fields"][0]["priority"], "rule_wins")
        self.assertEqual(profile["schema_version"], 7)


class DebugBundleTests(unittest.TestCase):
    def test_bundle_contains_profile_log_excel_parameters_and_no_credentials(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            log = root / "profilo_estrazione.log"
            excel = root / "ultimo.xlsx"
            log.write_text("LOG", encoding="utf-8")
            excel.write_bytes(b"FAKE-XLSX")
            destination = root / "debug.zip"
            profile = {"name": "Test", "fields": []}
            params = {"imap": {"host": "imap.example", "password": "<NON ESPORTATA>"}}
            manifest = build_debug_bundle(
                destination,
                profile=profile,
                profile_path=root / "profilo.json",
                log_path=log,
                excel_path=excel,
                parameters=params,
            )
            self.assertTrue(destination.exists())
            self.assertFalse(manifest["missing"])
            with zipfile.ZipFile(destination) as zf:
                names = set(zf.namelist())
                self.assertIn("profilo.json", names)
                self.assertIn(log.name, names)
                self.assertIn(excel.name, names)
                self.assertIn("parametri_debug.json", names)
                self.assertIn("manifest_debug.json", names)
                self.assertNotIn("imap_credentials.enc", names)
                loaded = json.loads(zf.read("parametri_debug.json").decode("utf-8"))
                self.assertEqual(loaded["imap"]["password"], "<NON ESPORTATA>")

    def test_bundle_reports_missing_optional_files(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            destination = root / "debug.zip"
            manifest = build_debug_bundle(
                destination,
                profile={"name": "Test"},
                profile_path=root / "profilo.json",
                log_path=root / "missing.log",
                excel_path=None,
                parameters={},
            )
            self.assertEqual(len(manifest["missing"]), 2)


if __name__ == "__main__":
    unittest.main()
