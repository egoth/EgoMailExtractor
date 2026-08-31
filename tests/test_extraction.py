import json
import tempfile
import unittest
from pathlib import Path

from egomail.mail import MailMessage
from egomail.extraction import extract_message, reconcile
from egomail.exporter import load_existing_records, merge_records, write_excel

ROOT = Path(__file__).resolve().parents[1]


class ExtractionTests(unittest.TestCase):
    def setUp(self):
        self.airbnb = json.loads((ROOT / "builtin_profiles" / "airbnb.json").read_text(encoding="utf-8"))
        self.vrbo = json.loads((ROOT / "builtin_profiles" / "vrbo.json").read_text(encoding="utf-8"))

    def test_airbnb_reconcile(self):
        reminder = MailMessage(
            uid=100,
            sender="Airbnb <automated@airbnb.com>",
            subject="Promemoria di prenotazione: Melanie arriverà presto!",
            date="Sun, 9 Aug 2026 10:00:00 +0200",
            raw_date="Sun, 9 Aug 2026 10:00:00 +0200",
            body="""Melanie arriverà domenica, 16 ago.\nMelanie Pereira\n2 Bedroom Apt w/ Terrace and Garden\nIntera casa/appart.\nCheck-in\ndom 16 ago\n15:00\nCheck-out\nven 21 ago\n11:00\nOspiti\n2 adulti, 3 bambini\nCodice di conferma\nHMWPYHB3RP\nTasse di soggiorno\n150,00 €\nTotale (EUR)\n763,94 €""",
        )
        payout = MailMessage(
            uid=101,
            sender="Airbnb <automated@airbnb.com>",
            subject="Abbiamo inviato un compenso di 395,54 € EUR",
            date="Mon, 17 Aug 2026 10:00:00 +0200",
            raw_date="Mon, 17 Aug 2026 10:00:00 +0200",
            body="""Melanie Pereira\n-110,25 € EUR\nRitenuta fiscale per il reddito italiano\nCasa con terrazzo super scontata! (9659089)\nHMWPYHB3RP\nMelanie Pereira\n505,79 € EUR\nAlloggio\nCasa con terrazzo super scontata! (9659089)\nHMWPYHB3RP\nTotale pagato:\n395,54 € EUR""",
        )
        _, r1 = extract_message(reminder, self.airbnb)
        _, r2 = extract_message(payout, self.airbnb)
        merged = reconcile([r1, r2], self.airbnb)
        self.assertEqual(len(merged), 1)
        rec = merged[0]
        self.assertEqual(rec["id_soggiorno"], "HMWPYHB3RP")
        self.assertEqual(rec["nome_ospite"], "Melanie Pereira")
        self.assertEqual(rec["numero_ospiti"], 5)
        self.assertEqual(rec["data_check_in"], "2026-08-16")
        self.assertEqual(rec["data_check_out"], "2026-08-21")
        self.assertEqual(rec["durata_giorni"], 5)
        self.assertAlmostEqual(rec["importo_pagato_ospite"], 763.94)
        self.assertAlmostEqual(rec["tasse_pagate"], 150.00)
        self.assertAlmostEqual(rec["importo_ricevuto"], 395.54)

    def test_vrbo_booking(self):
        mail = MailMessage(
            uid=200,
            sender="Vrbo",
            subject="Prenotazione per Nathan Beach: 31 mag - 8 giu 2026 - Vrbo Italia #10237222",
            date="Wed, 3 Jun 2026 12:00:00 +0200",
            raw_date="Wed, 3 Jun 2026 12:00:00 +0200",
            body="""Proprietà\n#10237222\nUnità\nunit_5262001\nNumero Prenotazione\nHA-NW84G6\nDate\n31 mag - 8 giu 2026, 8 Notti\nOspiti\n4 adulti, 1 bambino\nNome del viaggiatore\nNathan Beach""",
        )
        _, rec = extract_message(mail, self.vrbo)
        self.assertEqual(rec["id_soggiorno"], "HA-NW84G6")
        self.assertEqual(rec["durata_giorni"], 8)
        self.assertEqual(rec["numero_ospiti"], 5)
        self.assertEqual(rec["nome_ospite"], "Nathan Beach")

    def test_excel_roundtrip_and_update(self):
        original = [{"id_soggiorno": "ABC12345", "nome_ospite": "Mario", "importo_ricevuto": 100.0}]
        update = [{"id_soggiorno": "ABC12345", "nome_ospite": "Mario Rossi", "tasse_pagate": 30.0}]
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "airbnb.xlsx"
            write_excel(p, original, self.airbnb)
            existing = load_existing_records(p, self.airbnb)
            merged = merge_records(existing, update, self.airbnb)
            write_excel(p, merged, self.airbnb)
            reread = load_existing_records(p, self.airbnb)
            self.assertEqual(len(reread), 1)
            self.assertEqual(reread[0]["nome_ospite"], "Mario Rossi")
            self.assertEqual(reread[0]["tasse_pagate"], 30.0)
            self.assertEqual(reread[0]["importo_ricevuto"], 100.0)


if __name__ == "__main__":
    unittest.main()
