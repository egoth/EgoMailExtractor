from pathlib import Path

from egomail.extraction import header_prefilter_matches_profile
from egomail.exporter import build_period_summary_formula
from egomail.mail import MailMessage
from egomail.pipeline import event_to_dict, dict_to_event, atomic_write_json, read_json
from egomail.profile import normalize_profile


def _paypal_profile():
    return {
        "name": "Paypal",
        "fields": [{"name": "importo", "type": "number"}],
        "mail_types": [{
            "name": "Ricevuta",
            "match_rules": [
                {"source": "sender", "mode": "contains", "trigger": "paypal.it"},
                {"source": "subject", "mode": "contains", "trigger": "Ricevuta"},
            ],
            "rules": [],
        }],
        "computed_fields": [],
        "summary_formulas": [],
        "home_summary_fields": [],
    }


def test_header_prefilter_applies_sender_and_subject():
    p = normalize_profile(_paypal_profile())
    ok = MailMessage(uid=1, sender="service@paypal.it", subject="Ricevuta pagamento", date="", body="")
    bad_sender = MailMessage(uid=2, sender="x@example.com", subject="Ricevuta pagamento", date="", body="")
    bad_subject = MailMessage(uid=3, sender="service@paypal.it", subject="Avviso", date="", body="")
    assert header_prefilter_matches_profile(ok, p)
    assert not header_prefilter_matches_profile(bad_sender, p)
    assert not header_prefilter_matches_profile(bad_subject, p)


def test_body_rule_is_deferred_after_header_prefilter():
    p = _paypal_profile()
    p["mail_types"][0]["match_rules"].append(
        {"source": "body", "mode": "contains", "trigger": "codice segreto"}
    )
    p = normalize_profile(p)
    header_only = MailMessage(uid=1, sender="service@paypal.it", subject="Ricevuta pagamento", date="", body="")
    assert header_prefilter_matches_profile(header_only, p)


def test_period_summary_formula_supports_rolling_and_calendar():
    fields = ["data", "importo"]
    rolling = build_period_summary_formula({
        "kind": "period_aggregate", "operation": "sum", "field": "importo",
        "date_field": "data", "period": "month", "period_mode": "rolling",
    }, fields, 10)
    assert "SUMIFS" in rolling and "TODAY()-30" in rolling

    calendar = build_period_summary_formula({
        "kind": "period_aggregate", "operation": "count", "field": "importo",
        "date_field": "data", "period": "month", "period_mode": "calendar",
    }, fields, 10)
    assert "COUNTIFS" in calendar and "EOMONTH(TODAY(),-1)+1" in calendar


def test_profile_keeps_temporal_summary_and_six_home_refs():
    profile = normalize_profile({
        "name": "x",
        "fields": [
            {"name": "data", "type": "date"},
            {"name": "importo", "type": "number", "final_formula": "sum"},
            {"name": "a", "type": "number", "final_formula": "sum"},
            {"name": "b", "type": "number", "final_formula": "sum"},
            {"name": "c", "type": "number", "final_formula": "sum"},
            {"name": "d", "type": "number", "final_formula": "sum"},
            {"name": "calc", "type": "number"},
        ],
        "computed_fields": [{"field":"calc","kind":"days_between","start_field":"data","end_field":"data","priority":"rule_wins"}],
        "summary_formulas": [{
            "name": "mese_importo", "kind": "period_aggregate", "operation": "sum",
            "field": "importo", "date_field": "data", "period": "month", "period_mode": "calendar",
        }],
        "home_summary_fields": ["importo", "a", "b", "c", "computed:calc", "summary:mese_importo", "d"],
    })
    assert len(profile["home_summary_fields"]) == 6
    assert "computed:calc" in profile["home_summary_fields"]
    assert "summary:mese_importo" in profile["home_summary_fields"]
    assert profile["summary_formulas"][0]["period_mode"] == "calendar"


def test_operation_event_records_reconciliation_and_roundtrips(tmp_path: Path):
    mail = MailMessage(uid=9, sender="a", subject="s", date="d", body="b")
    event = {
        "mail": mail,
        "types": [{
            "mail_type": "Ricevuta",
            "selection_rules": [],
            "candidate": {"id": "X", "importo": 12},
            "updates": [{
                "rule_name": "Importo", "field": "importo", "value": 12,
                "policy": "replace", "matched": True,
            }],
        }],
    }
    payload = event_to_dict(event, {"reconcile_keys": ["id"]})
    assert payload["operations"][0]["requires_reconciliation"] is True
    assert payload["operations"][0]["reconcile_keys"] == {"id": "X"}
    restored = dict_to_event(payload)
    assert restored["mail"].uid == 9
    p = tmp_path / "x.json"
    atomic_write_json(p, {"events": [payload]})
    assert read_json(p)["events"][0]["mail"]["uid"] == 9
