
from egomail.exporter import apply_detailed_extractions


def test_incremental_log_fragment_has_no_session_header_or_footer():
    profile = {"name": "Test", "fields": [], "mail_types": [], "reconcile_keys": []}
    event = {"mail": type("M", (), {"uid": 1, "message_id": "x", "date": "", "sender": "a", "subject": "b"})(), "types": []}
    _, stats, text = apply_detailed_extractions(
        [], [event], profile, mode="estrazione incrementale",
        include_header=False, include_footer=False,
    )
    assert "MAIL ESAMINATA" in text
    assert "INIZIO ESTRAZIONE INCREMENTALE" not in text
    assert "FINE ESTRAZIONE INCREMENTALE" not in text
    assert stats["mails"] == 1
