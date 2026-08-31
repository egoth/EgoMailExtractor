import json
import queue
import threading
from pathlib import Path

import egomail.ui as ui
from egomail import __version__
from egomail.ui import MainWindow, SummaryFormulaEditorDialog


def test_version_134():
    assert __version__ == "1.45.2"


def test_imap_retry_automatic_then_user_retry(monkeypatch):
    monkeypatch.setattr(ui.time, "sleep", lambda _n: None)

    class Dummy:
        extraction_stop_event = threading.Event()
        ui_queue = queue.Queue()
        asked = 0
        def _ask_retry_imap_from_worker(self, error, context):
            self.asked += 1
            return True

    dummy = Dummy()
    calls = {"n": 0}
    def operation():
        calls["n"] += 1
        if calls["n"] < 3:
            raise OSError("socket error: EOF")
        return "ok"

    result = MainWindow._imap_call_with_reconnect(dummy, operation, "test")
    assert result == "ok"
    assert calls["n"] == 3
    assert dummy.asked == 1


def test_checkpoint_roundtrip(tmp_path):
    class Dummy:
        profile_path = tmp_path / "Paypal.json"
        _extraction_checkpoint_path = MainWindow._extraction_checkpoint_path

    dummy = Dummy()
    excel = tmp_path / "Paypal.xlsx"
    excel.write_bytes(b"excel")
    MainWindow._save_extraction_checkpoint(dummy, "INBOX", 123, 100, {101, 102}, excel)
    processed = MainWindow._load_extraction_checkpoint(dummy, "INBOX", 123, 100, excel)
    assert processed == {101, 102}
    data = json.loads((tmp_path / "Paypal_estrazione_checkpoint.json").read_text(encoding="utf-8"))
    assert data["processed_uids"] == [101, 102]
    MainWindow._clear_extraction_checkpoint(dummy)
    assert not (tmp_path / "Paypal_estrazione_checkpoint.json").exists()


def test_alias_click_inserts_at_cursor():
    class FakeList:
        def nearest(self, y): return 0
        def get(self, i): return "somma_importo"
    class FakeEntry:
        def __init__(self): self.calls = []
        def focus_set(self): self.calls.append(("focus",))
        def insert(self, where, text): self.calls.append(("insert", where, text))
    class Event: y = 3
    class Dummy:
        alias_list = FakeList()
        formula_entry = FakeEntry()
    d = Dummy()
    SummaryFormulaEditorDialog._insert_alias_from_click(d, Event())
    assert ("insert", "insert", "somma_importo") in d.formula_entry.calls
