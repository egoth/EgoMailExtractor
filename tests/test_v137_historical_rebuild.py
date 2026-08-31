from datetime import date
from pathlib import Path
from unittest.mock import Mock

from egomail import __version__
from egomail.mail import ImapService
import egomail.ui as ui


def test_version_137():
    assert __version__ == "1.45.2"


def test_uids_since_date_uses_imap_since():
    svc = object.__new__(ImapService)
    conn = Mock()
    conn.uid.return_value = ("OK", [b"10 20 30"])
    svc._connect = Mock(return_value=conn)
    svc._select = Mock(return_value=77)
    uidv, uids = svc.uids_since_date("INBOX", date(2026, 1, 2))
    assert uidv == 77
    assert uids == [10, 20, 30]
    conn.uid.assert_called_once_with("search", None, "SINCE", "02-Jan-2026")


def test_ui_exposes_rebuild_general_extraction():
    assert hasattr(ui.MainWindow, "rebuild_general_extraction")
