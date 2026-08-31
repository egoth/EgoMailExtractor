from egomail.support_mail import infer_smtp_from_imap


def test_gmail_smtp_inference():
    assert infer_smtp_from_imap("imap.gmail.com") == ("smtp.gmail.com", 465, "SSL")


def test_office365_smtp_inference():
    assert infer_smtp_from_imap("outlook.office365.com") == ("smtp.office365.com", 587, "STARTTLS")


def test_generic_imap_smtp_inference():
    assert infer_smtp_from_imap("imap.example.org") == ("smtp.example.org", 465, "SSL")


def test_default_support_recipient_is_configured():
    from pathlib import Path
    ui = (Path(__file__).parents[1] / "egomail" / "ui.py").read_text(encoding="utf-8")
    assert 'to_v = tk.StringVar(value="ennio@domoticachepassione.it")' in ui
