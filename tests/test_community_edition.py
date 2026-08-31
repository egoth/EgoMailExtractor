from pathlib import Path

from egomail.edition import EDITION, IS_PRO, DISPLAY_NAME, CONFIG_APP_NAME


def test_community_identity():
    assert EDITION == "Community"
    assert IS_PRO is False
    assert DISPLAY_NAME == "EgoMailExtractor Community"
    assert CONFIG_APP_NAME == "EgoMailExtractorCommunity"


def test_guided_wizard_source_not_distributed():
    package_dir = Path(__file__).resolve().parents[1] / "egomail"
    assert not (package_dir / "wizard.py").exists()


def test_apache_license_present():
    root = Path(__file__).resolve().parents[1]
    text = (root / "LICENSE.txt").read_text(encoding="utf-8")
    assert "Apache License" in text
    assert "Version 2.0" in text


def test_builtin_profiles_are_not_installed_by_community_startup():
    root = Path(__file__).resolve().parents[1]
    app_text = (root / "egomail" / "app.py").read_text(encoding="utf-8")
    assert "install_builtin_profiles" not in app_text
