import json
import zipfile
from pathlib import Path

from egomail import __version__
from egomail.profile import ProfileManager, new_profile
from egomail.profile_transfer import (
    ABBREVIATIONS_FILENAME,
    export_profile_package,
    import_profile_package,
    inspect_profile_package,
    merge_regex_abbreviations,
)


def test_version_138():
    assert __version__ == "1.45.2"


def test_export_and_import_profile_package(tmp_path):
    src = tmp_path / "src"; dst = tmp_path / "out"; imported = tmp_path / "imported"
    pm = ProfileManager(src)
    path = pm.create("Airbnb soggiorni")
    prof = pm.load(path)
    prof["fields"] = [{"name": "ospite", "type": "text"}]
    pm.save(path, prof)
    abbr = [{"token":"--MIO--", "replacement":"[A-Z]+"}]
    zpath = export_profile_package(path, dst, abbr)
    assert zpath.exists()
    with zipfile.ZipFile(zpath) as zf:
        names = zf.namelist()
        assert ABBREVIATIONS_FILENAME in names
        assert len([x for x in names if x.endswith('.json') and x != ABBREVIATIONS_FILENAME]) == 1
    data, got_abbr, member = inspect_profile_package(zpath)
    assert data["name"] == "Airbnb soggiorni"
    assert got_abbr[0]["token"] == "--MIO--"
    pm2 = ProfileManager(imported)
    out, got_abbr2 = import_profile_package(zpath, pm2)
    assert out.exists()
    assert pm2.load(out)["name"] == "Airbnb soggiorni"
    assert got_abbr2 == got_abbr


def test_import_duplicate_uses_new_filename(tmp_path):
    pm = ProfileManager(tmp_path / "src")
    p = pm.create("Profilo X")
    z = export_profile_package(p, tmp_path / "z", [])
    target = ProfileManager(tmp_path / "target")
    a, _ = import_profile_package(z, target)
    b, _ = import_profile_package(z, target)
    assert a != b
    assert a.exists() and b.exists()


def test_regex_merge_import_wins_on_conflict():
    current = [{"token":"--A--", "replacement":"old"}, {"token":"--B--", "replacement":"b"}]
    imported = [{"token":"--A--", "replacement":"new"}, {"token":"--C--", "replacement":"c"}]
    merged = merge_regex_abbreviations(current, imported)
    by = {x['token']: x['replacement'] for x in merged}
    assert by == {"--A--":"new", "--B--":"b", "--C--":"c"}
