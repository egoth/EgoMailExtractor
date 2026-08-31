from pathlib import Path
from egomail import __version__


def test_v1451_version_and_shutil_import():
    assert __version__ == "1.45.2"
    source = (Path(__file__).parents[1] / "egomail" / "ui.py").read_text(encoding="utf-8")
    assert any(line.strip() == "import shutil" for line in source.splitlines()[:30])


def test_reading_checkpoint_is_progressive_and_resumable():
    source = (Path(__file__).parents[1] / "egomail" / "ui.py").read_text(encoding="utf-8")
    assert '"scan_uids": unique_uids' in source
    assert '"processed_uids": sorted(processed_read_uids)' in source
    assert '"reading_complete": False' in source
    assert 'batch_size = 50' in source
    assert 'riprendo la Lettura dal checkpoint' in source
    assert 'checkpoint salvato' in source
