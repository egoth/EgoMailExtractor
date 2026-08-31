from pathlib import Path
from egomail import __version__
def test_version(): assert __version__ == '1.45.2'
def test_pro_link():
    s=Path('egomail/ui.py').read_text(encoding='utf-8')
    assert 'https://www.domoticachepassione.it/wp/acquista-egomailextractor-pro/' in s
