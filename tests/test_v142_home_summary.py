from pathlib import Path
from egomail.profile import normalize_profile


def test_home_summary_keeps_two_final_formula_fields():
    p = normalize_profile({
        'name':'x',
        'fields':[
            {'name':'a','type':'number','final_formula':'sum'},
            {'name':'b','type':'text','final_formula':'count'},
            {'name':'c','type':'number'},
        ],
        'home_summary_fields':['a','b','c'],
    })
    assert p['home_summary_fields'] == ['a','b']


def test_community_ui_displays_home_details():
    ui = Path(__file__).parents[1] / 'egomail' / 'ui.py'
    text = ui.read_text(encoding='utf-8')
    assert 'Data inizio estrazione:' in text
    assert 'Ultima estrazione:' in text
    assert 'Valori schermata iniziale…' in text
