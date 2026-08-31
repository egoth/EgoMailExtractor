import json, zipfile
from pathlib import Path
from egomail.profile_transfer import export_profiles_backup, inspect_profiles_backup, read_profile_from_backup

def test_multi_profile_backup_and_selective_inspection(tmp_path):
    p1=tmp_path/"a.json"; p2=tmp_path/"b.json"
    p1.write_text(json.dumps({"name":"A"}),encoding="utf-8")
    p2.write_text(json.dumps({"name":"B"}),encoding="utf-8")
    x=tmp_path/"A - Mail.xlsx"; x.write_bytes(b"excel")
    z=export_profiles_backup([(p1,{"mail1":{"last_processed_uid":10}},[(x.name,x)]),(p2,{"mail1":{"last_processed_uid":20}},[])],tmp_path,[])
    entries,abbr=inspect_profiles_backup(z)
    assert [e["name"] for e in entries]==["A","B"]
    prof,state=read_profile_from_backup(z,entries[1])
    assert prof["name"]=="B" and state["mail1"]["last_processed_uid"]==20
    with zipfile.ZipFile(z) as zf:
        assert any(n.endswith("A - Mail.xlsx") for n in zf.namelist())

def test_home_has_two_phase_progress_source():
    source=(Path(__file__).parents[1]/"egomail"/"ui.py").read_text(encoding="utf-8")
    assert "Lettura: 0%" in source
    assert "Estrazione: 0%" in source
    assert "Profilo in estrazione:" in source
