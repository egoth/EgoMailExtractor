import json
from pathlib import Path
from egomail.config import ConfigManager
from egomail.profile import ProfileManager
from egomail.profile_transfer import clean_profile_for_export, export_profile_package, inspect_profile_package

def test_state_is_pair_email_profile(tmp_path):
    cfg=ConfigManager(tmp_path/'cfg')
    cfg._active_imap_profile_id='mailA'
    a=cfg.extraction_state('airbnb.json'); a['last_processed_uid']=10; a['home_summary_values']=[{'field':'x','value':1}]
    b=cfg.extraction_state('paypal.json'); b['last_processed_uid']=20
    cfg._active_imap_profile_id='mailB'
    c=cfg.extraction_state('airbnb.json'); c['last_processed_uid']=30
    assert a['last_processed_uid']==10 and b['last_processed_uid']==20 and c['last_processed_uid']==30

def test_normal_export_strips_operational_state(tmp_path):
    d=tmp_path/'profiles'; mgr=ProfileManager(d); p=mgr.create('Airbnb')
    raw=mgr.load(p); raw['state']={'last_processed_uid':99,'home_summary_values':[{'field':'x','value':3}]}; raw['state_by_email']={'a':{'last_processed_uid':88}}; raw['extraction_start_date']='2026-01-01'
    # Scriviamo direttamente per simulare un profilo legacy.
    p.write_text(json.dumps(raw),encoding='utf-8')
    z=export_profile_package(p,tmp_path/'out',[])
    prof,_,_=inspect_profile_package(z)
    assert 'state' not in prof and 'state_by_email' not in prof and 'extraction_start_date' not in prof

def test_profile_save_does_not_persist_state(tmp_path):
    mgr=ProfileManager(tmp_path); p=mgr.create('Test'); prof=mgr.load(p); prof['state']['last_processed_uid']=44; prof['state_by_email']={'x':{'last_processed_uid':55}}
    mgr.save(p,prof); raw=json.loads(p.read_text(encoding='utf-8'))
    assert 'state' not in raw and 'state_by_email' not in raw
