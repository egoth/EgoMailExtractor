from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Iterable, Any

from .profile import ProfileManager, slugify
from .regex_utils import normalize_regex_abbreviations

ABBREVIATIONS_FILENAME = "regex_abbreviations.json"
ICON_PREFIX = "profile_icon"
BACKUP_STATE_FILENAME = "backup_state.json"
BACKUP_META_FILENAME = "backup_meta.json"


def safe_package_name(value: str, fallback: str = "profilo") -> str:
    value = str(value or "").strip()
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value)
    value = re.sub(r"\s+", " ", value).strip(" ._")
    return value[:120] or fallback



def clean_profile_for_export(profile: dict[str, Any]) -> dict[str, Any]:
    """Restituisce solo la configurazione trasferibile, senza stato operativo."""
    import copy
    clean = copy.deepcopy(profile)
    clean.pop("state", None)
    clean.pop("state_by_email", None)
    clean.pop("extraction_start_date", None)
    # I nomi dei due campi riepilogo fanno parte della configurazione; i loro
    # valori invece sono stato email/profilo e non sono nel JSON del profilo.
    return clean

def export_profile_package(
    profile_path: str | Path,
    destination_dir: str | Path,
    regex_abbreviations: Iterable[dict[str, Any]],
) -> Path:
    profile_path = Path(profile_path)
    destination_dir = Path(destination_dir)
    destination_dir.mkdir(parents=True, exist_ok=True)

    profile = clean_profile_for_export(json.loads(profile_path.read_text(encoding="utf-8")))
    display_name = str(profile.get("name") or profile_path.stem)
    base = safe_package_name(display_name, profile_path.stem)
    zip_path = destination_dir / f"{base}.zip"
    if zip_path.exists():
        n = 2
        while (destination_dir / f"{base}_{n}.zip").exists():
            n += 1
        zip_path = destination_dir / f"{base}_{n}.zip"

    profile_member = f"{safe_package_name(profile_path.stem, 'profile')}.json"
    abbreviations = normalize_regex_abbreviations(list(regex_abbreviations))
    icon_source = None
    icon_member = None
    icon_value = str(profile.get("icon", "") or "").strip()
    if icon_value:
        candidate = Path(icon_value)
        if not candidate.is_absolute():
            candidate = profile_path.parent / candidate
        if candidate.exists() and candidate.is_file():
            icon_source = candidate
            icon_member = ICON_PREFIX + candidate.suffix.lower()
            profile["icon"] = icon_member
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(profile_member, json.dumps(profile, indent=2, ensure_ascii=False))
        zf.writestr(ABBREVIATIONS_FILENAME, json.dumps(abbreviations, indent=2, ensure_ascii=False))
        if icon_source is not None and icon_member is not None:
            zf.write(icon_source, icon_member)
    return zip_path


def _safe_json_members(zf: zipfile.ZipFile) -> list[str]:
    names: list[str] = []
    for info in zf.infolist():
        if info.is_dir():
            continue
        p = Path(info.filename)
        # Pacchetto volutamente piatto: niente path traversal o sottocartelle.
        if len(p.parts) != 1 or p.name != info.filename:
            continue
        if p.suffix.lower() == ".json":
            names.append(info.filename)
    return names


def inspect_profile_package(zip_path: str | Path) -> tuple[dict[str, Any], list[dict[str, str]], str]:
    zip_path = Path(zip_path)
    with zipfile.ZipFile(zip_path, "r") as zf:
        json_members = _safe_json_members(zf)
        if ABBREVIATIONS_FILENAME not in json_members:
            raise ValueError(f"Nel pacchetto manca {ABBREVIATIONS_FILENAME}")
        profile_members = [n for n in json_members if n != ABBREVIATIONS_FILENAME]
        if len(profile_members) != 1:
            raise ValueError("Il pacchetto deve contenere esattamente un JSON di profilo")
        try:
            profile = json.loads(zf.read(profile_members[0]).decode("utf-8-sig"))
            abbreviations_raw = json.loads(zf.read(ABBREVIATIONS_FILENAME).decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError(f"JSON non valido nel pacchetto: {exc}") from exc
    if not isinstance(profile, dict):
        raise ValueError("Il JSON del profilo non contiene un oggetto valido")
    if not profile.get("name"):
        raise ValueError("Il profilo importato non contiene il nome")
    abbreviations = normalize_regex_abbreviations(abbreviations_raw)
    icon_member = str(profile.get("icon", "") or "").strip()
    if icon_member:
        p = Path(icon_member)
        if len(p.parts) != 1 or p.name != icon_member or p.suffix.lower() not in {".png",".jpg",".jpeg",".gif",".webp",".ico"}:
            profile["icon"] = ""
    return profile, abbreviations, profile_members[0]


def choose_import_path(profiles_dir: str | Path, profile: dict[str, Any]) -> Path:
    profiles_dir = Path(profiles_dir)
    base = slugify(str(profile.get("name") or "profilo"), "profilo")
    path = profiles_dir / f"{base}.json"
    n = 2
    while path.exists():
        path = profiles_dir / f"{base}_{n}.json"
        n += 1
    return path


def merge_regex_abbreviations(
    current: Iterable[dict[str, Any]], imported: Iterable[dict[str, Any]]
) -> list[dict[str, str]]:
    """Unisce le abbreviazioni; quelle del pacchetto vincono a parità di token.

    In questo modo il profilo importato conserva la semantica con cui è stato
    esportato, senza eliminare abbreviazioni locali non presenti nel pacchetto.
    """
    merged: dict[str, dict[str, str]] = {}
    order: list[str] = []
    for item in normalize_regex_abbreviations(list(current)):
        token = item["token"]
        if token not in merged:
            order.append(token)
        merged[token] = dict(item)
    for item in normalize_regex_abbreviations(list(imported)):
        token = item["token"]
        if token not in merged:
            order.append(token)
        merged[token] = dict(item)
    return [merged[t] for t in order]


def import_profile_package(
    zip_path: str | Path,
    profile_manager: ProfileManager,
    *,
    target_name: str | None = None,
    overwrite_path: str | Path | None = None,
) -> tuple[Path, list[dict[str, str]]]:
    profile, abbreviations, _member = inspect_profile_package(zip_path)
    if target_name:
        profile["name"] = str(target_name).strip()
    if overwrite_path is not None:
        dest = Path(overwrite_path)
    elif target_name:
        dest = profile_manager.profiles_dir / f"{slugify(profile['name'], 'profilo')}.json"
        if dest.exists():
            raise FileExistsError(dest)
    else:
        dest = choose_import_path(profile_manager.profiles_dir, profile)
    icon_member = str(profile.get("icon", "") or "").strip()
    if icon_member:
        with zipfile.ZipFile(zip_path, "r") as zf:
            if icon_member in zf.namelist():
                suffix = Path(icon_member).suffix.lower()
                icon_dest = dest.with_name(dest.stem + "_icon" + suffix)
                icon_dest.write_bytes(zf.read(icon_member))
                profile["icon"] = icon_dest.name
            else:
                profile["icon"] = ""
    profile_manager.save(dest, profile)
    return dest, abbreviations


def export_profile_backup(profile_path: str | Path, destination_dir: str | Path, regex_abbreviations, states_by_email: dict[str, Any], excel_files: list[tuple[str, str | Path]]) -> Path:
    """Backup completo: profilo, regex, stato per email e tutti gli Excel generali."""
    profile_path = Path(profile_path); destination_dir = Path(destination_dir); destination_dir.mkdir(parents=True, exist_ok=True)
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    base = safe_package_name(str(profile.get("name") or profile_path.stem), profile_path.stem)
    out = destination_dir / f"{base}_backup.zip"
    n=2
    while out.exists():
        out = destination_dir / f"{base}_backup_{n}.zip"; n += 1
    icon_source=None; icon_member=None
    icon_value=str(profile.get("icon","") or "").strip()
    if icon_value:
        c=Path(icon_value); c=c if c.is_absolute() else profile_path.parent/c
        if c.exists() and c.is_file():
            icon_source=c; icon_member=ICON_PREFIX+c.suffix.lower(); profile["icon"]=icon_member
    with zipfile.ZipFile(out,"w",compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{safe_package_name(profile_path.stem,'profile')}.json", json.dumps(profile,indent=2,ensure_ascii=False))
        zf.writestr(ABBREVIATIONS_FILENAME,json.dumps(normalize_regex_abbreviations(list(regex_abbreviations)),indent=2,ensure_ascii=False))
        zf.writestr(BACKUP_STATE_FILENAME,json.dumps(states_by_email or {},indent=2,ensure_ascii=False))
        zf.writestr(BACKUP_META_FILENAME,json.dumps({"format":"EgoMailExtractor profile backup","version":1},indent=2))
        if icon_source and icon_member: zf.write(icon_source,icon_member)
        for arcname,path in excel_files:
            pp=Path(path)
            if pp.exists() and pp.is_file(): zf.write(pp, f"excel/{safe_package_name(arcname,pp.name)}")
    return out

def inspect_profile_backup(zip_path: str | Path) -> tuple[dict[str,Any], list[dict[str,str]], dict[str,Any], str]:
    with zipfile.ZipFile(zip_path,"r") as zf:
        if BACKUP_STATE_FILENAME not in zf.namelist() or ABBREVIATIONS_FILENAME not in zf.namelist():
            raise ValueError("Il file non è un backup completo di EgoMailExtractor")
        json_members=[n for n in _safe_json_members(zf) if n not in {ABBREVIATIONS_FILENAME,BACKUP_STATE_FILENAME,BACKUP_META_FILENAME}]
        if len(json_members)!=1: raise ValueError("Backup non valido: JSON profilo mancante o ambiguo")
        profile=json.loads(zf.read(json_members[0]).decode("utf-8-sig"))
        abbr=normalize_regex_abbreviations(json.loads(zf.read(ABBREVIATIONS_FILENAME).decode("utf-8-sig")))
        states=json.loads(zf.read(BACKUP_STATE_FILENAME).decode("utf-8-sig"))
    return profile,abbr,states,json_members[0]


def export_profiles_backup(
    profiles: list[tuple[str | Path, dict[str, Any], list[tuple[str, str | Path]]]],
    destination_dir: str | Path,
    regex_abbreviations: Iterable[dict[str, Any]],
) -> Path:
    """Crea un unico backup contenente uno o più profili, stati per email, icone ed Excel."""
    destination_dir = Path(destination_dir)
    destination_dir.mkdir(parents=True, exist_ok=True)
    out = destination_dir / "EgoMailExtractor_backup_profili.zip"
    n = 2
    while out.exists():
        out = destination_dir / f"EgoMailExtractor_backup_profili_{n}.zip"; n += 1
    meta_profiles=[]
    with zipfile.ZipFile(out,"w",compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(ABBREVIATIONS_FILENAME, json.dumps(normalize_regex_abbreviations(list(regex_abbreviations)), indent=2, ensure_ascii=False))
        for profile_path, states, excel_files in profiles:
            profile_path=Path(profile_path)
            profile=json.loads(profile_path.read_text(encoding="utf-8"))
            display=str(profile.get("name") or profile_path.stem)
            key=safe_package_name(profile_path.stem,"profile")
            prefix=f"profiles/{key}/"
            icon_value=str(profile.get("icon","") or "").strip()
            icon_member=""
            if icon_value:
                c=Path(icon_value); c=c if c.is_absolute() else profile_path.parent/c
                if c.exists() and c.is_file():
                    icon_member=prefix+"profile_icon"+c.suffix.lower()
                    zf.write(c,icon_member)
                    profile["icon"]=Path(icon_member).name
            profile_member=prefix+"profile.json"
            state_member=prefix+BACKUP_STATE_FILENAME
            zf.writestr(profile_member,json.dumps(profile,indent=2,ensure_ascii=False))
            zf.writestr(state_member,json.dumps(states or {},indent=2,ensure_ascii=False))
            excel_members=[]
            for arcname,path in excel_files:
                pp=Path(path)
                if pp.exists() and pp.is_file():
                    member=prefix+"excel/"+safe_package_name(arcname,pp.name)
                    zf.write(pp,member); excel_members.append(member)
            meta_profiles.append({"name":display,"key":key,"profile_member":profile_member,"state_member":state_member,"icon_member":icon_member,"excel_members":excel_members})
        zf.writestr(BACKUP_META_FILENAME,json.dumps({"format":"EgoMailExtractor profiles backup","version":2,"profiles":meta_profiles},indent=2,ensure_ascii=False))
    return out


def inspect_profiles_backup(zip_path: str | Path) -> tuple[list[dict[str,Any]], list[dict[str,str]]]:
    """Restituisce metadati dei profili contenuti e abbreviazioni del backup multi-profilo."""
    with zipfile.ZipFile(zip_path,"r") as zf:
        names=set(zf.namelist())
        if BACKUP_META_FILENAME not in names or ABBREVIATIONS_FILENAME not in names:
            # Compatibilità con backup 1.43 a profilo singolo.
            profile,abbr,states,member=inspect_profile_backup(zip_path)
            return [{"name":str(profile.get("name") or "profilo"),"key":"legacy","profile_member":member,"state_member":BACKUP_STATE_FILENAME,"legacy_profile":profile,"legacy_states":states,"excel_members":[n for n in names if n.startswith("excel/")]}],abbr
        meta=json.loads(zf.read(BACKUP_META_FILENAME).decode("utf-8-sig"))
        if meta.get("format")!="EgoMailExtractor profiles backup":
            profile,abbr,states,member=inspect_profile_backup(zip_path)
            return [{"name":str(profile.get("name") or "profilo"),"key":"legacy","profile_member":member,"state_member":BACKUP_STATE_FILENAME,"legacy_profile":profile,"legacy_states":states,"excel_members":[n for n in names if n.startswith("excel/")]}],abbr
        profiles=list(meta.get("profiles") or [])
        abbr=normalize_regex_abbreviations(json.loads(zf.read(ABBREVIATIONS_FILENAME).decode("utf-8-sig")))
        return profiles,abbr


def read_profile_from_backup(zip_path: str | Path, entry: dict[str,Any]) -> tuple[dict[str,Any],dict[str,Any]]:
    with zipfile.ZipFile(zip_path,"r") as zf:
        if entry.get("legacy_profile") is not None:
            return dict(entry["legacy_profile"]),dict(entry.get("legacy_states") or {})
        profile=json.loads(zf.read(entry["profile_member"]).decode("utf-8-sig"))
        states=json.loads(zf.read(entry["state_member"]).decode("utf-8-sig"))
        return profile,states
