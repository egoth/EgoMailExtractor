from __future__ import annotations

import json
import os
import shutil
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .crypto_store import EncryptedJsonStore
from .regex_utils import normalize_regex_abbreviations

from .edition import CONFIG_APP_NAME

APP_NAME = CONFIG_APP_NAME
CREDENTIALS_FILENAME = "imap_credentials.enc"
LEGACY_KEYRING_SERVICE = "EgoMailExtractor.IMAP"

DEFAULT_REGEX_ABBREVIATIONS = [
    {"token": "--ALFANUM--", "replacement": r"[^\W_]+"},
    {"token": "--ALFANUMSP--", "replacement": r"[^\W_\ ]+"},
    {"token": "--ALFA--", "replacement": r"[^\W\d_]+"},
    {"token": "--ALFASP--", "replacement": r"[^\W\d_\ ]+"},
    {"token": "--NOME--", "replacement": r"[^\W\d_]+(?:[\s'’\-]+[^\W\d_]+)*"},
    {"token": "--NUMERO--", "replacement": r"[+-]?\d+(?:[.,]\d+)?"},
    {"token": "--INT--", "replacement": r"[+-]?\d+"},
    {"token": "--DATA_TESTO--", "replacement": r"\d{1,2}\s+[^\W\d_]+(?:\s+\d{4})?"},
    {"token": "--DATA_NUMERICA--", "replacement": r"\d{1,2}[./-]\d{1,2}[./-]\d{4}"},
    {"token": "--DATA_MAIL--", "replacement": r"(?:lun|mar|mer|gio|ven|sab|dom)?\s*\d{1,2}\s+[^\W\d_]+(?:\s+\d{4})?"},
    {"token": "--DATA_TESTO_GENERICA--", "replacement": r"(?:lun|mar|mer|gio|ven|sab|dom)?\s*\d{1,2}\s+[^\W\d_]+(?:\s+\d{4})?"},
    {"token": "--QUALSIASI--", "replacement": r".*?"},
    {"token": "--EURO--", "replacement": r"--NUMERO-- €"},
    {"token": "--STESSA_RIGA--", "replacement": r"[^\r\n]+"},
]

def _merge_default_regex_abbreviations(value):
    current = normalize_regex_abbreviations(value)
    present = {str(x.get("token", "")) for x in current}
    # Le abbreviazioni predefinite sono sempre disponibili senza sovrascrivere
    # eventuali personalizzazioni già salvate dall'utente.
    return current + [dict(x) for x in DEFAULT_REGEX_ABBREVIATIONS if x["token"] not in present]


def sys_platform() -> str:
    import sys
    return sys.platform


def user_config_dir() -> Path:
    """Directory di configurazione secondo le convenzioni del sistema operativo."""
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    elif sys_platform() == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / APP_NAME


@dataclass
class ImapSettings:
    host: str = ""
    port: int = 993
    ssl: bool = True
    username: str = ""
    folder: str = "INBOX"


class ConfigManager:
    """Preferenze non sensibili + profili IMAP cifrati.

    ``settings.json`` non contiene host, username o password IMAP. Tutti i profili
    IMAP, compresa la password di ciascun account, sono mantenuti dentro
    ``imap_credentials.enc`` e sono disponibili solo dopo lo sblocco con la
    password master.

    Dalla v1.12 il file cifrato può contenere più profili IMAP. Le API storiche
    ``imap`` e ``get_password()`` continuano a riferirsi al profilo IMAP attivo,
    così il resto dell'applicazione rimane compatibile.
    """

    def __init__(self, app_dir: Path | None = None):
        self.app_dir = Path(app_dir) if app_dir else user_config_dir()
        self.app_dir.mkdir(parents=True, exist_ok=True)
        self.path = self.app_dir / "settings.json"
        self.credentials_path = self.app_dir / CREDENTIALS_FILENAME
        self.credentials_store = EncryptedJsonStore(self.credentials_path)
        self.data: dict[str, Any] = {}
        self._legacy_imap: dict[str, Any] | None = None
        self._master_password: str | None = None
        self._imap_profiles: list[dict[str, Any]] = []
        self._active_imap_profile_id = ""
        self._default_imap_profile_id = ""
        self._imap = ImapSettings()
        self._imap_password = ""
        self.load()

    @property
    def credentials_initialized(self) -> bool:
        return self.credentials_store.exists

    @property
    def is_unlocked(self) -> bool:
        return self._master_password is not None

    def default_profiles_dir(self) -> Path:
        return self.app_dir / "profiles"

    @property
    def work_dir(self) -> Path:
        return Path(self.data.get("work_dir", self.data.get("profiles_dir", self.default_profiles_dir())))

    @property
    def profiles_dir(self) -> Path:
        return Path(self.data.get("profiles_dir", self.default_profiles_dir()))

    @property
    def email_dir(self) -> Path:
        return Path(self.data.get("email_dir", self.app_dir))

    def ensure_email_dir(self) -> Path:
        path = self.email_dir
        path.mkdir(parents=True, exist_ok=True)
        return path

    def load(self) -> None:
        defaults = {
            "profiles_dir": str(self.default_profiles_dir()),
            "work_dir": str(self.default_profiles_dir()),
            "email_dir": str(self.app_dir),
            "window": {"geometry": "1450x900"},
            "regex_abbreviations": [dict(x) for x in DEFAULT_REGEX_ABBREVIATIONS],
            "regex_defaults_version": 1,
            "last_excel_by_profile": {},
            "profile_associations_by_imap": {},
            "extraction_state_by_imap": {},
        }
        if self.path.exists():
            try:
                loaded = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    # Migrazione della v1: non rimettiamo mai IMAP nel nuovo settings.json.
                    legacy = loaded.get("imap")
                    if isinstance(legacy, dict):
                        self._legacy_imap = dict(legacy)
                    for key in ("profiles_dir", "work_dir", "email_dir", "window", "regex_abbreviations", "regex_defaults_version", "last_excel_by_profile", "profile_associations_by_imap", "extraction_state_by_imap"):
                        if key in loaded:
                            defaults[key] = loaded[key]
                    # Prima della v1.35 gli Excel generali erano accanto ai profili:
                    # se non esiste ancora la preferenza manteniamo quel comportamento.
                    if "work_dir" not in loaded:
                        defaults["work_dir"] = defaults["profiles_dir"]
            except Exception:
                pass
        if self.path.exists() and int(defaults.get("regex_defaults_version", 0) or 0) < 1:
            defaults["regex_abbreviations"] = _merge_default_regex_abbreviations(defaults.get("regex_abbreviations", []))
            defaults["regex_defaults_version"] = 1
        else:
            defaults["regex_abbreviations"] = normalize_regex_abbreviations(defaults.get("regex_abbreviations", []))
        self.data = defaults
        self.ensure_profiles_dir()
        self.ensure_email_dir()
        # La posizione del file cifrato email è configurabile dalle Preferenze.
        self.credentials_path = self.email_dir / CREDENTIALS_FILENAME
        self.credentials_store = EncryptedJsonStore(self.credentials_path)

    def save(self) -> None:
        self.app_dir.mkdir(parents=True, exist_ok=True)
        clean = {
            "profiles_dir": self.data.get("profiles_dir", str(self.default_profiles_dir())),
            "work_dir": self.data.get("work_dir", self.data.get("profiles_dir", str(self.default_profiles_dir()))),
            "email_dir": self.data.get("email_dir", str(self.app_dir)),
            "window": self.data.get("window", {"geometry": "1450x900"}),
            "regex_abbreviations": normalize_regex_abbreviations(
                self.data.get("regex_abbreviations", [])
            ),
            "regex_defaults_version": int(self.data.get("regex_defaults_version", 1) or 1),
            "last_excel_by_profile": dict(self.data.get("last_excel_by_profile", {}) or {}),
            "profile_associations_by_imap": dict(self.data.get("profile_associations_by_imap", {}) or {}),
            "extraction_state_by_imap": dict(self.data.get("extraction_state_by_imap", {}) or {}),
        }
        self.data = clean
        self.path.write_text(json.dumps(clean, indent=2, ensure_ascii=False), encoding="utf-8")

    @staticmethod
    def _normalize_imap_profile(raw: dict[str, Any], index: int = 1) -> dict[str, Any]:
        settings_raw = raw.get("settings", raw)
        if not isinstance(settings_raw, dict):
            settings_raw = {}
        profile_id = str(raw.get("id") or uuid.uuid4().hex)
        name = str(raw.get("name") or settings_raw.get("username") or settings_raw.get("host") or f"Email {index}").strip()
        return {
            "id": profile_id,
            "name": name or f"Email {index}",
            "settings": {
                "host": str(settings_raw.get("host", "")),
                "port": int(settings_raw.get("port", 993) or 993),
                "ssl": bool(settings_raw.get("ssl", True)),
                "username": str(settings_raw.get("username", "")),
                "folder": str(settings_raw.get("folder", "INBOX") or "INBOX"),
            },
            "password": str(raw.get("password", settings_raw.get("password", "")) or ""),
        }

    def _sync_active_cache(self) -> None:
        if not self._imap_profiles:
            self._active_imap_profile_id = ""
            self._default_imap_profile_id = ""
            self._imap = ImapSettings()
            self._imap_password = ""
            return
        ids = {p["id"] for p in self._imap_profiles}
        if self._default_imap_profile_id not in ids:
            self._default_imap_profile_id = self._imap_profiles[0]["id"]
        if self._active_imap_profile_id not in ids:
            # Un solo profilo viene sempre aperto automaticamente; con più profili
            # il default è la scelta iniziale sicura fino all'eventuale chooser.
            self._active_imap_profile_id = (
                self._imap_profiles[0]["id"] if len(self._imap_profiles) == 1
                else self._default_imap_profile_id
            )
        entry = next((p for p in self._imap_profiles if p["id"] == self._active_imap_profile_id), self._imap_profiles[0])
        raw = entry["settings"]
        self._imap = ImapSettings(
            host=str(raw.get("host", "")),
            port=int(raw.get("port", 993)),
            ssl=bool(raw.get("ssl", True)),
            username=str(raw.get("username", "")),
            folder=str(raw.get("folder", "INBOX")),
        )
        self._imap_password = str(entry.get("password", ""))

    def unlock_credentials(self, password: str, create_if_missing: bool = False) -> None:
        if self.credentials_store.exists:
            payload = self.credentials_store.load(password)
            self._master_password = password
            migrated_single = False
            profiles: list[dict[str, Any]] = []
            if isinstance(payload, dict) and isinstance(payload.get("imap_profiles"), list):
                for i, raw in enumerate(payload.get("imap_profiles", []), 1):
                    if isinstance(raw, dict):
                        profiles.append(self._normalize_imap_profile(raw, i))
                self._active_imap_profile_id = str(payload.get("active_imap_profile_id", "") or "")
                self._default_imap_profile_id = str(payload.get("default_imap_profile_id", "") or "")
            else:
                # Migrazione trasparente del formato cifrato v1 con singolo ``imap``.
                raw = payload.get("imap", {}) if isinstance(payload, dict) else {}
                if isinstance(raw, dict) and any(str(raw.get(k, "")).strip() for k in ("host", "username", "password")):
                    entry = self._normalize_imap_profile({
                        "name": str(raw.get("username") or raw.get("host") or "Predefinito"),
                        **raw,
                    })
                    profiles = [entry]
                    self._active_imap_profile_id = entry["id"]
                    self._default_imap_profile_id = entry["id"]
                    migrated_single = True
            self._imap_profiles = profiles
            self._sync_active_cache()
            if migrated_single:
                self._save_credentials()
            return
        if not create_if_missing:
            raise FileNotFoundError(self.credentials_path)
        self._master_password = password
        self._imap_profiles = []
        self._active_imap_profile_id = ""
        self._default_imap_profile_id = ""
        self._sync_active_cache()
        self._save_credentials()

    def _save_credentials(self) -> None:
        if self._master_password is None:
            raise RuntimeError("Credenziali bloccate: inserire prima la password master")
        payload = {
            "schema_version": 2,
            "imap_profiles": self._imap_profiles,
            "active_imap_profile_id": self._active_imap_profile_id,
            "default_imap_profile_id": self._default_imap_profile_id,
        }
        self.credentials_store.save(self._master_password, payload)

    def migrate_legacy_credentials(self) -> bool:
        """Importa una configurazione v1 da settings.json/keyring come profilo IMAP."""
        if not self._legacy_imap or not self.is_unlocked:
            self.save()
            return False
        raw = self._legacy_imap
        settings = ImapSettings(
            host=str(raw.get("host", "")),
            port=int(raw.get("port", 993)),
            ssl=bool(raw.get("ssl", True)),
            username=str(raw.get("username", "")),
            folder=str(raw.get("folder", "INBOX")),
        )
        password = ""
        if settings.username:
            try:
                import keyring  # type: ignore
                password = keyring.get_password(LEGACY_KEYRING_SERVICE, settings.username) or ""
                if password:
                    try:
                        keyring.delete_password(LEGACY_KEYRING_SERVICE, settings.username)
                    except Exception:
                        pass
            except Exception:
                pass
        self.create_imap_profile(
            name=settings.username or settings.host or "Predefinito",
            settings=settings,
            password=password,
            make_default=(len(self._imap_profiles) == 0),
            make_active=True,
        )
        self._legacy_imap = None
        self.save()
        return True

    def change_master_password(self, new_password: str) -> None:
        if not new_password:
            raise ValueError("La nuova password master non può essere vuota")
        if not self.is_unlocked:
            raise RuntimeError("Credenziali non sbloccate")
        old = self._master_password
        try:
            self._master_password = new_password
            self._save_credentials()
        except Exception:
            self._master_password = old
            raise

    def lock_credentials(self) -> None:
        self._master_password = None
        self._imap_profiles = []
        self._active_imap_profile_id = ""
        self._default_imap_profile_id = ""
        self._imap = ImapSettings()
        self._imap_password = ""

    @property
    def profiles_dir(self) -> Path:
        return Path(self.data.get("profiles_dir") or self.default_profiles_dir())

    @property
    def work_dir(self) -> Path:
        """Cartella di lavoro che contiene gli Excel generali dei profili."""
        value = self.data.get("work_dir") or self.data.get("profiles_dir") or self.default_profiles_dir()
        path = Path(value)
        path.mkdir(parents=True, exist_ok=True)
        return path

    def set_profiles_dir(self, path: str | Path) -> None:
        self.data["profiles_dir"] = str(Path(path))
        self.ensure_profiles_dir()
        self.save()

    def set_email_dir(self, path: str | Path) -> None:
        new_dir = Path(path)
        new_dir.mkdir(parents=True, exist_ok=True)
        old_path = Path(self.credentials_path)
        new_path = new_dir / CREDENTIALS_FILENAME
        if old_path.exists() and old_path.resolve() != new_path.resolve() and not new_path.exists():
            shutil.copy2(old_path, new_path)
        self.data["email_dir"] = str(new_dir)
        self.credentials_path = new_path
        self.credentials_store = EncryptedJsonStore(self.credentials_path)
        self.save()

    def associated_profile_names(self, imap_profile_id: str | None = None) -> list[str] | None:
        """Nomi file dei profili di estrazione associati a una email IMAP.

        ``None`` significa che per quell'email non è stata ancora fatta una scelta:
        per compatibilità/migrazione vengono quindi considerati associati tutti i profili.
        """
        pid = str(imap_profile_id or self._active_imap_profile_id or "")
        if not pid:
            return None
        mapping = self.data.get("profile_associations_by_imap", {}) or {}
        if not isinstance(mapping, dict) or pid not in mapping:
            return None
        value = mapping.get(pid, [])
        return [str(x) for x in value if str(x).strip()] if isinstance(value, list) else []

    def set_associated_profile_names(self, names: list[str], imap_profile_id: str | None = None) -> None:
        pid = str(imap_profile_id or self._active_imap_profile_id or "")
        if not pid:
            raise ValueError("Nessuna email attiva")
        mapping = self.data.setdefault("profile_associations_by_imap", {})
        if not isinstance(mapping, dict):
            mapping = {}
            self.data["profile_associations_by_imap"] = mapping
        mapping[pid] = sorted({str(Path(x).name) for x in names if str(x).strip()}, key=str.lower)
        self.save()

    def rename_associated_profile(self, old_name: str, new_name: str) -> None:
        mapping = self.data.get("profile_associations_by_imap", {}) or {}
        if not isinstance(mapping, dict):
            return
        changed = False
        for pid, values in list(mapping.items()):
            if not isinstance(values, list):
                continue
            updated = [new_name if str(x) == old_name else str(x) for x in values]
            if updated != values:
                mapping[pid] = updated
                changed = True
        if changed:
            self.save()

    @property
    def regex_abbreviations(self) -> list[dict[str, str]]:
        return normalize_regex_abbreviations(self.data.get("regex_abbreviations", []))

    def set_regex_abbreviations(self, items: list[dict[str, str]]) -> None:
        self.data["regex_abbreviations"] = normalize_regex_abbreviations(items)
        self.save()

    # ------------------------------------------------------------------
    # Profili IMAP cifrati
    # ------------------------------------------------------------------
    def extraction_state(self, profile_filename: str, imap_profile_id: str | None = None) -> dict[str, Any]:
        """Stato operativo separato per coppia email + profilo di estrazione."""
        pid = str(imap_profile_id or self._active_imap_profile_id or "default")
        profile_key = str(Path(profile_filename).name)
        root = self.data.setdefault("extraction_state_by_imap", {})
        if not isinstance(root, dict):
            root = {}
            self.data["extraction_state_by_imap"] = root
        by_profile = root.setdefault(pid, {})
        if not isinstance(by_profile, dict):
            by_profile = {}
            root[pid] = by_profile
        state = by_profile.setdefault(profile_key, {})
        if not isinstance(state, dict):
            state = {}
            by_profile[profile_key] = state
        state.setdefault("folder", "INBOX")
        state.setdefault("uidvalidity", 0)
        state.setdefault("last_processed_uid", 0)
        state.setdefault("last_processed_at", "")
        state.setdefault("first_processed_uid", 0)
        state.setdefault("first_processed_at", "")
        state.setdefault("extraction_start_date", "")
        state.setdefault("home_summary_values", [])
        return state

    def all_extraction_states_for_profile(self, profile_filename: str) -> dict[str, dict[str, Any]]:
        profile_key = str(Path(profile_filename).name)
        root = self.data.get("extraction_state_by_imap", {}) or {}
        out = {}
        if isinstance(root, dict):
            for pid, mapping in root.items():
                if isinstance(mapping, dict) and isinstance(mapping.get(profile_key), dict):
                    out[str(pid)] = dict(mapping[profile_key])
        return out

    def set_all_extraction_states_for_profile(self, profile_filename: str, states: dict[str, Any]) -> None:
        profile_key = str(Path(profile_filename).name)
        root = self.data.setdefault("extraction_state_by_imap", {})
        if not isinstance(root, dict):
            root = {}; self.data["extraction_state_by_imap"] = root
        # Rimuove gli stati esistenti di questo profilo, poi ripristina quelli ricevuti.
        for pid, mapping in list(root.items()):
            if isinstance(mapping, dict):
                mapping.pop(profile_key, None)
        for pid, state in (states or {}).items():
            mapping = root.setdefault(str(pid), {})
            if isinstance(mapping, dict) and isinstance(state, dict):
                mapping[profile_key] = dict(state)
        self.save()

    def rename_extraction_profile_state(self, old_name: str, new_name: str) -> None:
        old_key, new_key = str(Path(old_name).name), str(Path(new_name).name)
        root = self.data.get("extraction_state_by_imap", {}) or {}
        changed = False
        if isinstance(root, dict):
            for mapping in root.values():
                if isinstance(mapping, dict) and old_key in mapping:
                    mapping[new_key] = mapping.pop(old_key)
                    changed = True
        if changed:
            self.save()

    def remove_extraction_profile_state(self, profile_name: str) -> None:
        key = str(Path(profile_name).name)
        root = self.data.get("extraction_state_by_imap", {}) or {}
        changed = False
        if isinstance(root, dict):
            for mapping in root.values():
                if isinstance(mapping, dict) and key in mapping:
                    mapping.pop(key, None); changed = True
        if changed:
            self.save()

    def duplicate_extraction_profile_state(self, old_name: str, new_name: str) -> None:
        import copy
        old_key, new_key = str(Path(old_name).name), str(Path(new_name).name)
        root = self.data.setdefault("extraction_state_by_imap", {})
        if not isinstance(root, dict):
            return
        changed = False
        for mapping in root.values():
            if isinstance(mapping, dict) and isinstance(mapping.get(old_key), dict):
                mapping[new_key] = copy.deepcopy(mapping[old_key]); changed = True
        if changed:
            self.save()

    def list_imap_profiles(self) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for p in self._imap_profiles:
            s = p.get("settings", {})
            out.append({
                "id": p.get("id", ""),
                "name": p.get("name", ""),
                "host": s.get("host", ""),
                "port": s.get("port", 993),
                "ssl": s.get("ssl", True),
                "username": s.get("username", ""),
                "folder": s.get("folder", "INBOX"),
                "is_default": p.get("id") == self._default_imap_profile_id,
                "is_active": p.get("id") == self._active_imap_profile_id,
            })
        return out

    @property
    def imap_profile_count(self) -> int:
        return len(self._imap_profiles)

    @property
    def active_imap_profile_id(self) -> str:
        return self._active_imap_profile_id

    @property
    def default_imap_profile_id(self) -> str:
        return self._default_imap_profile_id

    @property
    def active_imap_profile_name(self) -> str:
        entry = next((p for p in self._imap_profiles if p.get("id") == self._active_imap_profile_id), None)
        return str(entry.get("name", "")) if entry else ""

    def get_imap_profile(self, profile_id: str) -> tuple[str, ImapSettings, str] | None:
        entry = next((p for p in self._imap_profiles if p.get("id") == profile_id), None)
        if not entry:
            return None
        raw = entry.get("settings", {})
        settings = ImapSettings(
            host=str(raw.get("host", "")), port=int(raw.get("port", 993)),
            ssl=bool(raw.get("ssl", True)), username=str(raw.get("username", "")),
            folder=str(raw.get("folder", "INBOX") or "INBOX"),
        )
        return str(entry.get("name", "")), settings, str(entry.get("password", ""))

    def create_imap_profile(
        self,
        name: str,
        settings: ImapSettings,
        password: str,
        *,
        make_default: bool = False,
        make_active: bool = True,
    ) -> str:
        if not self.is_unlocked:
            raise RuntimeError("Credenziali email bloccate")
        profile_id = uuid.uuid4().hex
        clean_name = str(name or settings.username or settings.host or f"Email {len(self._imap_profiles)+1}").strip()
        entry = {
            "id": profile_id,
            "name": clean_name,
            "settings": asdict(settings),
            "password": str(password or ""),
        }
        self._imap_profiles.append(entry)
        if len(self._imap_profiles) == 1 or make_default:
            self._default_imap_profile_id = profile_id
        if make_active or len(self._imap_profiles) == 1:
            self._active_imap_profile_id = profile_id
        self._sync_active_cache()
        self._save_credentials()
        return profile_id

    def update_imap_profile(
        self,
        profile_id: str,
        *,
        name: str,
        settings: ImapSettings,
        password: str,
    ) -> None:
        entry = next((p for p in self._imap_profiles if p.get("id") == profile_id), None)
        if entry is None:
            raise ValueError("Email non trovata")
        entry["name"] = str(name or settings.username or settings.host or "Email").strip()
        entry["settings"] = asdict(settings)
        entry["password"] = str(password or "")
        self._sync_active_cache()
        self._save_credentials()

    def delete_imap_profile(self, profile_id: str) -> bool:
        before = len(self._imap_profiles)
        self._imap_profiles = [p for p in self._imap_profiles if p.get("id") != profile_id]
        if len(self._imap_profiles) == before:
            return False
        if self._default_imap_profile_id == profile_id:
            self._default_imap_profile_id = self._imap_profiles[0]["id"] if self._imap_profiles else ""
        if self._active_imap_profile_id == profile_id:
            self._active_imap_profile_id = self._default_imap_profile_id or (self._imap_profiles[0]["id"] if self._imap_profiles else "")
        self._sync_active_cache()
        self._save_credentials()
        return True

    def set_active_imap_profile(self, profile_id: str, *, persist: bool = True) -> None:
        if not any(p.get("id") == profile_id for p in self._imap_profiles):
            raise ValueError("Email non trovata")
        self._active_imap_profile_id = profile_id
        self._sync_active_cache()
        if persist:
            self._save_credentials()

    def set_default_imap_profile(self, profile_id: str) -> None:
        if not any(p.get("id") == profile_id for p in self._imap_profiles):
            raise ValueError("Email non trovata")
        self._default_imap_profile_id = profile_id
        self._save_credentials()

    @property
    def imap(self) -> ImapSettings:
        return self._imap

    # API compatibili con le versioni precedenti: lavorano sul profilo attivo.
    def set_imap(self, settings: ImapSettings) -> None:
        self.set_imap_credentials(settings, self._imap_password)

    def set_imap_credentials(self, settings: ImapSettings, password: str) -> None:
        if self._active_imap_profile_id:
            name = self.active_imap_profile_name or settings.username or settings.host or "Predefinito"
            self.update_imap_profile(
                self._active_imap_profile_id, name=name, settings=settings, password=password
            )
        else:
            self.create_imap_profile(
                name=settings.username or settings.host or "Predefinito",
                settings=settings,
                password=password,
                make_default=True,
                make_active=True,
            )

    def get_password(self) -> str | None:
        if not self.is_unlocked:
            raise RuntimeError("Credenziali email bloccate")
        return self._imap_password or None

    def set_password(self, password: str) -> None:
        self.set_imap_credentials(self._imap, password)

    def clear_imap_credentials(self) -> None:
        """Svuota le credenziali del profilo IMAP attivo mantenendone il nome."""
        if self._active_imap_profile_id:
            self.update_imap_profile(
                self._active_imap_profile_id,
                name=self.active_imap_profile_name or "Email",
                settings=ImapSettings(),
                password="",
            )
        else:
            self._imap = ImapSettings()
            self._imap_password = ""
            self._save_credentials()

    def delete_password(self) -> None:
        self.set_password("")

    def ensure_profiles_dir(self) -> None:
        self.profiles_dir.mkdir(parents=True, exist_ok=True)

    def install_builtin_profiles(self, builtin_dir: Path) -> None:
        self.ensure_profiles_dir()
        for src in builtin_dir.glob("*.json"):
            dst = self.profiles_dir / src.name
            if not dst.exists():
                shutil.copy2(src, dst)
