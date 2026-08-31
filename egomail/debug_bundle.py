from __future__ import annotations

import json
import platform
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any


def build_debug_bundle(
    zip_path: str | Path,
    *,
    profile: dict[str, Any],
    profile_path: str | Path,
    log_path: str | Path | None,
    excel_path: str | Path | None,
    parameters: dict[str, Any],
    preferences_path: str | Path | None = None,
) -> dict[str, Any]:
    """Crea un pacchetto ZIP autosufficiente per il debug utente.

    Il profilo viene scritto dalla copia in memoria, così il pacchetto fotografa
    esattamente la configurazione usata anche se non è ancora stata salvata sul
    file JSON. I parametri devono essere già privi di segreti: questa funzione
    non legge né include il file cifrato delle credenziali.
    """
    destination = Path(zip_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    profile_path = Path(profile_path)
    log = Path(log_path) if log_path else None
    excel = Path(excel_path) if excel_path else None
    preferences = Path(preferences_path) if preferences_path else None

    included: list[str] = []
    missing: list[str] = []
    profile_arcname = profile_path.name or "profilo.json"
    params_arcname = "parametri_debug.json"
    readme_arcname = "LEGGIMI_DEBUG.txt"

    manifest: dict[str, Any] = {
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "profile_source_path": str(profile_path),
        "log_source_path": str(log) if log else "",
        "excel_source_path": str(excel) if excel else "",
        "preferences_source_path": str(preferences) if preferences else "",
        "included": included,
        "missing": missing,
    }

    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(profile_arcname, json.dumps(profile, indent=2, ensure_ascii=False, default=str))
        included.append(profile_arcname)

        zf.writestr(params_arcname, json.dumps(parameters, indent=2, ensure_ascii=False, default=str))
        included.append(params_arcname)

        if log and log.exists() and log.is_file():
            zf.write(log, arcname=log.name)
            included.append(log.name)
        else:
            missing.append("log estrazione")

        if excel and excel.exists() and excel.is_file():
            zf.write(excel, arcname=excel.name)
            included.append(excel.name)
        else:
            missing.append("ultimo file Excel salvato")

        if preferences is not None:
            if preferences.exists() and preferences.is_file():
                zf.write(preferences, arcname="preferenze/settings.json")
                included.append("preferenze/settings.json")
            else:
                missing.append("preferenze settings.json")

        readme = (
            "Pacchetto debug EgoMailExtractor\n"
            "================================\n\n"
            "Contiene il profilo effettivamente in uso, i parametri non sensibili, "
            "le preferenze, il log di estrazione e l'ultimo Excel salvato quando disponibili.\n\n"
            "Per sicurezza NON vengono inclusi la password master, la password IMAP, "
            "imap_credentials.enc o altri segreti.\n\n"
            f"Python: {sys.version.split()[0]}\n"
            f"Sistema: {platform.platform()}\n"
        )
        zf.writestr(readme_arcname, readme)
        included.append(readme_arcname)

        # Scriviamo il manifest per ultimo, dopo avere popolato included/missing.
        included.append("manifest_debug.json")
        zf.writestr("manifest_debug.json", json.dumps(manifest, indent=2, ensure_ascii=False))

    return manifest
