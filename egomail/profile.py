from __future__ import annotations

import copy
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


PROFILE_VERSION = 7


FIELD_TYPES = ("text", "date", "number")
MATCH_SOURCES = ("subject", "sender", "body")
MATCH_MODES = ("contains", "regex")
UPDATE_POLICIES = ("replace", "keep", "sum")
FINAL_FORMULAS = ("sum", "count")
EXTRACTION_SOURCES = ("body", "subject", "sender", "date", "message_id")


def normalize_field_type(value: str | None) -> str:
    """Converte i tipi delle versioni precedenti nei tre tipi Excel esposti."""
    value = (value or "text").strip().lower()
    if value in {"date", "datetime"}:
        return "date"
    if value in {"number", "int", "integer", "float", "money", "currency", "decimal"}:
        return "number"
    return "text"


def _migrate_criteria(criteria: dict[str, Any]) -> list[dict[str, Any]]:
    """Migra i vecchi ``criteria`` nelle nuove regole di selezione.

    Ogni regola compilata deve risultare vera (AND) affinché il tipo mail
    corrisponda. La lista permette più regole anche sulla stessa origine.
    """
    out: list[dict[str, Any]] = []
    mapping = (
        ("subject_regex", "subject", "regex"),
        ("sender_regex", "sender", "regex"),
        ("body_regex", "body", "regex"),
        ("subject_contains", "subject", "contains"),
        ("sender_contains", "sender", "contains"),
        ("body_contains", "body", "contains"),
    )
    for key, source, mode in mapping:
        value = criteria.get(key)
        if value is not None and str(value).strip():
            out.append({
                "name": f"match_{source}_{len(out)+1}",
                "source": source,
                "mode": mode,
                "trigger": str(value).strip(),
            })
    return out


def normalize_match_rules(mail_type: dict[str, Any]) -> list[dict[str, Any]]:
    raw = mail_type.get("match_rules")
    if raw is None:
        raw = _migrate_criteria(mail_type.get("criteria", {}) or {})
    normalized: list[dict[str, Any]] = []
    for i, item in enumerate(raw or [], 1):
        if not isinstance(item, dict):
            continue
        source = str(item.get("source", "body")).strip().lower()
        if source not in MATCH_SOURCES:
            source = "body"
        mode = str(item.get("mode", "contains")).strip().lower()
        if mode not in MATCH_MODES:
            mode = "contains"
        trigger = item.get("trigger", item.get("pattern", item.get("value", "")))
        normalized.append({
            "name": str(item.get("name") or f"match_{source}_{i}"),
            "source": source,
            "mode": mode,
            "trigger": str(trigger or ""),
        })
    mail_type["match_rules"] = normalized
    # Dalla v1.5 la fonte di verità è match_rules. Rimuoviamo criteria per
    # evitare che l'utente modifichi due strutture diverse per lo stesso match.
    mail_type.pop("criteria", None)
    mail_type.setdefault("rules", [])
    return normalized


def _merge_field_metadata(target: dict[str, Any], source: dict[str, Any]) -> dict[str, Any]:
    """Unisce due definizioni dello stesso campo senza creare duplicati.

    La definizione ``target`` è quella canonica: ne manteniamo nome e tipo.
    Le proprietà opzionali mancanti (formato, contatore, formula finale) vengono
    recuperate da ``source``. Questo rende sicura anche la migrazione di profili
    creati con versioni precedenti che potevano contenere due righe con lo
    stesso nome campo.
    """
    target_type = normalize_field_type(target.get("type"))
    source_type = normalize_field_type(source.get("type"))
    # Se una delle vecchie definizioni era rimasta al tipo generico Testo e
    # l'altra contiene un tipo più specifico, conserviamo quello specifico.
    # In caso di conflitto Numero/Data resta invece valida la prima definizione.
    if target_type == "text" and source_type != "text":
        target["type"] = source_type
    else:
        target["type"] = target_type
    for attr in ("format", "mail_type_counter", "final_formula", "summary_source_field"):
        if not target.get(attr) and source.get(attr):
            target[attr] = source[attr]
    return target


def _replace_field_references(profile: dict[str, Any], old_name: str, new_name: str) -> None:
    """Reindirizza tutti i riferimenti da un campo a un altro."""
    if not old_name or not new_name or old_name == new_name:
        return

    for mt in profile.get("mail_types", []) or []:
        if not isinstance(mt, dict):
            continue
        for rule in mt.get("rules", []) or []:
            if isinstance(rule, dict) and rule.get("field") == old_name:
                rule["field"] = new_name

    for comp in profile.get("computed_fields", []) or []:
        if not isinstance(comp, dict):
            continue
        for attr in ("field", "start_field", "end_field"):
            if comp.get(attr) == old_name:
                comp[attr] = new_name

    keys: list[str] = []
    for key in profile.get("reconcile_keys", []) or []:
        key = new_name if key == old_name else key
        if key not in keys:
            keys.append(key)
    profile["reconcile_keys"] = keys

    # Le formule riepilogative usano alias simbolici legati al nome campo.
    for sf in profile.get("summary_formulas", []) or []:
        if not isinstance(sf, dict):
            continue
        formula_text = str(sf.get("formula", "") or "")
        for prefix in ("somma_", "conta_"):
            formula_text = re.sub(
                rf"\b{re.escape(prefix + old_name)}\b",
                prefix + new_name, formula_text, flags=re.I,
            )
        sf["formula"] = formula_text


def merge_field_into(
    profile: dict[str, Any],
    source_name: str,
    target_name: str,
    *,
    source_index: int | None = None,
) -> dict[str, Any]:
    """Unisce ``source_name`` nel campo esistente ``target_name``.

    Tutte le regole di estrazione, i campi calcolati, le chiavi di
    riconciliazione e gli alias delle formule vengono reindirizzati al campo
    destinazione. La riga sorgente viene quindi rimossa da ``fields``.

    ``source_index`` permette alla GUI di distinguere due righe che, a causa
    del vecchio bug, hanno già lo stesso nome.
    """
    fields = profile.setdefault("fields", [])
    if source_index is None:
        source_index = next((
            i for i, f in enumerate(fields)
            if isinstance(f, dict) and str(f.get("name", "")) == source_name
        ), None)
    if source_index is None or source_index < 0 or source_index >= len(fields):
        raise ValueError(f"Campo sorgente '{source_name}' non trovato")

    source = fields[source_index]
    if not isinstance(source, dict):
        raise ValueError("Definizione campo sorgente non valida")
    actual_source_name = str(source.get("name", "") or source_name)

    target_index = next((
        i for i, f in enumerate(fields)
        if i != source_index and isinstance(f, dict) and str(f.get("name", "")) == target_name
    ), None)
    if target_index is None:
        raise ValueError(f"Campo destinazione '{target_name}' non trovato")

    target = fields[target_index]
    assert isinstance(target, dict)
    _merge_field_metadata(target, source)

    _replace_field_references(profile, actual_source_name, target_name)

    # Se i due nomi sono già uguali, i riferimenti sono già condivisi; resta
    # soltanto da eliminare la definizione duplicata.
    del fields[source_index]
    normalize_profile(profile)
    return get_field(profile, target_name) or target


def normalize_profile(profile: dict[str, Any]) -> dict[str, Any]:
    profile.setdefault("fields", [])
    normalized_fields: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in profile.get("fields", []):
        if isinstance(item, str):
            field = {"name": item, "type": "text"}
        elif isinstance(item, dict) and item.get("name"):
            field = dict(item)
            old_type = str(field.get("type", "text")).lower()
            field["type"] = normalize_field_type(old_type)
            if old_type in {"money", "currency"} and not field.get("format"):
                field["format"] = "currency_eur"
        else:
            continue
        formula = str(field.get("final_formula", "") or "").strip().lower()
        if formula not in FINAL_FORMULAS:
            field.pop("final_formula", None)
        elif formula == "sum" and field["type"] != "number":
            # La somma ha senso solo su colonne numeriche. Una configurazione
            # incoerente proveniente da un vecchio JSON viene disattivata.
            field.pop("final_formula", None)
        else:
            field["final_formula"] = formula
        name = str(field["name"])
        if name in seen:
            # Versioni precedenti permettevano di rinominare un campo con il
            # nome di uno già esistente, lasciando due definizioni distinte.
            # Da v1.28 le consideriamo lo stesso campo e ne fondiamo i metadati.
            existing = next((f for f in normalized_fields if str(f.get("name", "")) == name), None)
            if existing is not None:
                _merge_field_metadata(existing, field)
            continue
        seen.add(name)
        normalized_fields.append(field)
    profile["fields"] = normalized_fields

    profile.setdefault("mail_types", [])
    fields_by_name = {f.get("name"): f for f in normalized_fields if isinstance(f, dict)}
    # Un campo riepilogativo separato può calcolare SOMMA/CONTA prendendo i
    # valori da un altro campo del profilo. Se il riferimento non è valido,
    # torniamo alla colonna stessa per mantenere il profilo coerente.
    for field in normalized_fields:
        source = str(field.get("summary_source_field", "") or "").strip()
        if source and source in fields_by_name and source != field.get("name"):
            field["summary_source_field"] = source
        else:
            field.pop("summary_source_field", None)
    for mail_type in profile.get("mail_types", []):
        if isinstance(mail_type, dict):
            normalize_match_rules(mail_type)
            for rule in mail_type.setdefault("rules", []):
                if not isinstance(rule, dict):
                    continue
                policy = str(rule.get("existing_value_policy", "replace") or "replace").strip().lower()
                if policy not in UPDATE_POLICIES:
                    policy = "replace"
                field = fields_by_name.get(rule.get("field"), {})
                if policy == "sum" and normalize_field_type(field.get("type")) != "number":
                    policy = "replace"
                rule["existing_value_policy"] = policy

    # Un campo può essere un contatore automatico dei match di un Tipo mail.
    # Il contatore è sempre numerico e viene incrementato una sola volta per
    # ciascuna mail che fa match con quel Tipo mail sulla riga riconciliata.
    valid_mail_types = {
        str(mt.get("name")) for mt in profile.get("mail_types", [])
        if isinstance(mt, dict) and mt.get("name")
    }
    for field in normalized_fields:
        counter = str(field.get("mail_type_counter", "") or "").strip()
        if counter and counter in valid_mail_types:
            field["mail_type_counter"] = counter
            field["type"] = "number"
        else:
            field.pop("mail_type_counter", None)

    # Formule riepilogative nominate: formula composta oppure aggregazione
    # temporale condizionata da un campo data.
    summary_formulas: list[dict[str, Any]] = []
    seen_formula_names: set[str] = set()
    valid_field_names_for_summary = {
        str(f.get("name", "")) for f in normalized_fields if isinstance(f, dict) and f.get("name")
    }
    date_field_names_for_summary = {
        str(f.get("name", "")) for f in normalized_fields
        if isinstance(f, dict) and f.get("name") and normalize_field_type(f.get("type")) == "date"
    }
    for item in profile.get("summary_formulas", []) or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "") or "").strip()
        if not name or name.casefold() in seen_formula_names:
            continue
        if str(item.get("kind", "") or "") == "period_aggregate":
            operation = str(item.get("operation", "sum") or "sum").lower()
            field = str(item.get("field", "") or "").strip()
            date_field = str(item.get("date_field", "") or "").strip()
            period = str(item.get("period", "month") or "month").lower()
            period_mode = str(item.get("period_mode", "rolling") or "rolling").lower()
            if operation not in {"sum", "count"}:
                operation = "sum"
            if period not in {"month", "year"}:
                period = "month"
            if period_mode not in {"rolling", "calendar"}:
                period_mode = "rolling"
            if field not in valid_field_names_for_summary or date_field not in date_field_names_for_summary:
                continue
            normalized = {
                "name": name,
                "kind": "period_aggregate",
                "operation": operation,
                "field": field,
                "date_field": date_field,
                "period": period,
                "period_mode": period_mode,
            }
        else:
            formula = str(item.get("formula", "") or "").strip()
            if not formula:
                continue
            if not formula.startswith("="):
                formula = "=" + formula
            normalized = {"name": name, "formula": formula}
        seen_formula_names.add(name.casefold())
        summary_formulas.append(normalized)
    profile["summary_formulas"] = summary_formulas

    # Campi calcolati del profilo. Dalla v1.16 sono completamente gestibili
    # dalla GUI e possono scegliere se una regola esplicita della mail corrente
    # deve avere priorità sul ricalcolo automatico.
    computed_fields: list[dict[str, str]] = []
    seen_computed_destinations: set[str] = set()
    valid_field_names = {str(f.get("name", "")) for f in normalized_fields if isinstance(f, dict)}
    for item in profile.get("computed_fields", []) or []:
        if not isinstance(item, dict):
            continue
        field = str(item.get("field", "") or "").strip()
        kind = str(item.get("kind", "days_between") or "days_between").strip().lower()
        start_field = str(item.get("start_field", "") or "").strip()
        end_field = str(item.get("end_field", "") or "").strip()
        priority = str(item.get("priority", "rule_wins") or "rule_wins").strip().lower()
        if priority not in {"rule_wins", "always"}:
            priority = "rule_wins"
        if kind != "days_between":
            continue
        if not field or field in seen_computed_destinations:
            continue
        # Manteniamo configurazioni legacy anche se un campo è stato eliminato
        # manualmente dal JSON; la GUI le evidenzierà e permetterà di correggerle.
        computed_fields.append({
            "field": field,
            "kind": kind,
            "start_field": start_field,
            "end_field": end_field,
            "priority": priority,
        })
        seen_computed_destinations.add(field)
    profile["computed_fields"] = computed_fields

    profile["schema_version"] = PROFILE_VERSION
    profile.setdefault("reconcile_keys", [])
    profile.setdefault("state", {
        "folder": "INBOX", "uidvalidity": 0,
        "last_processed_uid": 0, "last_processed_at": ""
    })
    # Fino a sei valori possono essere mostrati nella home. I riferimenti
    # sono tipizzati per distinguere campi, campi calcolati e formule nominate.
    final_summary_names = {
        str(f.get("name", "")) for f in normalized_fields
        if isinstance(f, dict) and str(f.get("final_formula", "") or "").lower() in {"sum", "count"}
    }
    raw_computed_names = {
        str(x.get("field", "") or "") for x in profile.get("computed_fields", []) or []
        if isinstance(x, dict) and x.get("field")
    }
    summary_names = {
        str(x.get("name", "") or "") for x in profile.get("summary_formulas", []) or []
        if isinstance(x, dict) and x.get("name")
    }
    valid_home_refs = (
        {name for name in final_summary_names}
        | {f"computed:{name}" for name in raw_computed_names if name}
        | {f"summary:{name}" for name in summary_names if name}
    )
    home_summary_fields: list[str] = []
    for ref in profile.get("home_summary_fields", []) or []:
        ref = str(ref or "").strip()
        if not ref:
            continue
        if ref.startswith("field:") and ref.split(":", 1)[1] in final_summary_names:
            ref = ref.split(":", 1)[1]
        if ref in valid_home_refs and ref not in home_summary_fields:
            home_summary_fields.append(ref)
        if len(home_summary_fields) >= 6:
            break
    profile["home_summary_fields"] = home_summary_fields

    return profile


def get_field(profile: dict[str, Any], name: str) -> dict[str, Any] | None:
    for field in profile.get("fields", []):
        if isinstance(field, dict) and field.get("name") == name:
            return field
    return None


def field_generation_sources(profile: dict[str, Any], field_name: str) -> list[dict[str, Any]]:
    """Restituisce tutte le sorgenti che possono valorizzare ``field_name``.

    La funzione è volutamente indipendente dalla GUI: viene usata dalla
    schermata di tracciabilità dei campi e può essere verificata dai test.
    Sono incluse:
    - regole di estrazione dei singoli Tipi mail;
    - campi calcolati globali del profilo;
    - contatore automatico di match del Tipo mail associato al campo.

    Le formule finali Excel non sono incluse perché producono una cella di
    riepilogo in fondo alla colonna, non il valore del campo nelle righe dati.
    """
    target = str(field_name or "").strip()
    if not target:
        return []

    out: list[dict[str, Any]] = []

    for mt in profile.get("mail_types", []) or []:
        if not isinstance(mt, dict):
            continue
        mt_name = str(mt.get("name", "") or "")
        for index, rule in enumerate(mt.get("rules", []) or []):
            if not isinstance(rule, dict) or str(rule.get("field", "") or "") != target:
                continue
            strategy = str(rule.get("strategy", "") or "")
            source = str(rule.get("source", "") or "")
            detail_parts: list[str] = []
            if strategy == "regex" and rule.get("pattern"):
                detail_parts.append(f"regex: {rule.get('pattern')}")
            elif strategy == "after_trigger" and rule.get("trigger"):
                detail_parts.append(f"trigger: {rule.get('trigger')}")
            elif strategy == "constant":
                detail_parts.append(f"valore costante: {rule.get('value', '')}")
            elif rule.get("trigger"):
                detail_parts.append(f"trigger: {rule.get('trigger')}")
            policy = str(rule.get("existing_value_policy", "replace") or "replace")
            if policy:
                detail_parts.append(f"se già presente: {policy}")
            out.append({
                "kind": "extraction_rule",
                "kind_label": "Regola di estrazione",
                "scope": mt_name,
                "mail_type": mt_name,
                "name": str(rule.get("name", "") or f"regola_{index + 1}"),
                "source": source,
                "mechanism": strategy,
                "detail": " | ".join(detail_parts),
                "raw": rule,
            })

    for index, comp in enumerate(profile.get("computed_fields", []) or []):
        if not isinstance(comp, dict) or str(comp.get("field", "") or "") != target:
            continue
        kind = str(comp.get("kind", "") or "")
        if kind == "days_between":
            mechanism = "Differenza giorni tra due date"
            detail = (
                f"{comp.get('start_field', '')} → {comp.get('end_field', '')}; "
                f"priorità={comp.get('priority', 'rule_wins')}"
            )
        else:
            mechanism = kind
            detail = ""
        out.append({
            "kind": "computed_field",
            "kind_label": "Regola generale",
            "scope": "Profilo",
            "mail_type": "",
            "name": str(comp.get("name", "") or f"campo_calcolato_{index + 1}"),
            "source": "campi",
            "mechanism": mechanism,
            "detail": detail,
            "raw": comp,
        })

    field = get_field(profile, target)
    if field:
        counter = str(field.get("mail_type_counter", "") or "").strip()
        if counter:
            out.append({
                "kind": "mail_type_counter",
                "kind_label": "Regola generale",
                "scope": "Profilo",
                "mail_type": counter,
                "name": "Contatore Tipo mail",
                "source": "match Tipo mail",
                "mechanism": "Incremento contatore",
                "detail": (
                    f"Incrementa di 1 il campo '{target}' ogni volta che il Tipo mail "
                    f"'{counter}' ha riscontro sulla stessa riga riconciliata."
                ),
                "raw": {"mail_type_counter": counter, "field": target},
            })

    return out


def ensure_field(
    profile: dict[str, Any],
    name: str,
    field_type: str = "text",
    *,
    display_format: str | None = None,
) -> dict[str, Any]:
    field = get_field(profile, name)
    if field is None:
        field = {"name": name, "type": normalize_field_type(field_type)}
        profile.setdefault("fields", []).append(field)
    else:
        field["type"] = normalize_field_type(field_type or field.get("type"))
    if display_format:
        field["format"] = display_format
    elif field.get("format") == "currency_eur" and field["type"] != "number":
        field.pop("format", None)
    return field


def slugify(value: str, fallback: str = "regola") -> str:
    value = value.lower().strip()
    value = re.sub(r"[^a-z0-9àèéìòù]+", "_", value)
    value = value.strip("_")
    return value[:48] or fallback



def rename_mail_type(profile: dict[str, Any], old_name: str, new_name: str) -> str:
    """Rinomina un tipo mail e restituisce il nome normalizzato definitivo."""
    old_name = str(old_name or "")
    if not old_name:
        raise ValueError("Tipo mail non valido")
    new_slug = slugify(new_name, "tipo_mail")
    if not new_slug:
        raise ValueError("Inserisci il nuovo nome")
    for mt in profile.get("mail_types", []):
        if isinstance(mt, dict) and mt.get("name") == new_slug and mt.get("name") != old_name:
            raise ValueError(f"Esiste già un tipo mail chiamato '{new_slug}'")
    target = next((mt for mt in profile.get("mail_types", []) if isinstance(mt, dict) and mt.get("name") == old_name), None)
    if target is None:
        raise ValueError(f"Tipo mail '{old_name}' non trovato")
    target["name"] = new_slug
    for field in profile.get("fields", []):
        if isinstance(field, dict) and field.get("mail_type_counter") == old_name:
            field["mail_type_counter"] = new_slug
    return new_slug


def delete_mail_type(profile: dict[str, Any], name: str) -> bool:
    """Elimina un tipo mail e tutte le sue regole di selezione/estrazione."""
    items = profile.get("mail_types", [])
    for i, mt in enumerate(items):
        if isinstance(mt, dict) and mt.get("name") == name:
            del items[i]
            for field in profile.get("fields", []):
                if isinstance(field, dict) and field.get("mail_type_counter") == name:
                    field.pop("mail_type_counter", None)
            return True
    return False


def _unique_rule_name(existing_names: set[str], requested: str) -> str:
    """Restituisce un nome regola non ancora usato nel tipo mail destinazione."""
    base = str(requested or "regola_copiata").strip() or "regola_copiata"
    if base not in existing_names:
        return base
    candidate = f"{base}_copia"
    if candidate not in existing_names:
        return candidate
    n = 2
    while f"{base}_copia_{n}" in existing_names:
        n += 1
    return f"{base}_copia_{n}"


def copy_extraction_rules(
    source_profile: dict[str, Any],
    source_mail_type_name: str,
    rule_indexes: list[int] | tuple[int, ...],
    destination_profile: dict[str, Any],
    destination_mail_type_name: str,
) -> dict[str, Any]:
    """Copia una o più regole di estrazione tra profili/tipi mail.

    Le regole vengono copiate profondamente. Se il campo destinazione della
    regola non esiste nel profilo di arrivo, viene copiata anche la relativa
    definizione di campo. Lo stato di chiave di riconciliazione viene
    preservato. In caso di nome regola già presente, viene generato un nome
    univoco con suffisso ``_copia``.
    """
    normalize_profile(source_profile)
    normalize_profile(destination_profile)

    source_mt = next((
        mt for mt in source_profile.get("mail_types", [])
        if isinstance(mt, dict) and mt.get("name") == source_mail_type_name
    ), None)
    if source_mt is None:
        raise ValueError(f"Tipo mail sorgente '{source_mail_type_name}' non trovato")

    destination_mt = next((
        mt for mt in destination_profile.get("mail_types", [])
        if isinstance(mt, dict) and mt.get("name") == destination_mail_type_name
    ), None)
    if destination_mt is None:
        raise ValueError(f"Tipo mail destinazione '{destination_mail_type_name}' non trovato")

    source_rules = source_mt.get("rules", [])
    selected: list[dict[str, Any]] = []
    for raw_index in rule_indexes:
        index = int(raw_index)
        if index < 0 or index >= len(source_rules):
            raise ValueError(f"Indice regola non valido: {index}")
        rule = source_rules[index]
        if isinstance(rule, dict):
            selected.append(rule)

    if not selected:
        raise ValueError("Seleziona almeno una regola di estrazione")

    destination_rules = destination_mt.setdefault("rules", [])
    existing_names = {
        str(rule.get("name", "")) for rule in destination_rules if isinstance(rule, dict)
    }
    source_keys = set(source_profile.get("reconcile_keys", []))
    destination_keys = destination_profile.setdefault("reconcile_keys", [])

    copied_names: list[str] = []
    added_fields: list[str] = []
    added_reconcile_keys: list[str] = []

    for source_rule in selected:
        rule = copy.deepcopy(source_rule)
        old_name = str(rule.get("name", "regola_copiata"))
        new_name = _unique_rule_name(existing_names, old_name)
        rule["name"] = new_name
        existing_names.add(new_name)

        field_name = str(rule.get("field", "")).strip()
        if field_name:
            source_field = get_field(source_profile, field_name)
            if get_field(destination_profile, field_name) is None and source_field is not None:
                destination_profile.setdefault("fields", []).append(copy.deepcopy(source_field))
                added_fields.append(field_name)
            if field_name in source_keys and field_name not in destination_keys:
                destination_keys.append(field_name)
                added_reconcile_keys.append(field_name)

        destination_rules.append(rule)
        copied_names.append(new_name)

    normalize_profile(destination_profile)
    return {
        "copied_count": len(copied_names),
        "copied_names": copied_names,
        "added_fields": added_fields,
        "added_reconcile_keys": added_reconcile_keys,
    }

def new_profile(name: str) -> dict[str, Any]:
    return {
        "schema_version": PROFILE_VERSION,
        "name": name,
        "description": "",
        "fields": [],
        "reconcile_keys": [],
        "mail_types": [],
        "computed_fields": [],
        "summary_formulas": [],
        "home_summary_fields": [],
        "state": {
            "folder": "INBOX",
            "uidvalidity": 0,
            "last_processed_uid": 0,
            "last_processed_at": ""
        }
    }


class ProfileManager:
    def __init__(self, profiles_dir: Path):
        self.profiles_dir = Path(profiles_dir)
        self.profiles_dir.mkdir(parents=True, exist_ok=True)

    def list_profiles(self) -> list[Path]:
        return sorted(self.profiles_dir.glob("*.json"), key=lambda p: p.name.casefold())

    def load(self, path: str | Path) -> dict[str, Any]:
        return normalize_profile(json.loads(Path(path).read_text(encoding="utf-8")))

    def save(self, path: str | Path, profile: dict[str, Any]) -> None:
        import copy
        normalize_profile(profile)
        clean = copy.deepcopy(profile)
        # Stato operativo e checkpoint appartengono alla coppia email + profilo,
        # non alla definizione trasferibile del profilo di estrazione.
        clean.pop("state", None)
        clean.pop("state_by_email", None)
        clean.pop("extraction_start_date", None)
        Path(path).write_text(json.dumps(clean, indent=2, ensure_ascii=False), encoding="utf-8")

    def create(self, name: str) -> Path:
        base = slugify(name, "profilo")
        path = self.profiles_dir / f"{base}.json"
        n = 2
        while path.exists():
            path = self.profiles_dir / f"{base}_{n}.json"
            n += 1
        self.save(path, new_profile(name))
        return path

    @staticmethod
    def touch_state(profile: dict[str, Any], folder: str, uidvalidity: int, last_uid: int) -> None:
        profile.setdefault("state", {})
        profile["state"].update({
            "folder": folder,
            "uidvalidity": int(uidvalidity or 0),
            "last_processed_uid": int(last_uid or 0),
            "last_processed_at": datetime.now().isoformat(timespec="seconds"),
        })
