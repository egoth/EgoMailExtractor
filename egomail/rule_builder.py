from __future__ import annotations

import re
from typing import Any

from .profile import slugify


def propose_field_name(trigger: str, selected: str, kind: str) -> str:
    source = trigger.strip() or selected.strip()
    proposed = slugify(source, f"campo_{kind}")
    generic = {"totale": "importo_totale", "check_in": "data_check_in", "check_out": "data_check_out", "ospiti": "numero_ospiti"}
    return generic.get(proposed, proposed)


def infer_rule_from_selection(body: str, trigger: str, selected: str, field: str, kind: str) -> dict[str, Any]:
    transform = {
        "text": "text",
        "number": "int",
        "money": "money",
        "date": "date_it_email_year",
        "key": "upper",
    }.get(kind, "text")

    # Strategia principale: il valore è una delle righe non vuote successive al trigger.
    if trigger:
        lines = [x.strip() for x in body.splitlines()]
        trigger_indexes = [i for i, x in enumerate(lines) if trigger.casefold() in x.casefold()]
        selected_index = next((i for i, x in enumerate(lines) if selected.strip() and selected.strip() in x), None)
        if trigger_indexes and selected_index is not None:
            preceding = [i for i in trigger_indexes if i < selected_index]
            if preceding:
                t = preceding[-1]
                nonempty = [x for x in lines[t + 1:selected_index + 1] if x]
                offset = max(1, len(nonempty))
                return {
                    "name": f"estrai_{field}",
                    "field": field,
                    "source": "body",
                    "strategy": "after_trigger",
                    "trigger": trigger,
                    "nonempty_offset": offset,
                    "transform": transform,
                }

    # Fallback: regex basata sul campione, utile per codici/valori ben caratterizzati.
    if kind == "money":
        pattern = r"([+-]?\d{1,3}(?:\.\d{3})*,\d{2})\s*€"
    elif kind == "number":
        pattern = r"(\d+)"
    elif kind == "date":
        pattern = r"((?:lun|mar|mer|gio|ven|sab|dom)?\s*\d{1,2}\s+[A-Za-zàèéìòù]+(?:\s+\d{4})?)"
    elif kind == "key" and re.fullmatch(r"[A-Za-z0-9-]{5,}", selected.strip()):
        pattern = rf"\b({re.escape(selected.strip())})\b"
    else:
        pattern = rf"({re.escape(selected.strip())})"
    return {
        "name": f"estrai_{field}",
        "field": field,
        "source": "body",
        "strategy": "regex",
        "pattern": pattern,
        "group": 1,
        "transform": transform,
    }
