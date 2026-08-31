from __future__ import annotations

from typing import Any, Iterable


_REGEX_ABBREVIATIONS: list[tuple[str, str]] = []


def normalize_regex_abbreviations(value: Any) -> list[dict[str, str]]:
    """Normalizza la tabella globale delle abbreviazioni regex.

    Accetta sia il nuovo formato lista di oggetti sia, per robustezza, un dict.
    Le chiavi vuote vengono ignorate; in caso di duplicati vince l'ultima.
    """
    pairs: list[tuple[str, str]] = []
    if isinstance(value, dict):
        pairs = [(str(k), str(v)) for k, v in value.items()]
    elif isinstance(value, list):
        for item in value:
            if isinstance(item, dict):
                token = str(item.get("token", item.get("abbreviation", "")) or "")
                replacement = str(item.get("replacement", item.get("regex", "")) or "")
                pairs.append((token, replacement))
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                pairs.append((str(item[0]), str(item[1])))

    dedup: dict[str, str] = {}
    order: list[str] = []
    for token, replacement in pairs:
        token = token.strip()
        if not token:
            continue
        if token not in dedup:
            order.append(token)
        dedup[token] = replacement
    return [{"token": token, "replacement": dedup[token]} for token in order]


def set_regex_abbreviations(value: Any) -> None:
    """Imposta le abbreviazioni usate dal preprocessore regex dell'applicazione."""
    global _REGEX_ABBREVIATIONS
    normalized = normalize_regex_abbreviations(value)
    # Token più lunghi prima: evita che un'abbreviazione corta contenuta in una
    # più lunga venga espansa prematuramente.
    _REGEX_ABBREVIATIONS = sorted(
        [(item["token"], item["replacement"]) for item in normalized],
        key=lambda pair: len(pair[0]),
        reverse=True,
    )


def get_regex_abbreviations() -> list[dict[str, str]]:
    return [{"token": token, "replacement": replacement} for token, replacement in _REGEX_ABBREVIATIONS]


def expand_regex(pattern: str | None) -> str:
    """Espande letteralmente le abbreviazioni presenti in una regex.

    Il token è testo letterale, non una regex. La sostituzione viene inserita
    così com'è. Si eseguono più passaggi (max 20) per consentire abbreviazioni
    composte; un ciclo si arresta automaticamente quando il testo non cambia.
    """
    text = str(pattern or "")
    if not text or not _REGEX_ABBREVIATIONS:
        return text
    for _ in range(20):
        previous = text
        for token, replacement in _REGEX_ABBREVIATIONS:
            text = text.replace(token, replacement)
        if text == previous:
            break
    return text
