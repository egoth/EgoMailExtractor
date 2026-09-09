from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from typing import Any

from .mail import MailMessage, normalize_text
from .regex_utils import expand_regex

MONTHS_IT = {
    "gen": 1, "gennaio": 1,
    "feb": 2, "febbraio": 2,
    "mar": 3, "marzo": 3,
    "apr": 4, "aprile": 4,
    "mag": 5, "maggio": 5,
    "giu": 6, "giugno": 6,
    "lug": 7, "luglio": 7,
    "ago": 8, "agosto": 8,
    "set": 9, "sett": 9, "settembre": 9,
    "ott": 10, "ottobre": 10,
    "nov": 11, "novembre": 11,
    "dic": 12, "dicembre": 12,
}


def parse_money(value: str) -> float | None:
    if value is None:
        return None
    s = re.sub(r"[^0-9,.-]", "", str(value))
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_int(value: str) -> int | None:
    m = re.search(r"-?\d+", str(value))
    return int(m.group()) if m else None


def parse_guest_count(value: str) -> int | None:
    nums = [int(x) for x in re.findall(r"\d+", str(value))]
    return sum(nums) if nums else None


def parse_it_date(value: str, email_year: int | None) -> str | None:
    if not value:
        return None
    # Se la sorgente è la Data mail, il valore è spesso un header RFC 2822
    # (es. Sun, 25 Oct 2026 11:12:13 +0200). Proviamo prima il parser email.
    try:
        parsed = parsedate_to_datetime(str(value))
        if parsed is not None:
            return parsed.date().isoformat()
    except Exception:
        pass
    s = value.lower().replace(".", " ")
    # Rimuoviamo l'eventuale giorno della settimana solo se compare come
    # prefisso della data.  Non va eliminato globalmente: ``mar`` può infatti
    # significare sia martedì sia marzo.  Esempi corretti:
    #   "mer 25 mar" -> "25 mar"
    #   "mar 25 mar" -> "25 mar"
    #   "25 mar"     -> "25 mar"
    s = re.sub(
        r"^\s*(?:lun|mar|mer|gio|ven|sab|dom)\b[\s,]*",
        "",
        s,
        count=1,
        flags=re.I,
    )
    s = re.sub(r"\s+", " ", s).strip()
    m = re.search(r"(\d{1,2})\s+([a-zàèéìòù]+)(?:\s+(\d{4}))?", s)
    if not m:
        return None
    day = int(m.group(1))
    month = MONTHS_IT.get(m.group(2))
    if not month:
        return None
    year = int(m.group(3)) if m.group(3) else email_year
    if not year:
        year = datetime.now().year
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return None


def _convert_numeric_date(value: str, *, source: str, target: str) -> str | None:
    """Converte date numeriche USA/EU conservando un output ISO per Excel.

    source='us': MM/DD/YYYY; source='eu': DD/MM/YYYY.
    target è usato solo per chiarezza semantica: il valore restituito resta
    ISO YYYY-MM-DD, formato interno coerente con le altre conversioni Data.
    """
    text = str(value or "").strip()
    m = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})", text)
    if not m:
        return None
    a, b, y = map(int, m.groups())
    month, day = (a, b) if source == "us" else (b, a)
    try:
        return date(y, month, day).isoformat()
    except ValueError:
        return None


def _parse_numeric_datetime(value: str, *, source: str) -> datetime | None:
    """Legge date/date-ora numeriche EU o USA, con secondi e millisecondi opzionali.

    Sono accettati anche valori come ``13/07/2026, 16:42``.
    ``source='eu'`` interpreta GG/MM/AAAA; ``source='us'`` MM/GG/AAAA.
    """
    text = str(value or "").strip()
    m = re.search(
        r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})"
        r"(?:\s*,?\s+(\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d{1,6}))?)?)?",
        text,
    )
    if not m:
        return None
    a, b, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    month, day = (a, b) if source == "us" else (b, a)
    hour = int(m.group(4) or 0)
    minute = int(m.group(5) or 0)
    second = int(m.group(6) or 0)
    fraction = (m.group(7) or "")[:6]
    microsecond = int(fraction.ljust(6, "0")) if fraction else 0
    try:
        return datetime(year, month, day, hour, minute, second, microsecond)
    except ValueError:
        return None


def _convert_target_date(value: str, *, source: str, include_time: bool) -> str | None:
    parsed = _parse_numeric_datetime(value, source=source)
    if parsed is None:
        # La strategia email_date fornisce normalmente un header RFC 2822,
        # es. "Mon, 13 Jul 2026 16:42:00 +0000". Le conversioni Data/Data-Ora
        # devono quindi accettare anche questo formato senza richiedere Regex.
        try:
            parsed = parsedate_to_datetime(str(value or "").strip())
        except Exception:
            parsed = None
    if parsed is None:
        return None
    if not include_time:
        return parsed.date().isoformat()
    # Formato interno non ambiguo; Excel lo scriverà come vera cella Data-Ora.
    if parsed.microsecond:
        return parsed.isoformat(sep=" ", timespec="milliseconds")
    return parsed.isoformat(sep=" ", timespec="seconds")


def transform(value: Any, kind: str | None, mail: MailMessage) -> Any:
    if value is None:
        return None
    if not kind or kind == "text":
        return str(value).strip()
    if kind == "money":
        return parse_money(str(value))
    if kind == "int":
        return parse_int(str(value))
    if kind == "guest_count":
        return parse_guest_count(str(value))
    if kind == "date_it_email_year":
        return parse_it_date(str(value), mail.year)
    if kind == "date_us_to_eu":
        return _convert_numeric_date(str(value), source="us", target="eu")
    if kind == "date_eu_to_us":
        return _convert_numeric_date(str(value), source="eu", target="us")
    if kind == "date_eu":
        return _convert_target_date(str(value), source="eu", include_time=False)
    if kind == "datetime_eu":
        return _convert_target_date(str(value), source="eu", include_time=True)
    if kind == "date_us":
        return _convert_target_date(str(value), source="us", include_time=False)
    if kind == "datetime_us":
        return _convert_target_date(str(value), source="us", include_time=True)
    if kind == "upper":
        return str(value).strip().upper()
    return str(value).strip()


def _source(mail: MailMessage, name: str) -> str:
    if name == "subject":
        return mail.subject or ""
    if name == "sender":
        return mail.sender or ""
    if name == "recipient":
        return getattr(mail, "recipient", "") or ""
    if name == "date":
        return mail.date or mail.raw_date or ""
    if name == "message_id":
        return mail.message_id or ""
    return mail.body or ""


def match_selection_rule(mail: MailMessage, rule: dict[str, Any]) -> bool:
    """Valuta una singola regola di selezione del tipo mail.

    Oltre al trigger, rispetta l'eventuale intervallo valid_from/valid_to
    applicato alla data della mail, con estremi inclusi.
    """
    source_name = str(rule.get("source", "body") or "body")
    valid, _reason = rule_valid_for_mail_date(mail, rule)
    if not valid:
        return False
    # "Data" è un criterio autonomo: non usa Trigger/Regex ma soltanto i due
    # limiti valid_from/valid_to. Basta un solo estremo compilato.
    if source_name == "date":
        return bool(
            str(rule.get("valid_from", "") or "").strip()
            or str(rule.get("valid_to", "") or "").strip()
        )
    trigger = str(rule.get("trigger", "") or "").strip()
    if not trigger:
        return True
    source = _source(mail, source_name)
    mode = str(rule.get("mode", "contains") or "contains").lower()
    if mode == "regex":
        try:
            return bool(re.search(expand_regex(trigger), source, flags=re.I | re.S))
        except re.error:
            return False
    return trigger.casefold() in source.casefold()


def matches_mail_type(mail: MailMessage, mail_type: dict[str, Any]) -> bool:
    if not bool(mail_type.get("enabled", True)):
        return False
    # Nuovo formato v1.5: lista arbitraria di regole di match. Tutte le
    # regole compilate devono valere (AND). Un tipo senza trigger non seleziona
    # nessuna mail, così non può catturare accidentalmente l'intera casella.
    match_rules = mail_type.get("match_rules")
    if match_rules is not None:
        active = [
            r for r in match_rules
            if isinstance(r, dict) and (
                str(r.get("trigger", "") or "").strip()
                or (
                    str(r.get("source", "") or "") == "date"
                    and (
                        str(r.get("valid_from", "") or "").strip()
                        or str(r.get("valid_to", "") or "").strip()
                    )
                )
            )
        ]
        return bool(active) and all(match_selection_rule(mail, r) for r in active)

    # Compatibilità difensiva per profili non ancora normalizzati.
    criteria = mail_type.get("criteria", {})
    tests = []
    for key, source in (("subject_regex", mail.subject), ("sender_regex", mail.sender), ("body_regex", mail.body)):
        pattern = criteria.get(key)
        if pattern:
            try:
                tests.append(bool(re.search(expand_regex(str(pattern)), source or "", flags=re.I | re.S)))
            except re.error:
                tests.append(False)
    for key, source in (("subject_contains", mail.subject), ("sender_contains", mail.sender), ("body_contains", mail.body)):
        value = criteria.get(key)
        if value:
            tests.append(str(value).casefold() in (source or "").casefold())
    return all(tests) if tests else False


def header_prefilter_matches_profile(mail: MailMessage, profile: dict[str, Any]) -> bool:
    """Pre-filtra i Tipi mail usando esclusivamente dati già presenti negli header.

    Le regole su mittente/oggetto/data/destinatario/message-id vengono valutate
    subito. Le regole che richiedono il body vengono rinviate alla fase di
    Estrazione. Una mail è candidata se esiste almeno un Tipo mail per il quale
    tutte le regole valutabili dagli header risultano vere. Un Tipo composto
    esclusivamente da regole sul body non può essere escluso in Lettura e quindi
    mantiene la mail come candidata.
    """
    header_sources = {"sender", "subject", "date", "recipient", "message_id"}
    for mail_type in profile.get("mail_types", []) or []:
        if not isinstance(mail_type, dict):
            continue
        if not bool(mail_type.get("enabled", True)):
            continue
        rules = mail_type.get("match_rules")
        if rules is None:
            # Compatibilità con i profili storici: i criteri subject/sender sono
            # valutabili dagli header, quelli body vengono rimandati.
            criteria = mail_type.get("criteria", {}) or {}
            tests: list[bool] = []
            for key, source in (
                ("subject_regex", mail.subject),
                ("sender_regex", mail.sender),
                ("subject_contains", mail.subject),
                ("sender_contains", mail.sender),
            ):
                value = criteria.get(key)
                if not value:
                    continue
                if key.endswith("_regex"):
                    try:
                        tests.append(bool(re.search(expand_regex(str(value)), source or "", flags=re.I | re.S)))
                    except re.error:
                        tests.append(False)
                else:
                    tests.append(str(value).casefold() in (source or "").casefold())
            has_body_rule = bool(criteria.get("body_regex") or criteria.get("body_contains"))
            if tests and all(tests):
                return True
            if not tests and has_body_rule:
                return True
            continue

        active = [
            r for r in rules
            if isinstance(r, dict) and (
                str(r.get("trigger", "") or "").strip()
                or (
                    str(r.get("source", "") or "") == "date"
                    and (
                        str(r.get("valid_from", "") or "").strip()
                        or str(r.get("valid_to", "") or "").strip()
                    )
                )
            )
        ]
        if not active:
            continue
        header_rules = [r for r in active if str(r.get("source", "body") or "body") in header_sources]
        if all(match_selection_rule(mail, r) for r in header_rules):
            return True
    return False

def header_prefilter_matching_mail_types(mail: MailMessage, profile: dict[str, Any]) -> list[str]:
    """Nomi dei Tipi mail che non possono essere esclusi usando i soli header.

    Una regola sul body viene rinviata e quindi non rende falso il Tipo mail in
    Lettura. Serve anche per rendere esplicito nel log perché una mail è candidata.
    """
    header_sources = {"sender", "subject", "date", "recipient", "message_id"}
    names: list[str] = []
    for mail_type in profile.get("mail_types", []) or []:
        if not isinstance(mail_type, dict):
            continue
        if not bool(mail_type.get("enabled", True)):
            continue
        rules = mail_type.get("match_rules")
        if rules is None:
            # Per i profili storici deleghiamo al vecchio pre-filtro globale:
            # non è possibile attribuire con certezza il singolo Tipo senza
            # duplicarne tutta la logica legacy.
            if header_prefilter_matches_profile(mail, {"mail_types": [mail_type]}):
                names.append(str(mail_type.get("name", "") or ""))
            continue
        active = [
            r for r in rules
            if isinstance(r, dict) and (
                str(r.get("trigger", "") or "").strip()
                or (
                    str(r.get("source", "") or "") == "date"
                    and (str(r.get("valid_from", "") or "").strip() or str(r.get("valid_to", "") or "").strip())
                )
            )
        ]
        if not active:
            continue
        header_rules = [r for r in active if str(r.get("source", "body") or "body") in header_sources]
        if all(match_selection_rule(mail, r) for r in header_rules):
            names.append(str(mail_type.get("name", "") or ""))
    return names


def matching_mail_types(mail: MailMessage, profile: dict[str, Any]) -> list[dict[str, Any]]:
    """Restituisce al massimo il primo Tipo mail che fa match, secondo l'ordine del profilo."""
    for mt in profile.get("mail_types", []) or []:
        if isinstance(mt, dict) and matches_mail_type(mail, mt):
            return [mt]
    return []


@dataclass
class ExtractionTrace:
    """Risultato dettagliato dell'applicazione di una regola di estrazione.

    Gli span sono espressi rispetto alla sorgente normalizzata (corpo,
    oggetto o mittente) e permettono alla GUI di evidenziare esattamente la
    zona che ha fatto scattare la regola e il testo realmente acquisito.
    """

    value: Any = None
    raw_value: Any = None
    source_name: str = "body"
    source_text: str = ""
    matched: bool = False
    trigger_span: tuple[int, int] | None = None
    value_span: tuple[int, int] | None = None
    error: str = ""


def _occurrence_index(value: Any) -> int:
    try:
        occurrence = int(value or 1)
    except (TypeError, ValueError):
        occurrence = 1
    return occurrence - 1 if occurrence > 0 else occurrence


def _trimmed_line_spans(source: str) -> list[tuple[str, int, int]]:
    """Restituisce (testo_strippato, start, end) per ogni riga.

    start/end puntano al testo privo degli spazi ai bordi, così gli span
    possono essere applicati direttamente al widget Text.
    """
    out: list[tuple[str, int, int]] = []
    pos = 0
    for chunk in source.splitlines(keepends=True):
        content = chunk.rstrip("\r\n")
        left = len(content) - len(content.lstrip())
        right = len(content.rstrip())
        out.append((content.strip(), pos + left, pos + right))
        pos += len(chunk)
    # splitlines() restituisce lista vuota per stringa vuota e non aggiunge
    # l'ultima riga dopo un newline terminale: per l'estrazione non serve.
    if source and not out:
        out.append((source.strip(), 0, len(source.rstrip())))
    return out



def _mail_date_for_rule_validity(mail: MailMessage) -> date | None:
    """Restituisce la data civile della mail per il controllo validità regola."""
    raw = getattr(mail, "raw_date", "") or getattr(mail, "date", "") or ""
    if raw:
        try:
            parsed = parsedate_to_datetime(str(raw))
            if parsed is not None:
                return parsed.date()
        except Exception:
            pass
        text = str(raw).strip()
        for candidate in (text[:10], text):
            try:
                return date.fromisoformat(candidate)
            except Exception:
                pass
    return None


def rule_valid_for_mail_date(mail: MailMessage, rule: dict[str, Any]) -> tuple[bool, str]:
    """Controlla gli estremi inclusivi valid_from/valid_to della regola."""
    start_text = str(rule.get("valid_from", "") or "").strip()
    end_text = str(rule.get("valid_to", "") or "").strip()
    if not start_text and not end_text:
        return True, ""
    mail_day = _mail_date_for_rule_validity(mail)
    if mail_day is None:
        return False, "Data mail non interpretabile per il controllo di validità"
    try:
        start = date.fromisoformat(start_text) if start_text else None
        end = date.fromisoformat(end_text) if end_text else None
    except ValueError:
        return False, "Intervallo di validità della regola non valido"
    if start is not None and mail_day < start:
        return False, f"Regola non valida alla data mail {mail_day.isoformat()}: inizio validità {start.isoformat()}"
    if end is not None and mail_day > end:
        return False, f"Regola non valida alla data mail {mail_day.isoformat()}: fine validità {end.isoformat()}"
    return True, ""


def extract_rule_trace(mail: MailMessage, rule: dict[str, Any]) -> ExtractionTrace:
    """Applica una regola restituendo anche le posizioni usate dalla GUI."""
    source_name = str(rule.get("source", "body") or "body")
    source = normalize_text(_source(mail, source_name))
    trace = ExtractionTrace(source_name=source_name, source_text=source)
    strategy = str(rule.get("strategy", "regex") or "regex")

    valid, validity_message = rule_valid_for_mail_date(mail, rule)
    if not valid:
        trace.error = validity_message
        return trace

    try:
        if strategy == "regex":
            pattern = str(rule.get("pattern", "") or "")
            if not pattern:
                return trace
            expanded_pattern = expand_regex(pattern)
            matches = list(re.finditer(expanded_pattern, source, flags=re.I | re.S))
            if not matches:
                return trace
            idx = _occurrence_index(rule.get("occurrence", 1))
            try:
                m = matches[idx]
            except IndexError:
                return trace
            group = int(rule.get("group", 1) or 1)
            trace.trigger_span = m.span(0)
            try:
                trace.raw_value = m.group(group)
                trace.value_span = m.span(group)
            except (IndexError, ValueError):
                trace.raw_value = m.group(0)
                trace.value_span = m.span(0)
            trace.matched = True

        elif strategy == "after_trigger":
            trigger = str(rule.get("trigger", "") or "")
            if not trigger:
                return trace
            lines = _trimmed_line_spans(source)
            hits: list[tuple[int, tuple[int, int]]] = []
            for i, (line, start, _end) in enumerate(lines):
                if trigger.casefold() not in line.casefold():
                    continue
                m = re.search(re.escape(trigger), line, flags=re.I)
                if m:
                    span = (start + m.start(), start + m.end())
                else:
                    span = (start, start + len(line))
                hits.append((i, span))
            if not hits:
                return trace
            hit_idx = _occurrence_index(rule.get("occurrence", 1))
            try:
                start_line, trigger_span = hits[hit_idx]
            except IndexError:
                return trace
            trace.trigger_span = trigger_span

            try:
                offset = int(rule.get("nonempty_offset", 1) or 1)
            except (TypeError, ValueError):
                offset = 1
            candidates = [(line, st, en) for line, st, en in lines[start_line + 1:] if line]
            if offset <= 0 or len(candidates) < offset:
                return trace
            raw, raw_start, raw_end = candidates[offset - 1]
            trace.raw_value = raw
            trace.value_span = (raw_start, raw_end)

            pattern = str(rule.get("value_pattern", "") or "")
            if pattern:
                m = re.search(expand_regex(pattern), raw, flags=re.I | re.S)
                if not m:
                    trace.raw_value = None
                    trace.value_span = None
                    return trace
                group = int(rule.get("group", 1) or 1)
                try:
                    trace.raw_value = m.group(group)
                    gst, gen = m.span(group)
                except (IndexError, ValueError):
                    trace.raw_value = m.group(0)
                    gst, gen = m.span(0)
                trace.value_span = (raw_start + gst, raw_start + gen)
            trace.matched = True

        elif strategy == "email_date":
            trace.raw_value = mail.date or mail.raw_date or ""
            trace.matched = bool(trace.raw_value)
        elif strategy == "email_sender":
            trace.raw_value = mail.sender or ""
            trace.matched = bool(trace.raw_value)
        elif strategy == "email_recipient":
            trace.raw_value = getattr(mail, "recipient", "") or ""
            trace.matched = bool(trace.raw_value)
        elif strategy == "mail_type_counter":
            # Il valore viene incrementato una sola volta dall'exporter tramite
            # il metadato field.mail_type_counter. La regola resta visibile
            # nell'elenco ma non deve produrre un secondo +1.
            trace.raw_value = None
            trace.matched = True
        elif strategy == "constant":
            trace.raw_value = rule.get("value")
            trace.matched = True
        else:
            trace.error = f"Strategia sconosciuta: {strategy}"
            return trace

        trace.value = transform(trace.raw_value, rule.get("transform"), mail)
        return trace
    except re.error as exc:
        trace.error = f"Regex non valida: {exc}"
        return trace
    except Exception as exc:
        trace.error = str(exc)
        return trace


def extract_rule(mail: MailMessage, rule: dict[str, Any]) -> Any:
    return extract_rule_trace(mail, rule).value


def humanize_extraction_rule(
    rule: dict[str, Any],
    *,
    field_type: str = "text",
    currency: bool = False,
    reconcile_key: bool = False,
) -> str:
    """Traduce una regola JSON in una frase operativa leggibile."""
    source_labels = {
        "body": "CORPO", "subject": "OGGETTO", "sender": "MITTENTE", "recipient": "DESTINATARIO",
        "date": "DATA MAIL", "message_id": "ID MAIL",
    }
    transform_labels = {
        "text": "testo",
        "upper": "testo in maiuscolo",
        "int": "numero intero",
        "money": "importo numerico",
        "guest_count": "numero totale di ospiti",
        "date_it_email_year": "data italiana (usando l'anno della mail se manca)",
    }
    type_labels = {"text": "TESTO", "date": "DATA", "number": "NUMERO"}
    source = source_labels.get(str(rule.get("source", "body")), "CORPO")
    field = str(rule.get("field", "campo") or "campo")
    strategy = str(rule.get("strategy", "regex") or "regex")
    occurrence = int(rule.get("occurrence", 1) or 1)

    if strategy == "regex":
        pattern = str(rule.get("pattern", "") or "")
        group = int(rule.get("group", 1) or 1)
        action = (
            f"Nel {source}, cerca la {occurrence}ª corrispondenza della regex «{pattern}» "
            f"e preleva il gruppo {group}."
        )
        expanded = expand_regex(pattern)
        if expanded != pattern:
            action += f" Prima dell'uso la regex viene espansa in «{expanded}»."
    elif strategy == "after_trigger":
        trigger = str(rule.get("trigger", "") or "")
        offset = int(rule.get("nonempty_offset", 1) or 1)
        action = (
            f"Nel {source}, trova la {occurrence}ª occorrenza del trigger «{trigger}» e preleva "
            f"la {offset}ª riga non vuota successiva."
        )
        value_pattern = str(rule.get("value_pattern", "") or "")
        if value_pattern:
            group = int(rule.get("group", 1) or 1)
            action += f" Sul valore trovato applica la regex «{value_pattern}» e usa il gruppo {group}."
            expanded_value = expand_regex(value_pattern)
            if expanded_value != value_pattern:
                action += f" La regex sul valore viene espansa in «{expanded_value}»."
    elif strategy == "constant":
        action = f"Non legge la mail: usa il valore costante «{rule.get('value', '')}»."
    else:
        action = f"Usa la strategia «{strategy}» sul {source}."

    transform_name = str(rule.get("transform", "text") or "text")
    transform_text = transform_labels.get(transform_name, transform_name)
    result = (
        f"{action} Interpreta il risultato come {transform_text} e salvalo nel campo «{field}» "
        f"come {type_labels.get(field_type, field_type.upper())}."
    )
    if currency and field_type == "number":
        result += " In Excel applica il formato valuta €."
    policy_labels = {
        "replace": "se la cella Excel contiene già un valore, lo sostituisce",
        "keep": "se la cella Excel contiene già un valore, lo mantiene e non lo sostituisce",
        "sum": "se la cella Excel contiene già un valore, somma il nuovo valore a quello esistente",
    }
    policy = str(rule.get("existing_value_policy", "replace") or "replace")
    result += " " + policy_labels.get(policy, policy_labels["replace"]) + "."
    valid_from = str(rule.get("valid_from", "") or "").strip()
    valid_to = str(rule.get("valid_to", "") or "").strip()
    if valid_from and valid_to:
        result += f" La regola si applica alle mail dal {valid_from} al {valid_to}, estremi inclusi."
    elif valid_from:
        result += f" La regola si applica alle mail con data dal {valid_from} in poi."
    elif valid_to:
        result += f" La regola si applica alle mail con data fino al {valid_to} compreso."
    else:
        result += " La regola non ha limiti temporali di validità."
    if reconcile_key:
        result += " Il campo è una chiave di riconciliazione tra più mail."
    return result


def compute_fields(
    record: dict[str, Any],
    profile: dict[str, Any],
    *,
    protected_fields: set[str] | None = None,
) -> None:
    """Calcola i campi derivati del profilo.

    ``protected_fields`` contiene i campi valorizzati esplicitamente da una
    regola della mail corrente. Per i campi calcolati con priorità
    ``rule_wins`` questi valori non vengono sovrascritti. Con priorità
    ``always`` il ricalcolo avviene comunque.
    """
    protected = set(protected_fields or ())

    def as_date(value: Any) -> date:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        text = str(value).strip()
        return date.fromisoformat(text[:10])

    for comp in profile.get("computed_fields", []):
        if not isinstance(comp, dict):
            continue
        destination = str(comp.get("field", "") or "")
        priority = str(comp.get("priority", "rule_wins") or "rule_wins").lower()
        if destination in protected and priority != "always":
            continue
        if comp.get("kind") == "days_between":
            a = record.get(comp.get("start_field"))
            b = record.get(comp.get("end_field"))
            if a and b and destination:
                try:
                    d1 = as_date(a)
                    d2 = as_date(b)
                    if d2 < d1 and (d1 - d2).days > 300:
                        d2 = d2.replace(year=d2.year + 1)
                        record[comp.get("end_field")] = d2.isoformat()
                    record[destination] = (d2 - d1).days
                except Exception:
                    pass


def extract_for_mail_type(mail: MailMessage, mail_type: dict[str, Any], profile: dict[str, Any]) -> dict[str, Any]:
    record: dict[str, Any] = {}
    for rule in mail_type.get("rules", []):
        value = extract_rule(mail, rule)
        if value not in (None, ""):
            record[rule.get("field")] = value
    compute_fields(record, profile, protected_fields=set(record))
    return record


def extract_message_all(mail: MailMessage, profile: dict[str, Any]) -> list[tuple[str, dict[str, Any]]]:
    """Applica le regole solo al primo Tipo mail che fa match, in ordine di priorità."""
    out: list[tuple[str, dict[str, Any]]] = []
    for mt in matching_mail_types(mail, profile):
        out.append((str(mt.get("name", "")), extract_for_mail_type(mail, mt, profile)))
    return out


def extract_message(mail: MailMessage, profile: dict[str, Any]) -> tuple[str | None, dict[str, Any]]:
    """Compatibilità: restituisce il primo tipo mail che fa match."""
    results = extract_message_all(mail, profile)
    return results[0] if results else (None, {})


def record_key(record: dict[str, Any], profile: dict[str, Any]) -> tuple[Any, ...] | None:
    keys = profile.get("reconcile_keys", [])
    if not keys:
        return None
    vals = tuple(record.get(k) for k in keys)
    if any(v in (None, "") for v in vals):
        return None
    return vals


def reconcile(records: list[dict[str, Any]], profile: dict[str, Any]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    by_key: dict[tuple[Any, ...], dict[str, Any]] = {}
    for rec in records:
        key = record_key(rec, profile)
        if key is None:
            out.append(dict(rec))
            continue
        if key not in by_key:
            by_key[key] = {}
            out.append(by_key[key])
        target = by_key[key]
        for k, v in rec.items():
            if v not in (None, ""):
                target[k] = v
        compute_fields(target, profile)
    return out


def selection_rule_results(mail: MailMessage, mail_type: dict[str, Any]) -> list[dict[str, Any]]:
    """Restituisce l'esito dettagliato di tutte le regole di selezione attive.

    Le regole Data sono attive anche senza trigger: usano valid_from/valid_to.
    Questa funzione deve restare semanticamente allineata a ``matches_mail_type``
    perché viene usata dalla pipeline dettagliata di Estrazione.
    """
    out: list[dict[str, Any]] = []
    for rule in mail_type.get("match_rules", []) or []:
        if not isinstance(rule, dict):
            continue
        source = str(rule.get("source", "body") or "body")
        trigger = str(rule.get("trigger", "") or "").strip()
        valid_from = str(rule.get("valid_from", "") or "").strip()
        valid_to = str(rule.get("valid_to", "") or "").strip()
        is_date_rule = source == "date" and bool(valid_from or valid_to)
        if not trigger and not is_date_rule:
            continue
        mode = str(rule.get("mode", "contains") or "contains")
        out.append({
            "name": str(rule.get("name", "") or ""),
            "source": source,
            "mode": mode,
            "trigger": trigger,
            "expanded_trigger": expand_regex(trigger) if (mode == "regex" and trigger) else trigger,
            "valid_from": valid_from,
            "valid_to": valid_to,
            "matched": bool(match_selection_rule(mail, rule)),
        })
    return out


def extract_mail_detailed(mail: MailMessage, profile: dict[str, Any]) -> dict[str, Any]:
    """Estrae una mail mantenendo il dettaglio necessario per audit/log.

    Il formato compatto storico (``extract_message_all``) resta disponibile;
    questa funzione aggiunge nome regola, campo, politica su valore esistente,
    esito e valore trasformato per ogni regola di estrazione.
    """
    type_results: list[dict[str, Any]] = []
    for mt in profile.get("mail_types", []) or []:
        if not isinstance(mt, dict):
            continue
        selection = selection_rule_results(mail, mt)
        # Un tipo senza trigger attivi non seleziona nessuna mail.
        if not selection or not all(item["matched"] for item in selection):
            continue
        updates: list[dict[str, Any]] = []
        candidate: dict[str, Any] = {}
        for rule in mt.get("rules", []) or []:
            if not isinstance(rule, dict):
                continue
            trace = extract_rule_trace(mail, rule)
            field = str(rule.get("field", "") or "")
            value = trace.value
            if field and value not in (None, ""):
                candidate[field] = value
            updates.append({
                "rule_name": str(rule.get("name", "") or ""),
                "field": field,
                "policy": str(rule.get("existing_value_policy", "replace") or "replace"),
                "matched": bool(trace.matched),
                "raw_value": trace.raw_value,
                "value": value,
                "error": trace.error,
                "source": str(rule.get("source", "body") or "body"),
                "strategy": str(rule.get("strategy", "regex") or "regex"),
                "trigger": str(rule.get("trigger", "") or ""),
                "pattern": str(rule.get("pattern", "") or ""),
                "expanded_pattern": expand_regex(str(rule.get("pattern", "") or "")),
                "value_pattern": str(rule.get("value_pattern", "") or ""),
                "expanded_value_pattern": expand_regex(str(rule.get("value_pattern", "") or "")),
            })
        compute_fields(candidate, profile, protected_fields=set(candidate))
        type_results.append({
            "mail_type": str(mt.get("name", "") or ""),
            "selection_rules": selection,
            "updates": updates,
            "candidate": candidate,
        })
        # Priorità sequenziale: il primo Tipo mail che fa match vince.
        break
    return {"mail": mail, "types": type_results}
