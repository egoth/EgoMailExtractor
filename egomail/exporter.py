from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .extraction import compute_fields, parse_money, record_key
from .profile import normalize_field_type


def field_names(profile: dict[str, Any]) -> list[str]:
    fields = profile.get("fields", [])
    names = []
    for f in fields:
        if isinstance(f, str):
            names.append(f)
        elif isinstance(f, dict) and f.get("name"):
            names.append(f["name"])
    return names


def field_map(profile: dict[str, Any]) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for f in profile.get("fields", []):
        if isinstance(f, str):
            out[f] = {"name": f, "type": "text"}
        elif isinstance(f, dict) and f.get("name"):
            item = dict(f)
            item["type"] = normalize_field_type(item.get("type"))
            out[item["name"]] = item
    return out




def _strict_number(value: Any) -> float | int | None:
    """Converte solo stringhe che sono realmente numeriche.

    Evita il vecchio comportamento pericoloso in cui un codice alfanumerico
    come ``HM895ZFM4S`` diventava ``8954`` perché venivano eliminati tutti i
    caratteri non numerici prima della conversione.
    """
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return value
    if value in (None, ""):
        return None
    text = str(value).replace("\xa0", " ").strip()
    # Sono ammessi suffissi valuta, spazi, segno, separatore migliaia e decimali.
    text = re.sub(r"\s*(?:EUR|€)\s*$", "", text, flags=re.I).strip()
    numeric_re = re.compile(
        r"^[+-]?(?:(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?|\d+(?:\.\d+)?)$"
    )
    if not numeric_re.fullmatch(text):
        return None
    parsed = parse_money(text)
    if parsed is None:
        return None
    if re.fullmatch(r"[+-]?\d+", text):
        try:
            return int(text)
        except ValueError:
            pass
    return parsed

def excel_value(value: Any, field: dict[str, Any]) -> Any:
    """Converte il valore nel vero tipo cella richiesto dal profilo."""
    if value in (None, ""):
        return None
    kind = normalize_field_type(field.get("type"))
    if kind == "text":
        # Forza il testo, utile anche per codici con zeri iniziali.
        return str(value)
    if kind == "number":
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, (int, float)):
            return value
        parsed = _strict_number(value)
        # Se il campo è stato configurato erroneamente come Numero ma contiene
        # un codice alfanumerico, non distruggiamo il dato: lo lasciamo testo.
        return parsed if parsed is not None else str(value)
    if kind == "date":
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        try:
            return date.fromisoformat(str(value)[:10])
        except Exception:
            return value
    return value


def final_formula_aliases(profile: dict[str, Any]) -> dict[str, str]:
    """Alias simbolici disponibili nelle formule riepilogative nominate.

    Esempi: ``somma_importo`` e ``conta_nome_ospite``. Il valore restituito
    è il nome del campo; il riferimento alla cella viene determinato in fase
    di scrittura perché dipende dall'ordine corrente delle colonne.
    """
    aliases: dict[str, str] = {}
    for field in profile.get("fields", []) or []:
        if not isinstance(field, dict) or not field.get("name"):
            continue
        name = str(field["name"])
        kind = str(field.get("final_formula", "") or "").lower()
        if kind == "sum":
            aliases[f"somma_{name}"] = name
        elif kind == "count":
            aliases[f"conta_{name}"] = name
    return aliases


def expand_summary_formula(
    formula: str,
    profile: dict[str, Any],
    alias_cells: dict[str, str],
) -> str:
    """Converte gli alias ``somma_*``/``conta_*`` in riferimenti Excel."""
    text = str(formula or "").strip()
    if not text:
        raise ValueError("Formula riepilogativa vuota")
    if not text.startswith("="):
        text = "=" + text

    available = {k.casefold(): v for k, v in alias_cells.items()}
    token_re = re.compile(r"\b(?:somma|conta)_[A-Za-z0-9_À-ÖØ-öø-ÿ]+\b", re.I)
    unknown: set[str] = set()

    def repl(match: re.Match[str]) -> str:
        token = match.group(0)
        cell = available.get(token.casefold())
        if cell is None:
            unknown.add(token)
            return token
        return cell

    expanded = token_re.sub(repl, text)
    if unknown:
        valid = ", ".join(sorted(alias_cells)) or "<nessun alias disponibile>"
        raise ValueError(
            "Alias formula non definiti: " + ", ".join(sorted(unknown)) +
            ". Alias disponibili: " + valid
        )
    return expanded


def _is_generated_summary_row(
    ws,
    row_idx: int,
    fields: list[str],
    definitions: dict[str, dict[str, Any]],
    profile: dict[str, Any] | None = None,
) -> bool:
    """Riconosce righe riepilogative generate da EgoMailExtractor."""
    # Riga unica delle formule finali associate ai campi.
    for col_idx, field_name in enumerate(fields, 1):
        kind = str(definitions.get(field_name, {}).get("final_formula", "") or "").lower()
        if kind not in {"sum", "count"}:
            continue
        value = ws.cell(row_idx, col_idx).value
        if not isinstance(value, str) or not value.startswith("="):
            continue
        upper = value.upper()
        if kind == "sum" and upper.startswith("=SUM("):
            return True
        if kind == "count" and upper.startswith("=COUNTA("):
            return True

    # Righe delle formule riepilogative nominate: nome in colonna A e formula
    # Excel generata in colonna B.
    if profile:
        names = {
            str(item.get("name", "") or "").strip().casefold()
            for item in profile.get("summary_formulas", []) or []
            if isinstance(item, dict) and item.get("name")
        }
        if names:
            label = str(ws.cell(row_idx, 1).value or "").strip().casefold()
            formula = ws.cell(row_idx, 2).value
            if label in names and isinstance(formula, str) and formula.startswith("="):
                return True
    return False

def load_existing_records(path: str | Path, profile: dict[str, Any]) -> list[dict[str, Any]]:
    p = Path(path)
    if not p.exists():
        return []
    # Apriamo due viste: una conserva le formule per riconoscere la riga
    # riepilogativa, l'altra restituisce i valori calcolati delle celle normali.
    wb_formula = load_workbook(p, data_only=False)
    wb_values = load_workbook(p, data_only=True)
    ws_formula = wb_formula["Dati"] if "Dati" in wb_formula.sheetnames else wb_formula.active
    ws_values = wb_values["Dati"] if "Dati" in wb_values.sheetnames else wb_values.active
    headers = [c.value for c in ws_formula[1]]
    fields = [str(h) for h in headers if h]
    definitions = field_map(profile)
    out = []
    max_row = max(ws_formula.max_row, ws_values.max_row)
    for row_idx in range(2, max_row + 1):
        if _is_generated_summary_row(ws_formula, row_idx, fields, definitions, profile):
            continue
        row = [ws_values.cell(row_idx, i + 1).value for i in range(len(headers))]
        rec = {str(headers[i]): row[i] for i in range(min(len(headers), len(row))) if headers[i]}
        if any(v not in (None, "") for v in rec.values()):
            out.append(rec)
    return out


def merge_records(existing: list[dict[str, Any]], new: list[dict[str, Any]], profile: dict[str, Any]) -> list[dict[str, Any]]:
    out = [dict(x) for x in existing]
    index: dict[tuple[Any, ...], dict[str, Any]] = {}
    for rec in out:
        k = record_key(rec, profile)
        if k is not None:
            index[k] = rec
    for rec in new:
        k = record_key(rec, profile)
        if k is not None and k in index:
            target = index[k]
            for field, value in rec.items():
                if value not in (None, ""):
                    target[field] = value
            compute_fields(target, profile)
        else:
            copy = dict(rec)
            compute_fields(copy, profile)
            out.append(copy)
            if k is not None:
                index[k] = copy
    return out



def _write_totals_table(wb, records: list[dict[str, Any]], profile: dict[str, Any]) -> None:
    cfg = profile.get("totals_table", {}) or {}
    rows = [str(x) for x in cfg.get("row_fields", []) if x]
    cols = [str(x) for x in cfg.get("column_fields", []) if x]
    values = [str(x) for x in cfg.get("value_fields", []) if x]
    date_field = str(cfg.get("date_field", "") or "")
    group_year = bool(cfg.get("group_year")); group_month = bool(cfg.get("group_month"))
    if not (rows or cols or values or (date_field and (group_year or group_month))):
        return
    ws = wb.create_sheet("Tabella totali")
    row_dims = list(rows)
    if date_field and group_year: row_dims.append(f"{date_field} - Anno")
    if date_field and group_month: row_dims.append(f"{date_field} - Mese")
    def date_parts(rec):
        v=rec.get(date_field) if date_field else None
        if isinstance(v, datetime): d=v.date()
        elif isinstance(v,date): d=v
        else:
            try: d=date.fromisoformat(str(v)[:10])
            except Exception: d=None
        return (d.year if d and group_year else "", d.month if d and group_month else "")
    def rkey(rec):
        out=[rec.get(x,"") for x in rows]; y,m=date_parts(rec)
        if group_year: out.append(y)
        if group_month: out.append(m)
        return tuple(out)
    def ckey(rec): return tuple(rec.get(x,"") for x in cols)
    rowkeys=sorted({rkey(r) for r in records}, key=lambda x: tuple(str(v) for v in x)) or [tuple()]
    colkeys=sorted({ckey(r) for r in records}, key=lambda x: tuple(str(v) for v in x)) or [tuple()]
    metrics=values or ["Conteggio"]
    headers=row_dims[:]
    for ck in colkeys:
        prefix=" / ".join(str(v) for v in ck if v not in (None,""))
        for metric in metrics: headers.append((prefix+" / " if prefix else "")+metric)
    ws.append(headers)
    for cell in ws[1]: cell.font=Font(bold=True); cell.fill=PatternFill("solid",fgColor="D9EAF7")
    for rk in rowkeys:
        line=list(rk)
        for ck in colkeys:
            subset=[r for r in records if rkey(r)==rk and ckey(r)==ck]
            for metric in metrics:
                if metric=="Conteggio": val=len(subset)
                else:
                    nums=[_strict_number(r.get(metric)) for r in subset]; val=sum(x for x in nums if x is not None)
                line.append(val)
        ws.append(line)
    ws.freeze_panes="A2"
    for i in range(1,ws.max_column+1): ws.column_dimensions[get_column_letter(i)].width=min(40,max(12,max(len(str(ws.cell(r,i).value or "")) for r in range(1,ws.max_row+1))+2))


def build_period_summary_formula(
    item: dict[str, Any],
    fields: list[str],
    data_last_row: int,
) -> str:
    """Genera SOMMA/CONTA condizionata da una finestra temporale scelta dall'utente."""
    operation = str(item.get("operation", "sum") or "sum").lower()
    value_field = str(item.get("field", "") or "")
    date_field = str(item.get("date_field", "") or "")
    period = str(item.get("period", "month") or "month").lower()
    period_mode = str(item.get("period_mode", "rolling") or "rolling").lower()
    if value_field not in fields or date_field not in fields:
        return "=0"
    if data_last_row < 2:
        return "=0"
    value_col = get_column_letter(fields.index(value_field) + 1)
    date_col = get_column_letter(fields.index(date_field) + 1)
    value_range = f"{value_col}2:{value_col}{data_last_row}"
    date_range = f"{date_col}2:{date_col}{data_last_row}"

    if period_mode == "calendar":
        if period == "year":
            start_expr = "DATE(YEAR(TODAY()),1,1)"
            end_expr = "DATE(YEAR(TODAY()),12,31)"
        else:
            start_expr = "EOMONTH(TODAY(),-1)+1"
            end_expr = "EOMONTH(TODAY(),0)"
    else:
        days = 365 if period == "year" else 30
        start_expr = f"TODAY()-{days}"
        end_expr = "TODAY()"

    if operation == "count":
        return (
            f'=COUNTIFS({date_range},">="&{start_expr},{date_range},"<="&{end_expr},'
            f'{value_range},"<>")'
        )
    return (
        f'=SUMIFS({value_range},{date_range},">="&{start_expr},'
        f'{date_range},"<="&{end_expr})'
    )


def summary_formula_display(item: dict[str, Any]) -> str:
    if str(item.get("kind", "") or "") != "period_aggregate":
        return str(item.get("formula", "") or "")
    op = "Somma" if str(item.get("operation", "sum") or "sum") == "sum" else "Conta"
    period = "anno" if str(item.get("period", "month") or "month") == "year" else "mese"
    mode = "finestra mobile" if str(item.get("period_mode", "rolling") or "rolling") == "rolling" else "periodo di calendario"
    return f"{op} {item.get('field','')} se {item.get('date_field','')} è nell'ultimo {period} ({mode})"

def write_excel(path: str | Path, records: list[dict[str, Any]], profile: dict[str, Any]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Dati"
    fields = field_names(profile)
    definitions = field_map(profile)
    ws.append(fields)
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    for rec in records:
        ws.append([excel_value(rec.get(f), definitions.get(f, {"name": f, "type": "text"})) for f in fields])

    data_last_row = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=data_last_row):
        for idx, cell in enumerate(row):
            field = fields[idx]
            spec = definitions.get(field, {"type": "text"})
            kind = normalize_field_type(spec.get("type"))
            if spec.get("format") == "currency_eur" and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00 [$€-it-IT]'
            elif kind == "date" and isinstance(cell.value, (date, datetime)):
                cell.number_format = "dd/mm/yyyy"
            elif kind == "number" and isinstance(cell.value, (int, float)):
                cell.number_format = "0.00" if isinstance(cell.value, float) and not float(cell.value).is_integer() else "0"

    # Una sola riga riepilogativa, immediatamente dopo l'ultimo record. Ogni
    # campo può decidere indipendentemente se scrivere SOMMA o CONTA.
    formula_fields = [
        (
            i,
            f,
            str(definitions.get(f, {}).get("final_formula", "") or "").lower(),
            str(definitions.get(f, {}).get("summary_source_field", "") or f),
        )
        for i, f in enumerate(fields, 1)
        if str(definitions.get(f, {}).get("final_formula", "") or "").lower() in {"sum", "count"}
    ]
    summary_row: int | None = None
    alias_cells: dict[str, str] = {}
    if formula_fields:
        summary_row = data_last_row + 1
        for col_idx, field_name, formula_kind, source_field in formula_fields:
            col = get_column_letter(col_idx)
            try:
                source_col_idx = fields.index(source_field) + 1
            except ValueError:
                source_col_idx = col_idx
            source_col = get_column_letter(source_col_idx)
            if data_last_row >= 2:
                if formula_kind == "sum":
                    formula = f"=SUM({source_col}2:{source_col}{data_last_row})"
                else:
                    formula = f"=COUNTA({source_col}2:{source_col}{data_last_row})"
            else:
                formula = "=0"
            cell = ws.cell(summary_row, col_idx, formula)
            cell.font = Font(bold=True)
            spec = definitions.get(field_name, {"type": "text"})
            if formula_kind == "sum" and spec.get("format") == "currency_eur":
                cell.number_format = '#,##0.00 [$€-it-IT]'
            elif formula_kind == "sum":
                cell.number_format = "0.00" if normalize_field_type(spec.get("type")) == "number" else "0"
            else:
                cell.number_format = "0"
            alias_name = ("somma_" if formula_kind == "sum" else "conta_") + field_name
            alias_cells[alias_name] = f"{col}{summary_row}"

    # Formule riepilogative nominate. Sono scritte in righe successive con il
    # nome in colonna A e il risultato/formula in colonna B. Nella definizione
    # l'utente usa alias leggibili, ad esempio:
    # =somma_importo/somma_durata_giorni
    custom_formulas = [
        item for item in profile.get("summary_formulas", []) or []
        if isinstance(item, dict) and str(item.get("name", "") or "").strip()
        and (
            str(item.get("formula", "") or "").strip()
            or str(item.get("kind", "") or "") == "period_aggregate"
        )
    ]
    if custom_formulas:
        first_custom_row = (summary_row + 1) if summary_row is not None else (data_last_row + 1)
        for offset, item in enumerate(custom_formulas):
            row_idx = first_custom_row + offset
            name = str(item.get("name", "") or "").strip()
            if str(item.get("kind", "") or "") == "period_aggregate":
                expanded = build_period_summary_formula(item, fields, data_last_row)
            else:
                expanded = expand_summary_formula(str(item.get("formula", "") or ""), profile, alias_cells)
            name_cell = ws.cell(row_idx, 1, name)
            value_cell = ws.cell(row_idx, 2, expanded)
            name_cell.font = Font(bold=True)
            value_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    # Il filtro comprende solo intestazione + dati, non la riga di formule.
    if fields:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(fields))}{max(1, data_last_row)}"
    for i, f in enumerate(fields, 1):
        values = [str(f)] + [str(ws.cell(r, i).value or "") for r in range(2, data_last_row + 1)]
        width = min(45, max(12, max(len(v) for v in values) + 2))
        ws.column_dimensions[get_column_letter(i)].width = width

    _write_totals_table(wb, records, profile)

    # Chiede a Excel/LibreOffice di ricalcolare le formule all'apertura.
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    wb.save(p)



def read_excel_home_summary_values(
    path: str | Path,
    profile: dict[str, Any],
    selected_fields: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Legge dall'Excel i valori delle formule finali scelti per la home.

    I valori vengono letti con ``data_only=True`` dalla riga riepilogativa.
    Se Excel non ha ancora valorizzato la cache delle formule, su Windows viene
    chiesto a Microsoft Excel di ricalcolare e salvare il file, quindi la
    lettura viene ripetuta. In questo modo la home mostra il valore effettivo
    prodotto da Excel, non un calcolo parallelo dell'applicazione.
    """
    p = Path(path)
    if not p.exists():
        return []
    selected = [str(x) for x in (selected_fields if selected_fields is not None else profile.get("home_summary_fields", [])) if x][:6]
    if not selected:
        return []
    definitions = field_map(profile)

    def _read() -> list[dict[str, Any]]:
        wb = load_workbook(p, data_only=True, read_only=False)
        try:
            ws = wb["Dati"] if "Dati" in wb.sheetnames else wb.active
            headers = [str(c.value or "") for c in ws[1]]
            if "_EgoMeta" in wb.sheetnames:
                try:
                    data_last_row = int(wb["_EgoMeta"]["A1"].value or 1)
                except Exception:
                    data_last_row = max(1, ws.max_row - 1)
            else:
                data_last_row = max(1, ws.max_row - 1)
            summary_row = data_last_row + 1
            result: list[dict[str, Any]] = []
            summary_defs = {
                str(item.get("name", "") or ""): item
                for item in profile.get("summary_formulas", []) or []
                if isinstance(item, dict) and item.get("name")
            }
            computed_names = {
                str(item.get("field", "") or "")
                for item in profile.get("computed_fields", []) or []
                if isinstance(item, dict) and item.get("field")
            }
            for ref in selected:
                ref = str(ref)
                source_kind, _, raw_name = ref.partition(":")
                if not raw_name:
                    source_kind, raw_name = "field", ref

                if source_kind == "summary":
                    # Le formule riepilogative nominate sono scritte in colonna A/B
                    # dopo la riga delle formule finali.
                    value = None
                    for row_idx in range(summary_row, ws.max_row + 1):
                        if str(ws.cell(row_idx, 1).value or "") == raw_name:
                            value = ws.cell(row_idx, 2).value
                            break
                    result.append({
                        "field": ref, "label": raw_name, "kind": "summary",
                        "value": value, "format": "",
                    })
                    continue

                if raw_name not in headers:
                    continue
                spec = definitions.get(raw_name, {})
                col_idx = headers.index(raw_name) + 1
                final_kind = str(spec.get("final_formula", "") or "").lower()
                if final_kind in {"sum", "count"}:
                    value = ws.cell(summary_row, col_idx).value
                    result.append({
                        "field": ref, "label": raw_name, "kind": final_kind,
                        "value": value, "format": str(spec.get("format", "") or ""),
                    })
                    continue

                # Un campo calcolato privo di Formula finale mostra l'ultimo
                # valore non vuoto prodotto nelle righe dati.
                if source_kind == "computed" or raw_name in computed_names:
                    value = None
                    for row_idx in range(data_last_row, 1, -1):
                        candidate = ws.cell(row_idx, col_idx).value
                        if candidate not in (None, ""):
                            value = candidate
                            break
                    result.append({
                        "field": ref, "label": raw_name, "kind": "computed",
                        "value": value, "format": str(spec.get("format", "") or ""),
                    })
            return result
        finally:
            wb.close()

    result = _read()
    if result and all(item.get("value") is not None for item in result):
        return result

    # La cache formula può essere vuota subito dopo un salvataggio openpyxl.
    # Se possibile, facciamo ricalcolare il file a Microsoft Excel.
    import sys
    if sys.platform == "win32":
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            excel = None
            wb_com = None
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb_com = excel.Workbooks.Open(str(p.resolve()))
                try:
                    excel.CalculateFullRebuild()
                except Exception:
                    excel.Calculate()
                wb_com.Save()
            finally:
                try:
                    if wb_com is not None:
                        wb_com.Close(SaveChanges=True)
                except Exception:
                    pass
                try:
                    if excel is not None:
                        excel.Quit()
                except Exception:
                    pass
                pythoncom.CoUninitialize()
            result = _read()
        except Exception:
            pass
    return result


def _is_empty(value: Any) -> bool:
    return value in (None, "")


def _as_number(value: Any) -> float | int | None:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return value
    return _strict_number(value)


def apply_existing_value_policy(
    old_value: Any,
    new_value: Any,
    policy: str,
    field_spec: dict[str, Any],
) -> tuple[Any, str]:
    """Applica la politica della regola quando la cella contiene già un valore.

    Restituisce ``(valore_finale, descrizione_azione)`` per poter riportare
    esattamente l'operazione nel log di estrazione.
    """
    if _is_empty(new_value):
        return old_value, "NESSUN VALORE ESTRATTO"
    if _is_empty(old_value):
        return new_value, "SCRIVI (cella vuota)"

    policy = str(policy or "replace").strip().lower()
    if policy == "keep":
        return old_value, "NON SOSTITUIRE"
    if policy == "sum":
        if normalize_field_type(field_spec.get("type")) != "number":
            return new_value, "SOSTITUISCI (Somma non ammessa: campo non numerico)"
        old_num = _as_number(old_value)
        new_num = _as_number(new_value)
        if old_num is None or new_num is None:
            return old_value, "SOMMA NON ESEGUITA (valore non numerico)"
        total = old_num + new_num
        # Conserva un int quando possibile. Per gli importi evita gli artefatti
        # floating point (es. 650.8599999999999) arrotondando ai centesimi.
        if field_spec.get("format") == "currency_eur":
            total = round(float(total), 2)
        elif isinstance(old_num, int) and isinstance(new_num, int):
            total = int(total)
        return total, "SOMMA"
    return new_value, "SOSTITUISCI"


def _log_value(value: Any) -> str:
    if value is None:
        return "<vuoto>"
    text = str(value).replace("\r", " ").replace("\n", " ")
    return text if len(text) <= 500 else text[:497] + "..."


def _record_key_text(record: dict[str, Any], profile: dict[str, Any]) -> str:
    keys = profile.get("reconcile_keys", []) or []
    if not keys:
        return "nessuna chiave di riconciliazione"
    return ", ".join(f"{k}={_log_value(record.get(k))}" for k in keys)


def apply_detailed_extractions(
    existing: list[dict[str, Any]],
    mail_events: list[dict[str, Any]],
    profile: dict[str, Any],
    *,
    log_path: str | Path | None = None,
    mode: str = "estrazione",
    excel_path: str | Path | None = None,
    include_header: bool = True,
    include_footer: bool = True,
) -> tuple[list[dict[str, Any]], dict[str, int], str]:
    """Applica gli eventi dettagliati in ordine e produce un audit log.

    ``mail_events`` deve essere già ordinato dalla mail più vecchia alla più
    recente. Le politiche ``replace``/``keep``/``sum`` sono memorizzate sulla
    singola regola di estrazione, perciò due regole che scrivono lo stesso
    campo possono avere comportamenti differenti.
    """
    out = [dict(x) for x in existing]
    definitions = field_map(profile)
    index: dict[tuple[Any, ...], int] = {}
    for i, rec in enumerate(out):
        key = record_key(rec, profile)
        if key is not None:
            index[key] = i

    now = datetime.now().astimezone()
    lines: list[str] = []
    if include_header:
        lines.extend([
            "=" * 100,
            f"INIZIO {mode.upper()} - {now.isoformat(timespec='seconds')}",
            f"Profilo: {profile.get('name', '')}",
            f"File Excel: {Path(excel_path).resolve() if excel_path else ''}",
            f"Mail in scansione: {len(mail_events)}",
            "Ordine elaborazione: dalla mail più vecchia alla più recente",
            "=" * 100,
        ])
    stats = {"mails": len(mail_events), "matched_mails": 0, "type_matches": 0, "field_updates": 0}

    for event in mail_events:
        mail = event.get("mail")
        types = event.get("types", []) or []
        lines.extend([
            "",
            "-" * 100,
            "MAIL ESAMINATA",
            f"UID: {getattr(mail, 'uid', '')}",
            f"Message-ID: {getattr(mail, 'message_id', '') or '<assente>'}",
            f"Data/ora mail: {getattr(mail, 'date', '')}",
            f"Mittente: {getattr(mail, 'sender', '')}",
            f"Oggetto: {getattr(mail, 'subject', '')}",
            f"Selezionata per estrazione: {'SI' if types else 'NO'}",
        ])
        if not types:
            lines.append("Nessun tipo mail ha soddisfatto tutte le proprie regole di selezione.")
            continue
        stats["matched_mails"] += 1

        for type_event in types:
            stats["type_matches"] += 1
            mail_type = type_event.get("mail_type", "")
            selection = type_event.get("selection_rules", []) or []
            updates = type_event.get("updates", []) or []
            candidate = dict(type_event.get("candidate", {}) or {})
            extracted_nonempty = any(
                u.get("field") and u.get("value") not in (None, "") for u in updates
            )
            counter_fields = [
                (field_name, spec) for field_name, spec in definitions.items()
                if str(spec.get("mail_type_counter", "") or "") == str(mail_type)
            ]

            lines.extend([
                "",
                f"TIPO MAIL MATCH: {mail_type}",
                "Regole di selezione che hanno determinato il match:",
            ])
            for sr in selection:
                lines.append(
                    "  [MATCH] "
                    f"{sr.get('name', '')} | origine={sr.get('source', '')} | "
                    f"modalità={sr.get('mode', '')} | trigger={_log_value(sr.get('trigger', ''))}"
                )
                if sr.get("mode") == "regex" and sr.get("expanded_trigger") != sr.get("trigger"):
                    lines.append(
                        f"          regex espansa={_log_value(sr.get('expanded_trigger', ''))}"
                    )

            if not extracted_nonempty and not counter_fields:
                lines.append("Riga Excel: nessuna (le regole di estrazione non hanno prodotto valori).")
                lines.append("Regole di estrazione:")
                for upd in updates:
                    status = "MATCH" if upd.get("matched") else "NO MATCH"
                    lines.append(
                        f"  [{status}] {upd.get('rule_name', '')} -> campo={upd.get('field', '')} | "
                        f"valore={_log_value(upd.get('value'))} | errore={upd.get('error', '') or '<nessuno>'}"
                    )
                continue
            if not extracted_nonempty and counter_fields and profile.get("reconcile_keys"):
                lines.append(
                    "Riga Excel: nessuna. Il Tipo mail ha un contatore automatico, ma senza "
                    "un valore estratto non è possibile individuare la riga tramite la chiave di riconciliazione."
                )
                continue

            key = record_key(candidate, profile)
            if key is not None and key in index:
                row_index = index[key]
                target = out[row_index]
                row_status = "riga esistente riconciliata"
            else:
                target = {}
                out.append(target)
                row_index = len(out) - 1
                row_status = "nuova riga"
                if key is not None:
                    index[key] = row_index
            excel_row = row_index + 2
            lines.append(
                f"Riga Excel: {excel_row} ({row_status}) | {_record_key_text(candidate, profile)}"
            )
            lines.append("Regole di estrazione e campi:")

            for upd in updates:
                rule_name = str(upd.get("rule_name", "") or "")
                field = str(upd.get("field", "") or "")
                value = upd.get("value")
                matched = bool(upd.get("matched"))
                policy = str(upd.get("policy", "replace") or "replace")
                if not field:
                    lines.append(f"  [IGNORA] {rule_name}: campo destinazione non impostato")
                    continue
                if value in (None, ""):
                    status = "MATCH SENZA VALORE" if matched else "NO MATCH"
                    lines.append(
                        f"  [{status}] regola={rule_name} | campo={field} | "
                        f"strategia={upd.get('strategy', '')} | valore=<vuoto> | "
                        f"errore={upd.get('error', '') or '<nessuno>'}"
                    )
                    continue

                spec = definitions.get(field, {"name": field, "type": "text"})
                old = target.get(field)
                final, action = apply_existing_value_policy(old, value, policy, spec)
                target[field] = final
                stats["field_updates"] += 1
                lines.extend([
                    f"  [ESTRAI] regola={rule_name} | campo={field} | riga Excel={excel_row}",
                    f"           origine={upd.get('source', '')} | strategia={upd.get('strategy', '')}",
                    f"           valore estratto={_log_value(value)}",
                    f"           politica su valore esistente={policy}",
                    f"           valore prima={_log_value(old)}",
                    f"           azione={action}",
                    f"           valore dopo={_log_value(final)}",
                ])
                if upd.get("strategy") == "regex" and upd.get("expanded_pattern") != upd.get("pattern"):
                    lines.append(
                        f"           regex espansa={_log_value(upd.get('expanded_pattern', ''))}"
                    )
                if upd.get("expanded_value_pattern") and upd.get("expanded_value_pattern") != upd.get("value_pattern"):
                    lines.append(
                        f"           regex valore espansa={_log_value(upd.get('expanded_value_pattern', ''))}"
                    )

            # Contatori automatici dei match Tipo mail: una mail incrementa di
            # una unità ogni campo associato al tipo, indipendentemente dal
            # numero di regole di estrazione che hanno prodotto valori.
            for counter_field, counter_spec in counter_fields:
                old_counter = target.get(counter_field)
                old_num = _as_number(old_counter)
                if old_num is None:
                    old_num = 0
                final_counter = int(old_num) + 1
                target[counter_field] = final_counter
                stats["field_updates"] += 1
                lines.extend([
                    f"  [CONTA TIPO MAIL] tipo={mail_type} | campo={counter_field} | riga Excel={excel_row}",
                    f"           valore prima={_log_value(old_counter)}",
                    f"           azione=INCREMENTA MATCH TIPO MAIL",
                    f"           valore dopo={final_counter}",
                ])

            explicit_fields = {
                str(upd.get("field", "") or "")
                for upd in updates
                if upd.get("field") and upd.get("value") not in (None, "")
            }
            computed_before = {
                str(comp.get("field", "") or ""): target.get(str(comp.get("field", "") or ""))
                for comp in profile.get("computed_fields", []) or []
                if isinstance(comp, dict) and comp.get("field")
            }
            compute_fields(target, profile, protected_fields=explicit_fields)
            for comp in profile.get("computed_fields", []) or []:
                if not isinstance(comp, dict):
                    continue
                computed_field = str(comp.get("field", "") or "")
                if not computed_field:
                    continue
                priority = str(comp.get("priority", "rule_wins") or "rule_wins")
                old_computed = computed_before.get(computed_field)
                new_computed = target.get(computed_field)
                if computed_field in explicit_fields and priority != "always":
                    lines.extend([
                        f"  [CAMPO CALCOLATO] campo={computed_field} | tipo={comp.get('kind', '')}",
                        "           azione=NON RICALCOLATO (priorità al valore della regola corrente)",
                        f"           valore mantenuto={_log_value(new_computed)}",
                    ])
                elif new_computed != old_computed:
                    lines.extend([
                        f"  [CAMPO CALCOLATO] campo={computed_field} | tipo={comp.get('kind', '')}",
                        f"           valore prima={_log_value(old_computed)}",
                        "           azione=RICALCOLA",
                        f"           valore dopo={_log_value(new_computed)}",
                    ])
            # La chiave può essere stata popolata durante gli aggiornamenti.
            final_key = record_key(target, profile)
            if final_key is not None:
                index[final_key] = row_index

    end = datetime.now().astimezone()
    if include_footer:
        lines.extend([
            "",
            "=" * 100,
            f"FINE {mode.upper()} - {end.isoformat(timespec='seconds')}",
            f"Mail esaminate: {stats['mails']}",
            f"Mail con almeno un tipo matchato: {stats['matched_mails']}",
            f"Match di tipo mail: {stats['type_matches']}",
            f"Aggiornamenti campo eseguiti: {stats['field_updates']}",
            f"Righe dati finali: {len(out)}",
            "=" * 100,
            "",
        ])
    log_text = "\n".join(lines)
    if log_path:
        append_extraction_log(log_path, log_text)
    return out, stats, log_text


def append_extraction_log(log_path: str | Path, text: str) -> None:
    lp = Path(log_path)
    lp.parent.mkdir(parents=True, exist_ok=True)
    with lp.open("a", encoding="utf-8") as fh:
        fh.write(text)
        if not text.endswith("\n"):
            fh.write("\n")
