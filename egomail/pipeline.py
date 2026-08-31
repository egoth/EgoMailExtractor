from __future__ import annotations

import json
import shutil
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any

from .mail import MailHeader, MailMessage


PIPELINE_VERSION = 1


def atomic_write_json(path: str | Path, data: dict[str, Any]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(p)


def read_json(path: str | Path) -> dict[str, Any] | None:
    p = Path(path)
    if not p.exists():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def header_to_dict(header: MailHeader | MailMessage) -> dict[str, Any]:
    return {
        "uid": int(getattr(header, "uid", 0) or 0),
        "sender": str(getattr(header, "sender", "") or ""),
        "subject": str(getattr(header, "subject", "") or ""),
        "date": str(getattr(header, "date", "") or ""),
        "message_id": str(getattr(header, "message_id", "") or ""),
        "recipient": str(getattr(header, "recipient", "") or ""),
        "raw_date": str(getattr(header, "raw_date", "") or ""),
    }


def dict_to_header(data: dict[str, Any]) -> MailHeader:
    return MailHeader(
        uid=int(data.get("uid", 0) or 0),
        sender=str(data.get("sender", "") or ""),
        subject=str(data.get("subject", "") or ""),
        date=str(data.get("date", "") or ""),
        message_id=str(data.get("message_id", "") or ""),
        recipient=str(data.get("recipient", "") or ""),
    )


def mail_to_dict(mail: MailMessage) -> dict[str, Any]:
    out = header_to_dict(mail)
    out["body"] = str(getattr(mail, "body", "") or "")
    return out


def dict_to_mail(data: dict[str, Any]) -> MailMessage:
    return MailMessage(
        uid=int(data.get("uid", 0) or 0),
        sender=str(data.get("sender", "") or ""),
        subject=str(data.get("subject", "") or ""),
        date=str(data.get("date", "") or ""),
        message_id=str(data.get("message_id", "") or ""),
        recipient=str(data.get("recipient", "") or ""),
        body=str(data.get("body", "") or ""),
        raw_date=str(data.get("raw_date", "") or ""),
    )


def event_to_dict(event: dict[str, Any], profile: dict[str, Any]) -> dict[str, Any]:
    mail = event.get("mail")
    reconcile_keys = [str(x) for x in profile.get("reconcile_keys", []) or []]
    operations: list[dict[str, Any]] = []
    for type_event in event.get("types", []) or []:
        candidate = dict(type_event.get("candidate", {}) or {})
        for upd in type_event.get("updates", []) or []:
            field = str(upd.get("field", "") or "")
            value = upd.get("value")
            if not field or value in (None, ""):
                continue
            operations.append({
                "mail_type": str(type_event.get("mail_type", "") or ""),
                "rule_name": str(upd.get("rule_name", "") or ""),
                "field": field,
                "value": value,
                "policy": str(upd.get("policy", "replace") or "replace"),
                "requires_reconciliation": bool(reconcile_keys),
                "reconcile_keys": {
                    key: candidate.get(key) for key in reconcile_keys
                },
            })
    return {
        "mail": mail_to_dict(mail) if isinstance(mail, MailMessage) else {},
        "types": event.get("types", []) or [],
        "operations": operations,
    }


def dict_to_event(data: dict[str, Any]) -> dict[str, Any]:
    return {
        "mail": dict_to_mail(dict(data.get("mail", {}) or {})),
        "types": list(data.get("types", []) or []),
    }


def safe_unlink(path: str | Path) -> None:
    try:
        Path(path).unlink(missing_ok=True)
    except Exception:
        pass


def transactional_paths(final_path: str | Path) -> tuple[Path, Path]:
    p = Path(final_path)
    return (
        p.with_name(p.stem + " - originale" + p.suffix),
        p.with_name(p.stem + " - temporaneo" + p.suffix),
    )


def prepare_transaction_source(final_path: str | Path, create_empty_callback) -> tuple[Path, Path]:
    final = Path(final_path)
    original, temporary = transactional_paths(final)
    temporary.unlink(missing_ok=True)
    if original.exists():
        return original, temporary
    if final.exists():
        final.replace(original)
    else:
        create_empty_callback(original)
    shutil.copy2(original, temporary)
    return original, temporary
