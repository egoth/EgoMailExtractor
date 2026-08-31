from __future__ import annotations

import email
import html
import imaplib
import re
import ssl
import threading
from dataclasses import dataclass
from email.header import decode_header, make_header
from email.message import Message
from email.utils import parsedate_to_datetime
from typing import Callable, Iterable

from bs4 import BeautifulSoup

from .config import ConfigManager, ImapSettings


@dataclass
class MailHeader:
    uid: int
    sender: str
    subject: str
    date: str
    message_id: str = ""
    recipient: str = ""


@dataclass
class MailMessage(MailHeader):
    body: str = ""
    raw_date: str = ""

    @property
    def year(self) -> int | None:
        try:
            return parsedate_to_datetime(self.raw_date or self.date).year
        except Exception:
            return None


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def html_to_text(value: str) -> str:
    soup = BeautifulSoup(value, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text("\n")
    return normalize_text(text)


def normalize_text(text: str) -> str:
    text = html.unescape(text or "").replace("\xa0", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    out: list[str] = []
    last_blank = False
    for line in text.split("\n"):
        line = re.sub(r"[ \t]+", " ", line).strip()
        # Elimina elementi puramente decorativi tipici delle conversioni HTML/Markdown.
        if re.fullmatch(r"[|:\- ]+", line or " "):
            line = ""
        line = re.sub(r"^\*\*(.*?)\*\*$", r"\1", line)
        line = re.sub(r"^#{1,6}\s*", "", line)
        if not line:
            if not last_blank and out:
                out.append("")
            last_blank = True
        else:
            out.append(line)
            last_blank = False
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out)


def message_body(msg: Message) -> str:
    plain_parts: list[str] = []
    html_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", ""))
            if "attachment" in disp.lower():
                continue
            if ctype not in ("text/plain", "text/html"):
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                value = payload.decode(charset, errors="replace")
            except LookupError:
                value = payload.decode("utf-8", errors="replace")
            if ctype == "text/plain":
                plain_parts.append(value)
            else:
                html_parts.append(value)
    else:
        payload = msg.get_payload(decode=True)
        if payload is not None:
            charset = msg.get_content_charset() or "utf-8"
            try:
                value = payload.decode(charset, errors="replace")
            except LookupError:
                value = payload.decode("utf-8", errors="replace")
            if msg.get_content_type() == "text/html":
                html_parts.append(value)
            else:
                plain_parts.append(value)
    if plain_parts:
        return normalize_text("\n".join(plain_parts))
    return html_to_text("\n".join(html_parts))


def imap_fetch_payload(chunks) -> bytes:
    """Restituisce il payload reale di una risposta FETCH IMAP.

    ``imaplib`` normalmente restituisce il contenuto richiesto dentro una
    tupla ``(metadata, payload)`` e aggiunge uno o più elementi ``bytes`` di
    chiusura.  Le versioni precedenti dell'applicazione consideravano soltanto
    gli elementi ``bytes`` di primo livello e quindi, con server come quello
    usato dall'utente, finivano per leggere soltanto ``b')'``. Il risultato era
    una lista con i soli UID e un'anteprima vuota.

    La funzione è volutamente tollerante perché i server IMAP possono
    restituire forme leggermente diverse per HEADER e BODY.PEEK[].
    """
    payloads: list[bytes] = []
    for chunk in chunks or []:
        if isinstance(chunk, tuple):
            # Il secondo elemento è il literal richiesto da FETCH.
            if len(chunk) >= 2 and isinstance(chunk[1], (bytes, bytearray)):
                payloads.append(bytes(chunk[1]))
        elif isinstance(chunk, (bytes, bytearray)):
            raw = bytes(chunk)
            # I soli token di protocollo, ad esempio b')', non fanno parte
            # del messaggio. Conserviamo invece bytes "normali" per
            # compatibilità con server che restituiscono direttamente il
            # literal.
            if raw.strip() not in {b")", b"("}:
                payloads.append(raw)
    return b"".join(payloads)


class ImapService:
    def __init__(self, config: ConfigManager):
        self.config = config

    def _connect(self) -> imaplib.IMAP4:
        s = self.config.imap
        password = self.config.get_password()
        if not s.host or not s.username:
            raise RuntimeError("Configurazione IMAP incompleta")
        if password is None:
            raise RuntimeError("Password IMAP non presente nel file credenziali cifrato")
        if s.ssl:
            ctx = ssl.create_default_context()
            conn: imaplib.IMAP4 = imaplib.IMAP4_SSL(s.host, s.port, ssl_context=ctx)
        else:
            conn = imaplib.IMAP4(s.host, s.port)
        conn.login(s.username, password)
        return conn

    @staticmethod
    def _select(conn: imaplib.IMAP4, folder: str) -> int:
        status, data = conn.select(f'"{folder}"', readonly=True)
        if status != "OK":
            raise RuntimeError(f"Impossibile aprire la cartella IMAP: {folder}")
        uidvalidity = 0
        status, resp = conn.response("UIDVALIDITY")
        if status == "UIDVALIDITY" and resp:
            try:
                uidvalidity = int(resp[0])
            except Exception:
                pass
        if not uidvalidity:
            # Alcuni server lo espongono solo tramite STATUS.
            st, values = conn.status(f'"{folder}"', "(UIDVALIDITY)")
            if st == "OK" and values and values[0]:
                m = re.search(rb"UIDVALIDITY\s+(\d+)", values[0])
                if m:
                    uidvalidity = int(m.group(1))
        return uidvalidity

    def test_connection(self, folder: str | None = None) -> tuple[bool, str]:
        conn = None
        try:
            conn = self._connect()
            uidv = self._select(conn, folder or self.config.imap.folder)
            return True, f"Connessione riuscita. UIDVALIDITY={uidv or 'n/d'}"
        except Exception as e:
            return False, str(e)
        finally:
            if conn:
                try:
                    conn.logout()
                except Exception:
                    pass

    def search_progressive(
        self,
        folder: str,
        subject_filter: str,
        sender_filter: str,
        stop_event: threading.Event,
        on_item: Callable[[MailHeader | MailMessage], None],
        on_progress: Callable[[int, int], None] | None = None,
        *,
        include_body: bool = False,
        order: str = "desc",
        since_date=None,
    ) -> None:
        """Ricerca progressiva con prefiltri eseguiti, quando possibile, dal server IMAP.

        ``order`` può essere ``desc`` (mail più recenti prima) oppure ``asc``.
        ``since_date`` accetta un ``datetime.date`` e viene tradotto nel criterio
        IMAP SINCE, così le caselle molto grandi non devono essere percorse per intero.
        """
        conn = self._connect()
        try:
            self._select(conn, folder)
            criteria: list[str] = []
            if since_date is not None:
                months = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
                criteria.extend(["SINCE", f'{since_date.day:02d}-{months[since_date.month-1]}-{since_date.year:04d}'])
            # SUBJECT/FROM riducono drasticamente i candidati sui server che li indicizzano.
            # Il filtro viene comunque ricontrollato localmente sotto, quindi non cambia
            # il risultato funzionale del programma.
            if subject_filter.strip():
                criteria.extend(["SUBJECT", f'"{subject_filter.strip().replace(chr(34), "").replace(chr(92), "")}"'])
            if sender_filter.strip():
                criteria.extend(["FROM", f'"{sender_filter.strip().replace(chr(34), "").replace(chr(92), "")}"'])
            if not criteria:
                criteria = ["ALL"]
            status, data = conn.uid("search", None, *criteria)
            if status != "OK":
                # Alcuni server sono più restrittivi sui criteri testuali: fallback
                # trasparente alla ricerca completa, mantenendo poi i filtri locali.
                status, data = conn.uid("search", None, "ALL")
            if status != "OK":
                raise RuntimeError("Ricerca IMAP non riuscita")
            uids = [int(x) for x in (data[0] or b"").split()]
            uids.sort(reverse=(str(order).lower() != "asc"))
            total = len(uids)
            sf = subject_filter.casefold().strip()
            ff = sender_filter.casefold().strip()
            for i, uid in enumerate(uids, 1):
                if stop_event.is_set():
                    break
                status, chunks = conn.uid(
                    "fetch", str(uid),
                    "(BODY.PEEK[HEADER.FIELDS (FROM TO SUBJECT DATE MESSAGE-ID)])"
                )
                if status != "OK":
                    continue
                raw = imap_fetch_payload(chunks)
                if not raw:
                    if on_progress:
                        on_progress(i, total)
                    continue
                msg = email.message_from_bytes(raw)
                header = MailHeader(
                    uid=uid,
                    sender=decode_mime(msg.get("From")),
                    subject=decode_mime(msg.get("Subject")),
                    date=decode_mime(msg.get("Date")),
                    message_id=decode_mime(msg.get("Message-ID")),
                    recipient=decode_mime(msg.get("To")),
                )
                if on_progress:
                    on_progress(i, total)
                if sf and sf not in header.subject.casefold():
                    continue
                if ff and ff not in header.sender.casefold():
                    continue

                if include_body:
                    status, full_chunks = conn.uid("fetch", str(uid), "(BODY.PEEK[])")
                    if status != "OK":
                        continue
                    full_raw = imap_fetch_payload(full_chunks)
                    if not full_raw:
                        continue
                    full_msg = email.message_from_bytes(full_raw)
                    raw_date = full_msg.get("Date", "")
                    on_item(MailMessage(
                        uid=uid,
                        sender=decode_mime(full_msg.get("From")),
                        subject=decode_mime(full_msg.get("Subject")),
                        date=decode_mime(raw_date),
                        raw_date=raw_date,
                        message_id=decode_mime(full_msg.get("Message-ID")),
                        recipient=decode_mime(full_msg.get("To")),
                        body=message_body(full_msg),
                    ))
                else:
                    on_item(header)
        finally:
            try:
                conn.logout()
            except Exception:
                pass

    def fetch_headers(
        self,
        uids: Iterable[int],
        folder: str,
        stop_event: threading.Event | None = None,
        on_progress: Callable[[int, int], None] | None = None,
    ) -> list[MailHeader]:
        """Legge gli header di una lista di UID con una sola connessione IMAP.

        È usato prima dell'estrazione per stabilire l'ordine cronologico reale
        senza dover scaricare in anticipo il corpo delle mail più recenti.
        """
        uid_list = sorted({int(uid) for uid in uids})
        conn = self._connect()
        out: list[MailHeader] = []
        try:
            self._select(conn, folder)
            for i, uid in enumerate(uid_list, 1):
                if stop_event is not None and stop_event.is_set():
                    break
                status, chunks = conn.uid(
                    "fetch", str(uid),
                    "(BODY.PEEK[HEADER.FIELDS (FROM TO SUBJECT DATE MESSAGE-ID)])",
                )
                if status == "OK":
                    raw = imap_fetch_payload(chunks)
                    if raw:
                        msg = email.message_from_bytes(raw)
                        out.append(MailHeader(
                            uid=uid,
                            sender=decode_mime(msg.get("From")),
                            subject=decode_mime(msg.get("Subject")),
                            date=decode_mime(msg.get("Date")),
                            message_id=decode_mime(msg.get("Message-ID")),
                            recipient=decode_mime(msg.get("To")),
                        ))
                if on_progress:
                    on_progress(i, len(uid_list))
            return out
        finally:
            try:
                conn.logout()
            except Exception:
                pass

    def fetch_message(self, uid: int, folder: str) -> MailMessage:
        conn = self._connect()
        try:
            self._select(conn, folder)
            status, chunks = conn.uid("fetch", str(uid), "(BODY.PEEK[])")
            if status != "OK":
                raise RuntimeError(f"Impossibile leggere UID {uid}")
            raw = imap_fetch_payload(chunks)
            if not raw:
                raise RuntimeError(
                    f"Il server IMAP non ha restituito il contenuto della mail UID {uid}."
                )
            msg = email.message_from_bytes(raw)
            raw_date = msg.get("Date", "")
            return MailMessage(
                uid=uid,
                sender=decode_mime(msg.get("From")),
                subject=decode_mime(msg.get("Subject")),
                date=decode_mime(raw_date),
                raw_date=raw_date,
                message_id=decode_mime(msg.get("Message-ID")),
                recipient=decode_mime(msg.get("To")),
                body=message_body(msg),
            )
        finally:
            try:
                conn.logout()
            except Exception:
                pass

    def fetch_message_images(self, uid: int, folder: str) -> list[dict]:
        """Restituisce le immagini inline/allegate di una mail per scegliere l'icona profilo."""
        conn = self._connect()
        try:
            self._select(conn, folder)
            status, chunks = conn.uid("fetch", str(uid), "(BODY.PEEK[])")
            if status != "OK":
                raise RuntimeError(f"Impossibile leggere UID {uid}")
            raw = imap_fetch_payload(chunks)
            if not raw:
                return []
            msg = email.message_from_bytes(raw)
            out = []
            n = 0
            for part in msg.walk():
                ctype = str(part.get_content_type() or "")
                if not ctype.startswith("image/"):
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                n += 1
                filename = decode_mime(part.get_filename()) or f"immagine_{n}.{ctype.split('/',1)[1].split('+',1)[0]}"
                out.append({"filename": filename, "content_type": ctype, "data": payload})
            return out
        finally:
            try:
                conn.logout()
            except Exception:
                pass

    def uids_since_date(self, folder: str, since_date) -> tuple[int, list[int]]:
        """Restituisce gli UID delle mail dalla data indicata usando SEARCH SINCE lato server."""
        conn = self._connect()
        try:
            uidvalidity = self._select(conn, folder)
            criterion = since_date.strftime("%d-%b-%Y")
            status, data = conn.uid("search", None, "SINCE", criterion)
            if status != "OK":
                raise RuntimeError("Ricerca storica IMAP per data non riuscita")
            uids = [int(x) for x in (data[0] or b"").split()]
            return uidvalidity, sorted(set(uids))
        finally:
            try:
                conn.logout()
            except Exception:
                pass

    def incremental_uids(self, folder: str, after_uid: int) -> tuple[int, list[int]]:
        conn = self._connect()
        try:
            uidvalidity = self._select(conn, folder)
            criterion = f"UID {max(1, after_uid + 1)}:*" if after_uid else "ALL"
            status, data = conn.uid("search", None, criterion)
            if status != "OK":
                raise RuntimeError("Ricerca incrementale IMAP non riuscita")
            uids = [int(x) for x in (data[0] or b"").split()]
            # Alcuni server possono includere after_uid; filtriamo esplicitamente.
            uids = sorted(u for u in uids if u > after_uid)
            return uidvalidity, uids
        finally:
            try:
                conn.logout()
            except Exception:
                pass
