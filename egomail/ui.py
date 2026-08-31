from __future__ import annotations

from . import __version__

import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from email.utils import parsedate_to_datetime
from datetime import datetime
import webbrowser
import tempfile
import shutil
import zipfile
import tkinter as tk

try:
    from PIL import Image, ImageTk
except Exception:  # Pillow è consigliato ma la GUI resta avviabile senza.
    Image = ImageTk = None
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from .config import ConfigManager, ImapSettings
from .debug_bundle import build_debug_bundle
from .exporter import (
    append_extraction_log, apply_detailed_extractions, final_formula_aliases,
    expand_summary_formula, load_existing_records, merge_records, write_excel,
    read_excel_home_summary_values, summary_formula_display,
)
from .extraction import (
    compute_fields,
    extract_message,
    extract_mail_detailed,
    extract_message_all,
    extract_rule,
    extract_rule_trace,
    humanize_extraction_rule,
    matching_mail_types,
    matches_mail_type, header_prefilter_matches_profile,
    reconcile,
)
from .mail import ImapService, MailHeader, MailMessage
from .profile import (
    FIELD_TYPES,
    EXTRACTION_SOURCES,
    MATCH_MODES,
    MATCH_SOURCES,
    ProfileManager,
    copy_extraction_rules,
    delete_mail_type,
    ensure_field,
    field_generation_sources,
    get_field,
    normalize_field_type,
    merge_field_into,
    rename_mail_type,
    slugify,
)
from .rule_builder import infer_rule_from_selection, propose_field_name
from .regex_utils import expand_regex, set_regex_abbreviations, normalize_regex_abbreviations
from .edition import IS_PRO, DISPLAY_NAME
from .support_mail import SmtpSettings, infer_smtp_from_imap, send_support_email
from .pipeline import atomic_write_json, read_json, header_to_dict, dict_to_header, event_to_dict, dict_to_event, safe_unlink, transactional_paths, prepare_transaction_source
from .profile_transfer import export_profile_package, import_profile_package, merge_regex_abbreviations, inspect_profile_package, export_profile_backup, inspect_profile_backup, export_profiles_backup, inspect_profiles_backup, read_profile_from_backup, ABBREVIATIONS_FILENAME, BACKUP_STATE_FILENAME


RULE_COLORS = (
    "#FFF59D", "#C8E6C9", "#BBDEFB", "#F8BBD0",
    "#FFE0B2", "#D1C4E9", "#B2EBF2", "#DCEDC8",
)

APP_TITLE = DISPLAY_NAME


class ExtractionRetryCancelled(RuntimeError):
    """L'utente ha scelto di non continuare i tentativi IMAP durante l'estrazione."""

APP_VERSION = __version__
APP_DATE = "31/08/2026"
APP_AUTHOR = ""
APP_WEBSITE = "https://www.domoticachepassione.it/wp/acquista-egomailextractor-pro/"


def center_window(win: tk.Misc, parent: tk.Misc | None = None) -> None:
    """Centra una finestra rispetto al parent visibile, altrimenti allo schermo."""
    try:
        win.update_idletasks()
        width = win.winfo_width() or win.winfo_reqwidth()
        height = win.winfo_height() or win.winfo_reqheight()
        if parent is not None and parent.winfo_exists() and parent.winfo_viewable():
            px = parent.winfo_rootx(); py = parent.winfo_rooty()
            pw = parent.winfo_width(); ph = parent.winfo_height()
            x = px + max(0, (pw - width) // 2)
            y = py + max(0, (ph - height) // 2)
        else:
            x = max(0, (win.winfo_screenwidth() - width) // 2)
            y = max(0, (win.winfo_screenheight() - height) // 2)
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass


def center_dialog_later(win: tk.Misc, parent: tk.Misc | None = None) -> None:
    try:
        win.after_idle(lambda: center_window(win, parent))
    except Exception:
        pass


def matches_active_type_filters(row_matches: set[str], active_filters: set[str]) -> bool:
    """True se la riga deve essere visibile con il filtro tipi mail corrente.

    Nessun filtro = mostra tutto. Con più tipi selezionati si usa una logica OR:
    basta che la mail corrisponda ad almeno uno dei tipi selezionati.
    """
    return not active_filters or bool(set(row_matches) & set(active_filters))


def safe_filename_component(value: str, fallback: str = "profilo") -> str:
    text = re.sub(r'[<>:"/\\|?*]+', "_", str(value or "").strip())
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or fallback


def show_splash(parent: tk.Misc, duration_ms: int = 1700, *, closable: bool = False) -> None:
    """Mostra lo splash/informazioni di EgoMailExtractor.

    Lo stesso contenuto viene usato sia all'avvio sia da Help → Informazioni su.
    All'avvio la finestra si chiude automaticamente; in modalità ``closable``
    resta aperta finché l'utente la chiude. Il sito è sempre cliccabile.
    """
    splash = tk.Toplevel(parent)
    if closable:
        splash.title(f"Informazioni su {APP_TITLE}")
        splash.resizable(False, False)
        try:
            splash.transient(parent)
        except Exception:
            pass
    else:
        splash.overrideredirect(True)
        try:
            splash.attributes("-topmost", True)
        except Exception:
            pass

    outer = tk.Frame(splash, bg="#ffffff", bd=1, relief="solid")
    outer.pack(fill="both", expand=True)

    # Titolo prima, simbolo subito dopo come richiesto.
    tk.Label(
        outer, text=APP_TITLE, bg="#ffffff", fg="#0F4761",
        font=("Segoe UI", 24, "bold"),
    ).pack(padx=42, pady=(24, 8))

    icon_img = apply_app_icon(splash)
    if icon_img is not None:
        try:
            # 512 -> circa 128 px, sufficiente per lo splash senza dipendenze esterne.
            splash_img = icon_img.subsample(4, 4)
            splash._splash_icon_ref = splash_img
            tk.Label(outer, image=splash_img, bg="#ffffff", bd=0).pack(pady=(0, 8))
        except Exception:
            pass

    tk.Label(
        outer, text=f"Versione {APP_VERSION}", bg="#ffffff", fg="#333333",
        font=("Segoe UI", 11),
    ).pack(pady=(0, 4))
    tk.Label(
        outer, text=APP_DATE, bg="#ffffff", fg="#555555",
        font=("Segoe UI", 9),
    ).pack()

    website = tk.Label(
        outer,
        text=APP_WEBSITE,
        bg="#ffffff",
        fg="#0563C1",
        cursor="hand2",
        font=("Segoe UI", 9, "underline"),
    )
    website.pack(pady=(0, 26))
    website.bind("<Button-1>", lambda _event: webbrowser.open(APP_WEBSITE))

    splash.update_idletasks()
    w = splash.winfo_reqwidth()
    h = splash.winfo_reqheight()
    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")
    splash.lift()
    try:
        splash.focus_force()
    except Exception:
        pass

    if closable:
        splash.protocol("WM_DELETE_WINDOW", splash.destroy)
        splash.bind("<Escape>", lambda _e: splash.destroy())
    elif duration_ms and duration_ms > 0:
        splash.after(duration_ms, splash.destroy)
    splash.wait_window()


def _resource_path(*parts: str) -> Path:
    """Percorso di una risorsa sia da sorgenti sia da bundle PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)




def load_app_icon(size: int = 128) -> tk.PhotoImage | None:
    """Carica l'icona applicativa dalle risorse del pacchetto."""
    path = _resource_path("assets", "egomail_icon.png")
    if not path.exists():
        return None
    try:
        img = tk.PhotoImage(file=str(path))
        # Tk PhotoImage non supporta resize arbitrario senza Pillow; la PNG 512x512
        # viene lasciata intatta per iconphoto e ridotta con subsample nello splash.
        return img
    except Exception:
        return None


def apply_app_icon(window: tk.Misc) -> tk.PhotoImage | None:
    """Imposta l'icona EgoMailExtractor su una finestra Tk/Toplevel.

    Su Windows usiamo anche ``iconbitmap`` con il file ICO multi-risoluzione:
    in questo modo la stessa icona viene mostrata in barra applicazioni, Alt-Tab
    e finestre Tk anche nel bundle PyInstaller. ``iconphoto`` resta attivo come
    fallback multipiattaforma e per le finestre secondarie.
    """
    img = load_app_icon()
    if img is not None:
        try:
            window.iconphoto(True, img)
            setattr(window, "_egomail_icon_ref", img)
        except Exception:
            pass

    if sys.platform.startswith("win"):
        ico = _resource_path("assets", "egomail_icon.ico")
        if ico.exists():
            try:
                window.iconbitmap(default=str(ico))
                setattr(window, "_egomail_ico_path", str(ico))
            except Exception:
                pass
    return img

def open_extraction_help(parent: tk.Misc | None = None) -> None:
    path = _resource_path("help", "regole_estrazione.html")
    if not path.exists():
        if parent:
            messagebox.showerror("Help", f"File di help non trovato:\n{path}", parent=parent)
        return
    try:
        webbrowser.open(path.resolve().as_uri())
    except Exception as exc:
        if parent:
            messagebox.showerror("Help", str(exc), parent=parent)


class HoverTip:
    """Tooltip leggero usato sui tag del contenuto della mail."""
    def __init__(self, owner: tk.Misc):
        self.owner = owner
        self.tip: tk.Toplevel | None = None

    def show(self, text: str, x: int, y: int):
        self.hide()
        tip = tk.Toplevel(self.owner)
        tip.wm_overrideredirect(True)
        try:
            tip.attributes("-topmost", True)
        except Exception:
            pass
        label = tk.Label(
            tip, text=text, justify="left", relief="solid", borderwidth=1,
            background="#ffffe0", padx=6, pady=4, wraplength=420,
        )
        label.pack()
        tip.wm_geometry(f"+{x + 14}+{y + 16}")
        self.tip = tip

    def hide(self):
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


class MasterPasswordDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, creating: bool = False):
        super().__init__(parent)
        self.creating = creating
        self.result: str | None = None
        self.title("Crea password master" if creating else "Sblocca EgoMailExtractor")
        self.resizable(False, False)
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        msg = (
            "Crea la password master che proteggerà le credenziali email."
            if creating else
            "Inserisci la password master per decriptare le credenziali email."
        )
        ttk.Label(frm, text=msg, wraplength=430, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )
        ttk.Label(frm, text="Password master").grid(row=1, column=0, sticky="w", pady=4)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(frm, textvariable=self.password_var, width=38, show="•")
        self.password_entry.grid(row=1, column=1, sticky="ew", pady=4)
        self.confirm_var = tk.StringVar()
        if creating:
            ttk.Label(frm, text="Conferma password").grid(row=2, column=0, sticky="w", pady=4)
            ttk.Entry(frm, textvariable=self.confirm_var, width=38, show="•").grid(
                row=2, column=1, sticky="ew", pady=4
            )
        note_row = 3 if creating else 2
        ttk.Label(
            frm,
            text="La password master non viene salvata. Se viene dimenticata, le credenziali cifrate non possono essere recuperate.",
            wraplength=430,
            foreground="#666666",
            justify="left",
        ).grid(row=note_row, column=0, columnspan=2, sticky="w", pady=(10, 8))
        btns = ttk.Frame(frm)
        btns.grid(row=note_row + 1, column=0, columnspan=2, sticky="e", pady=(6, 0))
        ttk.Button(btns, text="Annulla", command=self.cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Crea" if creating else "Sblocca", command=self.accept).pack(side="right")
        frm.columnconfigure(1, weight=1)
        self.bind("<Return>", lambda _e: self.accept())
        self.bind("<Escape>", lambda _e: self.cancel())
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        # Non rendere il dialog transient rispetto a una root ritirata: su alcuni
        # sistemi Windows puo' non venire portato correttamente in primo piano.
        try:
            if parent.winfo_viewable():
                self.transient(parent)
        except Exception:
            pass
        self.grab_set()
        center_dialog_later(self, parent)
        self.after(50, self._bring_to_front)

    def _bring_to_front(self):
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.password_entry.focus_set()
            self.after(350, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def accept(self):
        password = self.password_var.get()
        if not password:
            messagebox.showwarning("Password master", "Inserisci una password master.", parent=self)
            return
        if self.creating and password != self.confirm_var.get():
            messagebox.showerror("Password master", "Le due password non coincidono.", parent=self)
            return
        self.result = password
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class ImapSettingsDialog(tk.Toplevel):
    """Editor di un singolo profilo IMAP cifrato."""

    def __init__(
        self,
        parent: tk.Misc,
        config: ConfigManager,
        profile_id: str | None = None,
        *,
        creating: bool = False,
    ):
        super().__init__(parent)
        self.config_mgr = config
        self.profile_id = None if creating else (profile_id or config.active_imap_profile_id or None)
        self.result: str | None = None
        current = config.get_imap_profile(self.profile_id) if self.profile_id else None
        if current:
            current_name, s, pwd = current
        else:
            current_name, s, pwd = "", ImapSettings(), ""
        self.title("Nuova email" if not current else f"Email — {current_name}")
        self.resizable(False, False)
        self.vars = {
            "protocol": tk.StringVar(value="IMAP"),
            "name": tk.StringVar(value=current_name),
            "host": tk.StringVar(value=s.host),
            "port": tk.StringVar(value=str(s.port)),
            "username": tk.StringVar(value=s.username),
            "folder": tk.StringVar(value=s.folder),
            "ssl": tk.BooleanVar(value=s.ssl),
            "password": tk.StringVar(value=pwd),
            "default": tk.BooleanVar(value=(
                self.profile_id == config.default_imap_profile_id if self.profile_id else config.imap_profile_count == 0
            )),
        }
        frm = ttk.Frame(self, padding=12)
        frm.grid(sticky="nsew")
        ttk.Label(frm, text="Protocollo").grid(row=0, column=0, sticky="w", pady=4)
        protocol_combo = ttk.Combobox(frm, textvariable=self.vars["protocol"], values=("IMAP",), state="readonly", width=41)
        protocol_combo.grid(row=0, column=1, sticky="ew", pady=4)
        labels = [
            ("Nome profilo", "name"),
            ("Server email", "host"),
            ("Porta", "port"),
            ("Utente", "username"),
            ("Password", "password"),
            ("Cartella predefinita", "folder"),
        ]
        for r, (lab, key) in enumerate(labels, start=1):
            ttk.Label(frm, text=lab).grid(row=r, column=0, sticky="w", pady=4)
            e = ttk.Entry(frm, textvariable=self.vars[key], width=44, show="*" if key == "password" else "")
            e.grid(row=r, column=1, sticky="ew", pady=4)
        ttk.Checkbutton(frm, text="Usa SSL/TLS", variable=self.vars["ssl"]).grid(row=7, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            frm, text="Email predefinita", variable=self.vars["default"]
        ).grid(row=8, column=1, sticky="w", pady=4)
        ttk.Label(
            frm,
            text="Tutti i dati di questo profilo, compresa la password, sono salvati nel file cifrato.",
            foreground="#666666",
            wraplength=420,
        ).grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))
        btns = ttk.Frame(frm)
        btns.grid(row=10, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Prova connessione", command=self.test).pack(side="left", padx=4)
        ttk.Button(btns, text="Salva", command=self.save).pack(side="left", padx=4)
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="left", padx=4)
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def _settings(self) -> ImapSettings:
        try:
            port = int(self.vars["port"].get())
        except ValueError:
            raise ValueError("La porta deve essere numerica")
        name = self.vars["name"].get().strip()
        if not name:
            raise ValueError("Inserisci un nome per l'email")
        return ImapSettings(
            host=self.vars["host"].get().strip(), port=port,
            ssl=bool(self.vars["ssl"].get()), username=self.vars["username"].get().strip(),
            folder=self.vars["folder"].get().strip() or "INBOX"
        )

    def _apply(self) -> str:
        settings = self._settings()
        name = self.vars["name"].get().strip()
        password = self.vars["password"].get()
        if self.profile_id:
            self.config_mgr.update_imap_profile(
                self.profile_id, name=name, settings=settings, password=password
            )
        else:
            self.profile_id = self.config_mgr.create_imap_profile(
                name, settings, password, make_default=bool(self.vars["default"].get()), make_active=True
            )
        self.config_mgr.set_active_imap_profile(self.profile_id)
        if self.vars["default"].get():
            self.config_mgr.set_default_imap_profile(self.profile_id)
        return self.profile_id

    def test(self):
        try:
            self._apply()
            ok, msg = ImapService(self.config_mgr).test_connection(self.vars["folder"].get())
            messagebox.showinfo("Connessione email" if ok else "Errore email", msg, parent=self)
        except Exception as e:
            messagebox.showerror("Errore", str(e), parent=self)

    def save(self):
        try:
            self.result = self._apply()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Errore", str(e), parent=self)


class ImapProfileChooser(tk.Toplevel):
    """Scelta del profilo IMAP quando sono disponibili più account."""

    def __init__(self, parent: tk.Misc, config: ConfigManager, title: str = "Scegli email"):
        super().__init__(parent)
        self.config_mgr = config
        self.result: str | None = None
        self.title(title)
        self.geometry("680x390")
        self.minsize(560, 320)
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(
            frm, text="Email disponibili", font=("TkDefaultFont", 11, "bold")
        ).pack(anchor="w")
        ttk.Label(
            frm,
            text="Seleziona l'account da usare. Il profilo predefinito è preselezionato.",
            foreground="#666666",
        ).pack(anchor="w", pady=(2, 8))
        self.tree = ttk.Treeview(
            frm, columns=("name", "user", "server", "default"), show="headings", selectmode="browse"
        )
        for c, label, width in [
            ("name", "Profilo", 180), ("user", "Utente", 200),
            ("server", "Server", 180), ("default", "Default", 70),
        ]:
            self.tree.heading(c, text=label)
            self.tree.column(c, width=width, stretch=True)
        self.tree.pack(fill="both", expand=True)
        self.profiles = config.list_imap_profiles()
        default_iid = None
        for item in self.profiles:
            iid = str(item["id"])
            self.tree.insert("", "end", iid=iid, values=(
                item.get("name", ""), item.get("username", ""), item.get("host", ""),
                "Sì" if item.get("is_default") else "",
            ))
            if item.get("is_default"):
                default_iid = iid
        choose_iid = default_iid or (str(self.profiles[0]["id"]) if self.profiles else None)
        if choose_iid:
            self.tree.selection_set(choose_iid)
            self.tree.focus(choose_iid)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Apri", command=self.open_selected).pack(side="left")
        ttk.Button(btns, text="Annulla", command=self.cancel).pack(side="right")
        self.tree.bind("<Double-1>", lambda _e: self.open_selected())
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        try:
            if parent.winfo_viewable():
                self.transient(parent)
        except Exception:
            pass
        self.grab_set()
        center_dialog_later(self, parent)
        self.after(50, self._bring_to_front)

    def _bring_to_front(self):
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.after(350, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def open_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.result = str(sel[0])
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class ImapProfilesDialog(tk.Toplevel):
    """Gestione completa dei profili IMAP cifrati."""

    def __init__(self, parent: tk.Misc, config: ConfigManager, on_active_changed=None):
        super().__init__(parent)
        self.config_mgr = config
        self.on_active_changed = on_active_changed
        self.title("Email")
        self.geometry("900x470")
        self.minsize(720, 380)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        holder = ttk.Frame(frm)
        holder.pack(fill="both", expand=True)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        cols = ("name", "user", "server", "folder", "default", "active")
        self.tree = ttk.Treeview(holder, columns=cols, show="headings", selectmode="browse")
        for c, label, width in [
            ("name", "Profilo", 170), ("user", "Utente", 200), ("server", "Server", 180),
            ("folder", "Cartella", 110), ("default", "Default", 70), ("active", "Attivo", 60),
        ]:
            self.tree.heading(c, text=label)
            self.tree.column(c, width=width, stretch=True)
        y = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=y.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Nuovo…", command=self.add_profile).pack(side="left")
        ttk.Button(btns, text="Modifica…", command=self.edit_profile).pack(side="left", padx=5)
        ttk.Button(btns, text="Elimina", command=self.delete_profile).pack(side="left")
        ttk.Button(btns, text="Imposta predefinito", command=self.set_default).pack(side="left", padx=(14, 5))
        ttk.Button(btns, text="Usa questo profilo", command=self.activate).pack(side="left")
        ttk.Button(btns, text="Chiudi", command=self.destroy).pack(side="right")
        self.tree.bind("<Double-1>", lambda _e: self.edit_profile())
        self.refresh()
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def selected_id(self) -> str | None:
        sel = self.tree.selection()
        return str(sel[0]) if sel else None

    def refresh(self, select_id: str | None = None):
        old = select_id or self.selected_id()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for p in self.config_mgr.list_imap_profiles():
            iid = str(p["id"])
            self.tree.insert("", "end", iid=iid, values=(
                p.get("name", ""), p.get("username", ""), p.get("host", ""), p.get("folder", "INBOX"),
                "Sì" if p.get("is_default") else "", "Sì" if p.get("is_active") else "",
            ))
        if old and self.tree.exists(old):
            self.tree.selection_set(old)
            self.tree.focus(old)

    def add_profile(self):
        dlg = ImapSettingsDialog(self, self.config_mgr, creating=True)
        self.wait_window(dlg)
        if dlg.result:
            self.refresh(dlg.result)
            if self.on_active_changed:
                self.on_active_changed()

    def edit_profile(self):
        pid = self.selected_id()
        if not pid:
            return
        dlg = ImapSettingsDialog(self, self.config_mgr, pid)
        self.wait_window(dlg)
        if dlg.result:
            self.refresh(pid)
            if self.on_active_changed:
                self.on_active_changed()

    def delete_profile(self):
        pid = self.selected_id()
        if not pid:
            return
        info = next((x for x in self.config_mgr.list_imap_profiles() if x.get("id") == pid), {})
        if not messagebox.askyesno(
            "Elimina email",
            f"Eliminare l'email «{info.get('name', '')}» e le sue credenziali cifrate?",
            parent=self,
        ):
            return
        self.config_mgr.delete_imap_profile(pid)
        self.refresh()
        if self.on_active_changed:
            self.on_active_changed()

    def set_default(self):
        pid = self.selected_id()
        if not pid:
            return
        self.config_mgr.set_default_imap_profile(pid)
        self.refresh(pid)

    def activate(self):
        pid = self.selected_id()
        if not pid:
            return
        self.config_mgr.set_active_imap_profile(pid)
        self.refresh(pid)
        if self.on_active_changed:
            self.on_active_changed()


class ProfileChooser(tk.Toplevel):
    def __init__(self, parent: tk.Misc, manager: ProfileManager, *, paths: list[Path] | None = None, associated_only: bool = False, config: ConfigManager | None = None):
        super().__init__(parent)
        self.title("Cambia profilo")
        self.geometry("700x440")
        self.manager = manager
        self.config_mgr = config
        self.associated_only = associated_only
        self.result: Path | None = None
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        title = "Profili associati alla email" if associated_only else "Profili disponibili"
        ttk.Label(frm, text=title, font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        if associated_only:
            ttk.Label(
                frm,
                text="Se non si vede il profilo di estrazione, controllare i profili di estrazione associati alla email",
                foreground="#666666", wraplength=650,
            ).pack(anchor="w", pady=(3, 6))
        self.list = tk.Listbox(frm)
        self.list.pack(fill="both", expand=True, pady=8)
        self.paths = list(paths) if paths is not None else manager.list_profiles()
        for p in self.paths:
            try:
                prof = manager.load(p)
                self.list.insert("end", f"{prof.get('name', p.stem)}   [{p.name}]")
            except Exception:
                self.list.insert("end", p.name)
        if self.paths:
            self.list.selection_set(0)
        btns = ttk.Frame(frm)
        btns.pack(fill="x")
        ttk.Button(btns, text="Apri", command=self.open_selected).pack(side="left")
        ttk.Button(btns, text="Nuovo profilo", command=self.new_profile).pack(side="left", padx=6)
        ttk.Button(btns, text="Esci", command=self.cancel).pack(side="right")
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        try:
            if parent.winfo_viewable():
                self.transient(parent)
        except Exception:
            pass
        self.grab_set()
        center_dialog_later(self, parent)
        self.after(50, self._bring_to_front)

    def _bring_to_front(self):
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.after(350, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def open_selected(self):
        sel = self.list.curselection()
        if not sel:
            return
        self.result = self.paths[sel[0]]
        self.destroy()

    def new_profile(self):
        name = simpledialog.askstring("Nuovo profilo", "Nome del profilo:", parent=self)
        if not name:
            return
        self.result = self.manager.create(name)
        if self.associated_only and self.config_mgr is not None and self.config_mgr.active_imap_profile_id:
            current = self.config_mgr.associated_profile_names()
            if current is None:
                current = [p.name for p in self.manager.list_profiles()]
            if self.result.name not in current:
                current.append(self.result.name)
            self.config_mgr.set_associated_profile_names(current)
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class EmailExtractionAssociationsDialog(tk.Toplevel):
    """Associa i profili di estrazione alla email attiva."""

    def __init__(self, parent: tk.Misc, config: ConfigManager, manager: ProfileManager):
        super().__init__(parent)
        self.config_mgr = config
        self.manager = manager
        self.title("Profili di estrazione associati all'email")
        self.geometry("760x520")
        self.minsize(620, 420)
        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        email_name = config.active_imap_profile_name or "<nessuna email attiva>"
        ttk.Label(frm, text=f"Email attiva: {email_name}", font=("TkDefaultFont", 11, "bold")).pack(anchor="w")
        ttk.Label(frm, text="Metti la spunta sui profili di estrazione che vuoi rendere disponibili per questa email.", foreground="#666666").pack(anchor="w", pady=(2, 8))
        self.paths = manager.list_profiles()
        explicit = config.associated_profile_names()
        selected = {p.name for p in self.paths} if explicit is None else set(explicit)
        holder = ttk.Frame(frm); holder.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(holder, columns=("assoc", "name", "file"), show="headings", selectmode="browse")
        for col, lab, width in [("assoc","Associato",90),("name","Profilo di estrazione",310),("file","File",240)]:
            self.tree.heading(col, text=lab); self.tree.column(col, width=width, stretch=(col != "assoc"), anchor="center" if col=="assoc" else "w")
        sy=ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=sy.set)
        self.tree.pack(side="left", fill="both", expand=True); sy.pack(side="right", fill="y")
        self._selected = set(selected)
        for p in self.paths:
            try: name = str(manager.load(p).get("name", p.stem) or p.stem)
            except Exception: name = p.stem
            self.tree.insert("", "end", iid=p.name, values=("✓" if p.name in self._selected else "", name, p.name))
        self.tree.bind("<Double-1>", self._toggle_event); self.tree.bind("<space>", self._toggle_event)
        btn=ttk.Frame(frm); btn.pack(fill="x", pady=(8,0))
        ttk.Button(btn, text="Seleziona tutti", command=lambda:self._set_all(True)).pack(side="left")
        ttk.Button(btn, text="Deseleziona tutti", command=lambda:self._set_all(False)).pack(side="left", padx=6)
        ttk.Button(btn, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btn, text="Salva", command=self._save).pack(side="right", padx=6)
        self.transient(parent); self.grab_set(); center_dialog_later(self, parent)

    def _toggle_event(self, _event=None):
        sel=self.tree.selection()
        if not sel: return
        iid=str(sel[0])
        if iid in self._selected: self._selected.remove(iid)
        else: self._selected.add(iid)
        vals=list(self.tree.item(iid, "values")); vals[0]="✓" if iid in self._selected else ""; self.tree.item(iid, values=vals)

    def _set_all(self, value: bool):
        self._selected={p.name for p in self.paths} if value else set()
        for iid in self.tree.get_children():
            vals=list(self.tree.item(iid,"values")); vals[0]="✓" if iid in self._selected else ""; self.tree.item(iid, values=vals)

    def _save(self):
        self.config_mgr.set_associated_profile_names(sorted(self._selected))
        self.destroy()


MATCH_SOURCE_LABELS = {
    "subject": "Oggetto",
    "sender": "Mittente",
    "body": "Corpo",
}
MATCH_SOURCE_FROM_LABEL = {v: k for k, v in MATCH_SOURCE_LABELS.items()}
MATCH_MODE_LABELS = {
    "contains": "Contiene",
    "regex": "Espressione regolare",
}
MATCH_MODE_FROM_LABEL = {v: k for k, v in MATCH_MODE_LABELS.items()}


class MatchRuleEditorDialog(tk.Toplevel):
    """Editor di una singola regola usata per selezionare un tipo mail."""

    def __init__(self, parent: tk.Misc, rule: dict | None = None, title: str = "Regola di selezione mail"):
        super().__init__(parent)
        rule = dict(rule or {})
        self.result: dict | None = None
        self.title(title)
        self.resizable(False, False)
        source = str(rule.get("source", "subject"))
        mode = str(rule.get("mode", "contains"))
        self.name_var = tk.StringVar(value=str(rule.get("name", "")))
        self.source_var = tk.StringVar(value=MATCH_SOURCE_LABELS.get(source, "Oggetto"))
        self.mode_var = tk.StringVar(value=MATCH_MODE_LABELS.get(mode, "Contiene"))
        self.trigger_var = tk.StringVar(value=str(rule.get("trigger", "")))

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text="Nome regola").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.name_var, width=52).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(frm, text="Dove cercare").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(
            frm, textvariable=self.source_var,
            values=list(MATCH_SOURCE_FROM_LABEL), state="readonly", width=30
        ).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(frm, text="Tipo di match").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(
            frm, textvariable=self.mode_var,
            values=list(MATCH_MODE_FROM_LABEL), state="readonly", width=30
        ).grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Label(frm, text="Trigger / match").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.trigger_var, width=52).grid(row=3, column=1, sticky="ew", pady=5)
        ttk.Label(
            frm,
            text=(
                "La regola seleziona la mail solo se il trigger è presente. "
                "Per un tipo mail, tutte le regole con trigger compilato devono essere vere. "
                "Se scegli Espressione regolare, il testo inserito viene interpretato come regex."
            ),
            wraplength=520, justify="left", foreground="#555555"
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 8))
        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Salva", command=self.save).pack(side="right")
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def save(self):
        trigger = self.trigger_var.get().strip()
        source = MATCH_SOURCE_FROM_LABEL.get(self.source_var.get(), "subject")
        mode = MATCH_MODE_FROM_LABEL.get(self.mode_var.get(), "contains")
        if mode == "regex" and trigger:
            try:
                re.compile(expand_regex(trigger))
            except re.error as exc:
                messagebox.showerror("Regex non valida", str(exc), parent=self)
                return
        name = self.name_var.get().strip() or f"match_{source}"
        self.result = {
            "name": slugify(name, f"match_{source}"),
            "source": source,
            "mode": mode,
            "trigger": trigger,
        }
        self.destroy()


class MatchRulesDialog(tk.Toplevel):
    """Gestione delle regole AND che identificano un tipo mail."""

    def __init__(self, parent: tk.Misc, mail_type: dict, on_change=None):
        super().__init__(parent)
        self.mail_type = mail_type
        self.on_change = on_change
        self.title(f"Selezione mail — {mail_type.get('name', '')}")
        self.geometry("900x480")
        self.minsize(760, 400)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(
            frm,
            text=(
                "Queste regole servono SOLO a riconoscere il tipo di mail. "
                "Tutti i trigger compilati devono corrispondere (AND) prima di applicare le regole di estrazione."
            ),
            wraplength=850, justify="left"
        ).pack(anchor="w", pady=(0, 8))
        cols = ("name", "source", "mode", "trigger")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for col, label, width in (
            ("name", "Regola", 180),
            ("source", "Match su", 110),
            ("mode", "Modalità", 150),
            ("trigger", "Trigger / valore da cercare", 420),
        ):
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, stretch=True)
        self.tree.pack(fill="both", expand=True)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Nuova regola", command=self.add_rule).pack(side="left")
        ttk.Button(btns, text="Modifica", command=self.edit_rule).pack(side="left", padx=5)
        ttk.Button(btns, text="Elimina", command=self.delete_rule).pack(side="left", padx=5)
        ttk.Button(btns, text="Chiudi", command=self.destroy).pack(side="right")
        self.tree.bind("<Double-1>", lambda _e: self.edit_rule())
        self.refresh()
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def _changed(self):
        if callable(self.on_change):
            self.on_change()

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, rule in enumerate(self.mail_type.setdefault("match_rules", [])):
            self.tree.insert("", "end", iid=str(i), values=(
                rule.get("name", ""),
                MATCH_SOURCE_LABELS.get(rule.get("source", "body"), rule.get("source", "body")),
                MATCH_MODE_LABELS.get(rule.get("mode", "contains"), rule.get("mode", "contains")),
                rule.get("trigger", ""),
            ))

    def add_rule(self):
        dlg = MatchRuleEditorDialog(self, title="Nuova regola di selezione")
        self.wait_window(dlg)
        if dlg.result:
            self.mail_type.setdefault("match_rules", []).append(dlg.result)
            self.refresh()
            self._changed()

    def edit_rule(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        rules = self.mail_type.setdefault("match_rules", [])
        dlg = MatchRuleEditorDialog(self, rules[i], title="Modifica regola di selezione")
        self.wait_window(dlg)
        if dlg.result:
            rules[i] = dlg.result
            self.refresh()
            self._changed()

    def delete_rule(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        rules = self.mail_type.setdefault("match_rules", [])
        rule = rules[i]
        if messagebox.askyesno(
            "Elimina regola", f"Eliminare la regola '{rule.get('name', '')}'?", parent=self
        ):
            del rules[i]
            self.refresh()
            self._changed()


FIELD_TYPE_LABELS = {
    "text": "Testo",
    "date": "Data",
    "number": "Numero",
}
FIELD_LABEL_TO_TYPE = {v: k for k, v in FIELD_TYPE_LABELS.items()}
FINAL_FORMULA_LABELS = {
    "": "Nessuna",
    "sum": "Somma",
    "count": "Conta",
}
FINAL_FORMULA_FROM_LABEL = {v: k for k, v in FINAL_FORMULA_LABELS.items()}
UPDATE_POLICY_LABELS = {
    "replace": "Sostituisci",
    "keep": "Non sostituire",
    "sum": "Somma",
}
UPDATE_POLICY_FROM_LABEL = {v: k for k, v in UPDATE_POLICY_LABELS.items()}
COMPUTED_KIND_LABELS = {
    "days_between": "Differenza giorni tra due date",
}
COMPUTED_KIND_FROM_LABEL = {v: k for k, v in COMPUTED_KIND_LABELS.items()}
COMPUTED_PRIORITY_LABELS = {
    "rule_wins": "Non sovrascrivere valori impostati da una regola",
    "always": "Ricalcola sempre",
}
COMPUTED_PRIORITY_FROM_LABEL = {v: k for k, v in COMPUTED_PRIORITY_LABELS.items()}


class RuleEditorDialog(tk.Toplevel):
    """Editor completo di una regola di estrazione con anteprima live."""

    TRANSFORMS = (
        "text", "upper", "int", "money", "guest_count", "date_it_email_year"
    )
    STRATEGIES = ("regex", "after_trigger", "constant")
    SOURCES = EXTRACTION_SOURCES

    def __init__(
        self,
        parent: tk.Misc,
        profile: dict,
        rule: dict,
        title: str = "Modifica regola",
        current_mail: MailMessage | None = None,
    ):
        super().__init__(parent)
        self.profile = profile
        self.original = dict(rule)
        self.current_mail = current_mail
        self.result: dict | None = None
        self.title(title)
        self.geometry("840x830")
        self.minsize(760, 740)

        field_name = str(rule.get("field", ""))
        field_spec = get_field(profile, field_name) or {"type": "text"}
        type_label = FIELD_TYPE_LABELS[normalize_field_type(field_spec.get("type"))]

        self.vars = {
            "name": tk.StringVar(value=str(rule.get("name", ""))),
            "field": tk.StringVar(value=field_name),
            "field_type": tk.StringVar(value=type_label),
            "source": tk.StringVar(value=str(rule.get("source", "body"))),
            "strategy": tk.StringVar(value=str(rule.get("strategy", "regex"))),
            "transform": tk.StringVar(value=str(rule.get("transform", "text"))),
            "pattern": tk.StringVar(value=str(rule.get("pattern", ""))),
            "trigger": tk.StringVar(value=str(rule.get("trigger", ""))),
            "value_pattern": tk.StringVar(value=str(rule.get("value_pattern", ""))),
            "occurrence": tk.StringVar(value=str(rule.get("occurrence", 1))),
            "group": tk.StringVar(value=str(rule.get("group", 1))),
            "nonempty_offset": tk.StringVar(value=str(rule.get("nonempty_offset", 1))),
            "value": tk.StringVar(value=str(rule.get("value", ""))),
            "currency": tk.BooleanVar(value=field_spec.get("format") == "currency_eur" or rule.get("transform") == "money"),
            "key": tk.BooleanVar(value=field_name in profile.get("reconcile_keys", [])),
            "existing_value_policy": tk.StringVar(
                value=UPDATE_POLICY_LABELS.get(
                    str(rule.get("existing_value_policy", "replace") or "replace"),
                    "Sostituisci",
                )
            ),
        }
        self.human_var = tk.StringVar()
        self.application_var = tk.StringVar()

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(outer, text="Nome regola").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.vars["name"]).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(outer, text="Campo da popolare").grid(row=row, column=0, sticky="w", pady=4)
        fields = [f.get("name", "") for f in profile.get("fields", []) if isinstance(f, dict)]
        ttk.Combobox(outer, textvariable=self.vars["field"], values=fields).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(outer, text="Tipo cella Excel").grid(row=row, column=0, sticky="w", pady=4)
        self.field_type_combo = ttk.Combobox(
            outer, textvariable=self.vars["field_type"],
            values=list(FIELD_LABEL_TO_TYPE), state="readonly"
        )
        self.field_type_combo.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        opts = ttk.Frame(outer)
        opts.grid(row=row, column=1, sticky="w", pady=(0, 7))
        ttk.Checkbutton(opts, text="Formato valuta €", variable=self.vars["currency"]).pack(side="left")
        ttk.Checkbutton(opts, text="Chiave di riconciliazione", variable=self.vars["key"]).pack(side="left", padx=18)
        row += 1

        ttk.Label(outer, text="Se la cella Excel ha già un valore").grid(row=row, column=0, sticky="w", pady=4)
        self.policy_combo = ttk.Combobox(
            outer, textvariable=self.vars["existing_value_policy"], state="readonly"
        )
        self.policy_combo.grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(
            outer,
            text="Somma è disponibile solo per campi di tipo Numero.",
            foreground="#666666",
        ).grid(row=row + 1, column=1, sticky="w", pady=(0, 4))
        row += 2
        self._sync_policy_choices()
        self.field_type_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_policy_choices())

        ttk.Separator(outer).grid(row=row, column=0, columnspan=2, sticky="ew", pady=8)
        row += 1

        ttk.Label(outer, text="Origine").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Combobox(outer, textvariable=self.vars["source"], values=self.SOURCES, state="readonly").grid(row=row, column=1, sticky="ew", pady=4)
        row += 1
        ttk.Label(outer, text="Strategia").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Combobox(outer, textvariable=self.vars["strategy"], values=self.STRATEGIES, state="readonly").grid(row=row, column=1, sticky="ew", pady=4)
        row += 1
        ttk.Label(outer, text="Conversione valore").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Combobox(outer, textvariable=self.vars["transform"], values=self.TRANSFORMS, state="readonly").grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Label(outer, text="Regex").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.vars["pattern"]).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1
        ttk.Label(outer, text="Trigger").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.vars["trigger"]).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1
        ttk.Label(outer, text="Regex sul valore (facoltativa)").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.vars["value_pattern"]).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        nums = ttk.Frame(outer)
        nums.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Label(nums, text="Occorrenza").pack(side="left")
        ttk.Entry(nums, textvariable=self.vars["occurrence"], width=7).pack(side="left", padx=(5, 18))
        ttk.Label(nums, text="Gruppo regex").pack(side="left")
        ttk.Entry(nums, textvariable=self.vars["group"], width=7).pack(side="left", padx=(5, 18))
        ttk.Label(nums, text="Offset righe non vuote").pack(side="left")
        ttk.Entry(nums, textvariable=self.vars["nonempty_offset"], width=7).pack(side="left", padx=5)
        row += 1

        ttk.Label(outer, text="Valore costante").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(outer, textvariable=self.vars["value"]).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        # Traduzione umana richiesta: viene aggiornata mentre si modificano i campi.
        human_box = ttk.LabelFrame(outer, text="Traduzione della regola in forma umana", padding=8)
        human_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(12, 5))
        ttk.Label(human_box, textvariable=self.human_var, wraplength=770, justify="left").pack(fill="x")
        row += 1

        application_box = ttk.LabelFrame(outer, text="Applicazione sulla mail corrente", padding=8)
        application_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        ttk.Label(application_box, textvariable=self.application_var, wraplength=770, justify="left").pack(fill="x")
        row += 1

        btns = ttk.Frame(outer)
        btns.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="Help regole…", command=lambda: open_extraction_help(self)).pack(side="left")
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Salva regola", command=self.save).pack(side="right")

        for var in self.vars.values():
            try:
                var.trace_add("write", lambda *_args: self.after_idle(self._update_live_preview))
            except Exception:
                pass
        self.after_idle(self._update_live_preview)

        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def _sync_policy_choices(self):
        field_type = FIELD_LABEL_TO_TYPE.get(self.vars["field_type"].get(), "text")
        values = [UPDATE_POLICY_LABELS["replace"], UPDATE_POLICY_LABELS["keep"]]
        if field_type == "number":
            values.append(UPDATE_POLICY_LABELS["sum"])
        try:
            self.policy_combo.configure(values=values)
        except Exception:
            return
        if self.vars["existing_value_policy"].get() not in values:
            self.vars["existing_value_policy"].set(UPDATE_POLICY_LABELS["replace"])

    def _safe_int(self, key: str, default: int = 1) -> int:
        try:
            return int(self.vars[key].get().strip() or default)
        except Exception:
            return default

    def _rule_from_vars(self, validate: bool = False) -> dict:
        name = self.vars["name"].get().strip()
        field_name = slugify(self.vars["field"].get(), "campo")
        strategy = self.vars["strategy"].get() or "regex"
        pattern = self.vars["pattern"].get().strip()
        trigger = self.vars["trigger"].get().strip()
        value_pattern = self.vars["value_pattern"].get().strip()

        if validate:
            if not name:
                raise ValueError("Inserisci il nome della regola")
            if not field_name:
                raise ValueError("Inserisci il nome del campo")
            if strategy == "regex" and not pattern:
                raise ValueError("Per la strategia regex devi specificare la Regex")
            if strategy == "regex":
                re.compile(expand_regex(pattern))
            if strategy == "after_trigger" and not trigger:
                raise ValueError("Per after_trigger devi specificare il Trigger")
            if value_pattern:
                re.compile(expand_regex(value_pattern))

        rule = {
            "name": name or "regola",
            "field": field_name,
            "source": self.vars["source"].get() or "body",
            "strategy": strategy,
            "transform": self.vars["transform"].get() or "text",
            "existing_value_policy": UPDATE_POLICY_FROM_LABEL.get(
                self.vars["existing_value_policy"].get(), "replace"
            ),
        }
        occurrence = self._safe_int("occurrence", 1)
        if occurrence != 1:
            rule["occurrence"] = occurrence
        if strategy == "regex":
            rule["pattern"] = pattern
            rule["group"] = self._safe_int("group", 1)
        elif strategy == "after_trigger":
            rule["trigger"] = trigger
            rule["nonempty_offset"] = self._safe_int("nonempty_offset", 1)
            if value_pattern:
                rule["value_pattern"] = value_pattern
                rule["group"] = self._safe_int("group", 1)
        elif strategy == "constant":
            rule["value"] = self.vars["value"].get()
        return rule

    def _update_live_preview(self):
        try:
            rule = self._rule_from_vars(validate=False)
            field_type = FIELD_LABEL_TO_TYPE.get(self.vars["field_type"].get(), "text")
            self.human_var.set(humanize_extraction_rule(
                rule,
                field_type=field_type,
                currency=bool(self.vars["currency"].get()),
                reconcile_key=bool(self.vars["key"].get()),
            ))

            if self.current_mail is None:
                self.application_var.set("Nessuna mail è aperta nella finestra principale. Apri una mail per provare la regola in tempo reale.")
                return
            trace = extract_rule_trace(self.current_mail, rule)
            if trace.error:
                self.application_var.set(f"Errore nell'applicazione: {trace.error}")
            elif trace.value not in (None, ""):
                raw_text = "" if trace.raw_value is None else str(trace.raw_value)
                converted = str(trace.value)
                if raw_text == converted:
                    self.application_var.set(f"✓ Valore estratto: {converted}")
                else:
                    self.application_var.set(f"✓ Testo trovato: {raw_text}  →  valore memorizzato: {converted}")
            elif trace.matched:
                self.application_var.set("La regola ha trovato la zona prevista, ma il valore finale è vuoto o non convertibile.")
            else:
                self.application_var.set("✗ La regola, con i parametri attuali, non trova alcun valore nella mail corrente.")
        except Exception as exc:
            self.human_var.set(f"Regola incompleta: {exc}")
            self.application_var.set("Completa i parametri per provarla sulla mail corrente.")

    def save(self):
        try:
            rule = self._rule_from_vars(validate=True)
            field_name = rule["field"]
            field_type = FIELD_LABEL_TO_TYPE.get(self.vars["field_type"].get(), "text")
            if rule.get("existing_value_policy") == "sum" and field_type != "number":
                raise ValueError("La politica Somma può essere usata solo con un campo di tipo Numero")
            fmt = "currency_eur" if self.vars["currency"].get() and field_type == "number" else None
            spec = ensure_field(self.profile, field_name, field_type, display_format=fmt)
            if not fmt:
                spec.pop("format", None)

            keys = self.profile.setdefault("reconcile_keys", [])
            if self.vars["key"].get():
                if field_name not in keys:
                    keys.append(field_name)
            elif field_name in keys:
                keys.remove(field_name)

            self.result = rule
            self.destroy()
        except re.error as exc:
            messagebox.showerror("Regex non valida", str(exc), parent=self)
        except Exception as exc:
            messagebox.showerror("Regola", str(exc), parent=self)


class FieldEditorDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, profile: dict, field: dict | None = None):
        super().__init__(parent)
        self.profile = profile
        self.field = dict(field or {})
        self.result: dict | None = None
        self.title("Modifica campo" if field else "Nuovo campo")
        self.resizable(False, False)
        self.name_var = tk.StringVar(value=self.field.get("name", ""))
        self.type_var = tk.StringVar(value=FIELD_TYPE_LABELS[normalize_field_type(self.field.get("type"))])
        self.currency_var = tk.BooleanVar(value=self.field.get("format") == "currency_eur")
        self.key_var = tk.BooleanVar(value=self.field.get("name") in profile.get("reconcile_keys", []))
        self.mail_type_counter_var = tk.StringVar(value=str(self.field.get("mail_type_counter", "") or ""))
        self.final_formula_var = tk.StringVar(
            value=FINAL_FORMULA_LABELS.get(str(self.field.get("final_formula", "") or "").lower(), "Nessuna")
        )
        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Nome campo").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.name_var, width=36).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="Tipo cella Excel").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(frm, textvariable=self.type_var, values=list(FIELD_LABEL_TO_TYPE), state="readonly", width=20).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Checkbutton(frm, text="Formato valuta €", variable=self.currency_var).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Checkbutton(frm, text="Chiave di riconciliazione", variable=self.key_var).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Label(frm, text="Contatore riscontri Tipo mail").grid(row=4, column=0, sticky="w", pady=4)
        mail_type_names = [
            str(mt.get("name", "")) for mt in profile.get("mail_types", [])
            if isinstance(mt, dict) and mt.get("name")
        ]
        self.mail_type_counter_combo = ttk.Combobox(
            frm, textvariable=self.mail_type_counter_var,
            values=[""] + mail_type_names, state="readonly", width=28,
        )
        self.mail_type_counter_combo.grid(row=4, column=1, sticky="w", pady=4)
        self.mail_type_counter_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self.type_var.set("Numero") if self.mail_type_counter_var.get().strip() else None,
        )
        ttk.Label(
            frm,
            text="Se impostato: 1 al primo match sulla riga, poi 2, 3, ...; il campo deve essere Numero.",
            foreground="#555555",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 4))
        ttk.Label(frm, text="Formula finale Excel").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=self.final_formula_var,
            values=list(FINAL_FORMULA_FROM_LABEL), state="readonly", width=20,
        ).grid(row=6, column=1, sticky="w", pady=4)
        ttk.Label(
            frm,
            text="Somma: solo campi Numero. Conta: conta le celle non vuote della colonna.",
            foreground="#555555",
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(0, 4))
        btns = ttk.Frame(frm); btns.grid(row=8, column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=(6,0))
        ttk.Button(btns, text="Salva", command=self.save).pack(side="right")
        self.transient(parent); self.grab_set(); center_dialog_later(self, parent)

    def save(self):
        name = slugify(self.name_var.get(), "campo")
        if not name:
            return
        field_type = FIELD_LABEL_TO_TYPE.get(self.type_var.get(), "text")
        mail_type_counter = self.mail_type_counter_var.get().strip()
        if mail_type_counter and field_type != "number":
            messagebox.showerror(
                "Contatore Tipo mail",
                "Un campo usato come contatore dei riscontri Tipo mail deve essere di tipo Numero.",
                parent=self,
            )
            return
        final_formula = FINAL_FORMULA_FROM_LABEL.get(self.final_formula_var.get(), "")
        if final_formula == "sum" and field_type != "number":
            messagebox.showerror(
                "Formula finale Excel",
                "La formula Somma può essere usata solo su un campo di tipo Numero.",
                parent=self,
            )
            return
        result = {"name": name, "type": field_type}
        if self.currency_var.get() and field_type == "number":
            result["format"] = "currency_eur"
        if mail_type_counter:
            result["mail_type_counter"] = mail_type_counter
        if final_formula:
            result["final_formula"] = final_formula
        self.result = result
        self.destroy()



class HomeSummaryFieldsDialog(tk.Toplevel):
    """Sceglie fino a sei valori riepilogativi da mostrare nella home."""

    def __init__(self, parent: tk.Misc, profile: dict, on_change=None):
        super().__init__(parent)
        self.profile = profile
        self.on_change = on_change
        self.title("Valori nella schermata iniziale")
        self.resizable(False, False)

        candidates: list[tuple[str, str]] = []
        field_defs = {
            str(f.get("name", "")): f
            for f in profile.get("fields", []) or []
            if isinstance(f, dict) and f.get("name")
        }
        computed_names = {
            str(x.get("field", "") or "")
            for x in profile.get("computed_fields", []) or []
            if isinstance(x, dict) and x.get("field")
        }

        for name, field in field_defs.items():
            kind = str(field.get("final_formula", "") or "").lower()
            if kind == "sum":
                candidates.append((name, f"Somma {name.replace('_', ' ')}"))
            elif kind == "count":
                candidates.append((name, f"Conta {name.replace('_', ' ')}"))

        for name in sorted(computed_names):
            # Se il campo calcolato ha già Somma/Conta la voce precedente è più
            # significativa; altrimenti è comunque selezionabile e mostra
            # l'ultimo valore calcolato disponibile nell'Excel.
            if str(field_defs.get(name, {}).get("final_formula", "") or "").lower() not in {"sum", "count"}:
                candidates.append((f"computed:{name}", f"Campo calcolato: {name.replace('_', ' ')}"))

        for item in profile.get("summary_formulas", []) or []:
            if not isinstance(item, dict) or not item.get("name"):
                continue
            name = str(item.get("name"))
            candidates.append((f"summary:{name}", f"Formula riepilogativa: {name}"))

        self._ref_to_label = dict(candidates)
        self._label_to_ref = {label: ref for ref, label in candidates}
        values = [""] + [label for _ref, label in candidates]

        selected_raw = [str(x) for x in profile.get("home_summary_fields", []) or []][:6]
        selected = selected_raw
        labels = [self._ref_to_label.get(x, "") for x in selected]
        while len(labels) < 6:
            labels.append("")
        self.vars = [tk.StringVar(value=labels[i]) for i in range(6)]

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(
            frm,
            text=(
                "Scegli fino a sei valori da mostrare sul bottone del profilo. "
                "Sono disponibili Somma/Conta finali, formule riepilogative e campi calcolati."
            ),
            foreground="#555555", wraplength=620, justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        for i, var in enumerate(self.vars, 1):
            ttk.Label(frm, text=f"Campo {i}").grid(row=i, column=0, sticky="w", pady=4)
            ttk.Combobox(
                frm, textvariable=var, values=values, state="readonly", width=52
            ).grid(row=i, column=1, sticky="ew", pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Salva", command=self.save).pack(side="right")
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def save(self):
        refs: list[str] = []
        for var in self.vars:
            label = var.get().strip()
            if not label:
                continue
            ref = self._label_to_ref.get(label, "")
            if ref and ref not in refs:
                refs.append(ref)
        chosen = [v.get().strip() for v in self.vars if v.get().strip()]
        if len(refs) != len(chosen):
            messagebox.showerror(
                "Valori schermata iniziale",
                "Ogni valore può essere selezionato una sola volta.",
                parent=self,
            )
            return
        self.profile["home_summary_fields"] = refs[:6]
        if self.on_change:
            self.on_change()
        self.destroy()


class ComputedFieldEditorDialog(tk.Toplevel):
    """Editor di un campo calcolato globale del profilo."""

    def __init__(self, parent: tk.Misc, profile: dict, item: dict | None = None):
        super().__init__(parent)
        self.profile = profile
        self.item = dict(item or {})
        self.result: dict | None = None
        self.title("Modifica campo calcolato" if item else "Nuovo campo calcolato")
        self.resizable(False, False)

        fields = [
            f for f in profile.get("fields", [])
            if isinstance(f, dict) and f.get("name")
        ]
        all_names = [str(f.get("name")) for f in fields]
        date_names = [
            str(f.get("name")) for f in fields
            if normalize_field_type(f.get("type")) == "date"
        ]
        number_names = [
            str(f.get("name")) for f in fields
            if normalize_field_type(f.get("type")) == "number"
        ]

        self.field_var = tk.StringVar(value=str(self.item.get("field", "") or ""))
        self.kind_var = tk.StringVar(
            value=COMPUTED_KIND_LABELS.get(str(self.item.get("kind", "days_between") or "days_between"),
                                           COMPUTED_KIND_LABELS["days_between"])
        )
        self.start_var = tk.StringVar(value=str(self.item.get("start_field", "") or ""))
        self.end_var = tk.StringVar(value=str(self.item.get("end_field", "") or ""))
        self.priority_var = tk.StringVar(
            value=COMPUTED_PRIORITY_LABELS.get(
                str(self.item.get("priority", "rule_wins") or "rule_wins"),
                COMPUTED_PRIORITY_LABELS["rule_wins"],
            )
        )

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Campo destinazione").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=self.field_var,
            values=number_names or all_names, state="readonly", width=36,
        ).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="Tipo calcolo").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=self.kind_var,
            values=list(COMPUTED_KIND_FROM_LABEL), state="readonly", width=36,
        ).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="Campo data iniziale").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=self.start_var,
            values=date_names or all_names, state="readonly", width=36,
        ).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="Campo data finale").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=self.end_var,
            values=date_names or all_names, state="readonly", width=36,
        ).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="Priorità").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frm, textvariable=self.priority_var,
            values=list(COMPUTED_PRIORITY_FROM_LABEL), state="readonly", width=48,
        ).grid(row=4, column=1, sticky="ew", pady=4)
        ttk.Label(
            frm,
            text=(
                "Con 'Non sovrascrivere...' una regola esplicita della mail corrente ha priorità. "
                "Esempio: una cancellazione può impostare durata_giorni=0 senza che il calcolo tra date la ripristini."
            ),
            foreground="#555555", wraplength=600, justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(2, 8))

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Salva", command=self.save).pack(side="right")
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def save(self):
        field = self.field_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        kind = COMPUTED_KIND_FROM_LABEL.get(self.kind_var.get(), "days_between")
        priority = COMPUTED_PRIORITY_FROM_LABEL.get(self.priority_var.get(), "rule_wins")
        if not field or not start or not end:
            messagebox.showerror(
                "Campo calcolato",
                "Seleziona campo destinazione, data iniziale e data finale.",
                parent=self,
            )
            return
        if start == end:
            messagebox.showerror(
                "Campo calcolato", "I due campi data devono essere diversi.", parent=self
            )
            return
        dest_spec = get_field(self.profile, field)
        if kind == "days_between" and dest_spec and normalize_field_type(dest_spec.get("type")) != "number":
            messagebox.showerror(
                "Campo calcolato",
                "La Differenza giorni richiede un campo destinazione di tipo Numero.",
                parent=self,
            )
            return
        for source_name in (start, end):
            spec = get_field(self.profile, source_name)
            if spec and normalize_field_type(spec.get("type")) != "date":
                messagebox.showerror(
                    "Campo calcolato",
                    f"Il campo '{source_name}' deve essere di tipo Data.",
                    parent=self,
                )
                return
        self.result = {
            "field": field,
            "kind": kind,
            "start_field": start,
            "end_field": end,
            "priority": priority,
        }
        self.destroy()


class ComputedFieldsDialog(tk.Toplevel):
    """Gestione GUI dei campi calcolati globali del profilo."""

    def __init__(self, parent: tk.Misc, profile: dict, on_change=None):
        super().__init__(parent)
        self.profile = profile
        self.on_change = on_change
        self.title("Campi calcolati del profilo")
        self.geometry("1040x430")
        self.minsize(780, 360)

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        cols = ("field", "kind", "start", "end", "priority")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for c, lab, width in [
            ("field", "Campo destinazione", 170),
            ("kind", "Tipo calcolo", 220),
            ("start", "Campo iniziale", 150),
            ("end", "Campo finale", 150),
            ("priority", "Priorità", 320),
        ]:
            self.tree.heading(c, text=lab)
            self.tree.column(c, width=width, stretch=True)
        self.tree.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text=(
                "Questi calcoli sono globali al profilo e vengono eseguiti dopo le regole di estrazione. "
                "La priorità stabilisce se possono sovrascrivere un valore scritto esplicitamente dalla mail corrente."
            ),
            foreground="#555555", wraplength=980, justify="left",
        ).pack(fill="x", pady=(6, 0))

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Nuovo", command=self.add).pack(side="left")
        ttk.Button(btns, text="Modifica", command=self.edit).pack(side="left", padx=5)
        ttk.Button(btns, text="Elimina", command=self.delete).pack(side="left", padx=5)
        ttk.Button(btns, text="Chiudi", command=self.destroy).pack(side="right")
        self.tree.bind("<Double-1>", lambda _e: self.edit())
        self.refresh()
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, comp in enumerate(self.profile.get("computed_fields", []) or []):
            if not isinstance(comp, dict):
                continue
            kind = COMPUTED_KIND_LABELS.get(str(comp.get("kind", "") or ""), str(comp.get("kind", "") or ""))
            priority = COMPUTED_PRIORITY_LABELS.get(
                str(comp.get("priority", "rule_wins") or "rule_wins"),
                COMPUTED_PRIORITY_LABELS["rule_wins"],
            )
            self.tree.insert("", "end", iid=str(i), values=(
                comp.get("field", ""), kind, comp.get("start_field", ""),
                comp.get("end_field", ""), priority,
            ))

    def _changed(self):
        if self.on_change:
            self.on_change()

    def add(self):
        dlg = ComputedFieldEditorDialog(self, self.profile)
        self.wait_window(dlg)
        if not dlg.result:
            return
        existing = {
            str(x.get("field", "")) for x in self.profile.get("computed_fields", [])
            if isinstance(x, dict)
        }
        if dlg.result["field"] in existing:
            messagebox.showerror(
                "Campo calcolato",
                "Esiste già un calcolo per questo campo destinazione.",
                parent=self,
            )
            return
        self.profile.setdefault("computed_fields", []).append(dlg.result)
        self.refresh()
        self._changed()

    def edit(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        items = self.profile.setdefault("computed_fields", [])
        dlg = ComputedFieldEditorDialog(self, self.profile, items[i])
        self.wait_window(dlg)
        if not dlg.result:
            return
        for j, item in enumerate(items):
            if j != i and isinstance(item, dict) and item.get("field") == dlg.result["field"]:
                messagebox.showerror(
                    "Campo calcolato",
                    "Esiste già un calcolo per questo campo destinazione.",
                    parent=self,
                )
                return
        items[i] = dlg.result
        self.refresh()
        self._changed()

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        items = self.profile.setdefault("computed_fields", [])
        item = items[i]
        if messagebox.askyesno(
            "Elimina campo calcolato",
            f"Eliminare il calcolo del campo '{item.get('field', '')}'?",
            parent=self,
        ):
            del items[i]
            self.refresh()
            self._changed()


class FieldGenerationRulesDialog(tk.Toplevel):
    """Mostra tutte le regole che possono valorizzare un campo del profilo."""

    def __init__(self, parent: tk.Misc, profile: dict, field_name: str):
        super().__init__(parent)
        self.profile = profile
        self.field_name = str(field_name or "")
        self.items = field_generation_sources(profile, self.field_name)
        self.title(f"Regole che generano il campo: {self.field_name}")
        self.geometry("1180x560")
        self.minsize(820, 420)

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text=f"Campo: {self.field_name}",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            frm,
            text=(
                "Sono elencati tutti i meccanismi del profilo che possono scrivere il valore del campo "
                "nelle righe dati: regole di estrazione dei Tipi mail, regole generali/calcoli automatici "
                "e contatori di Tipo mail. Le formule finali Excel non sono incluse perché producono solo "
                "la riga riepilogativa in fondo ai dati."
            ),
            foreground="#555555", wraplength=1120, justify="left",
        ).pack(fill="x", pady=(3, 8))

        cols = ("kind", "scope", "mail_type", "name", "source", "mechanism", "detail")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for c, lab, width in [
            ("kind", "Tipo", 135),
            ("scope", "Ambito", 115),
            ("mail_type", "Tipo mail", 165),
            ("name", "Regola", 175),
            ("source", "Origine", 110),
            ("mechanism", "Strategia / calcolo", 200),
            ("detail", "Dettaglio", 380),
        ]:
            self.tree.heading(c, text=lab)
            self.tree.column(c, width=width, stretch=True)
        self.tree.pack(fill="both", expand=True)

        detail_frame = ttk.LabelFrame(frm, text="Dettaglio della regola selezionata", padding=6)
        detail_frame.pack(fill="both", expand=False, pady=(8, 0))
        self.detail = tk.Text(detail_frame, height=7, wrap="word")
        detail_scroll = ttk.Scrollbar(detail_frame, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=detail_scroll.set)
        self.detail.pack(side="left", fill="both", expand=True)
        detail_scroll.pack(side="right", fill="y")

        if not self.items:
            self.detail.insert(
                "1.0",
                "Nessuna regola o meccanismo del profilo genera direttamente questo campo.\n"
                "Il campo può restare vuoto oppure essere popolato da dati già presenti nell'Excel caricato.",
            )
            self.detail.configure(state="disabled")
        else:
            for i, item in enumerate(self.items):
                self.tree.insert("", "end", iid=str(i), values=(
                    item.get("kind_label", ""),
                    item.get("scope", ""),
                    item.get("mail_type", ""),
                    item.get("name", ""),
                    item.get("source", ""),
                    item.get("mechanism", ""),
                    item.get("detail", ""),
                ))
            self.tree.selection_set("0")
            self.tree.focus("0")
            self._show_selected_detail()

        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_selected_detail())
        self.tree.bind("<Double-1>", lambda _e: self._show_selected_detail())

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Chiudi", command=self.destroy).pack(side="right")

        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def _show_selected_detail(self):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            item = self.items[int(sel[0])]
        except (ValueError, IndexError):
            return
        raw = item.get("raw", {})
        lines = [
            f"Tipo: {item.get('kind_label', '')}",
            f"Ambito: {item.get('scope', '')}",
        ]
        if item.get("mail_type"):
            lines.append(f"Tipo mail: {item.get('mail_type')}")
        lines.extend([
            f"Nome: {item.get('name', '')}",
            f"Origine: {item.get('source', '')}",
            f"Strategia / calcolo: {item.get('mechanism', '')}",
            f"Dettaglio: {item.get('detail', '')}",
            "",
            "Configurazione JSON:",
            json.dumps(raw, indent=2, ensure_ascii=False),
        ])
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", "\n".join(lines))
        self.detail.configure(state="disabled")



class SummaryFormulaEditorDialog(tk.Toplevel):
    """Editor di una formula riepilogativa nominata o aggregazione temporale."""

    def __init__(self, parent: tk.Misc, profile: dict, item: dict | None = None):
        super().__init__(parent)
        self.profile = profile
        self.item = dict(item or {})
        self.result: dict | None = None
        self.title("Modifica formula riepilogativa" if item else "Nuova formula riepilogativa")
        self.geometry("790x540")
        self.minsize(700, 500)

        self.name_var = tk.StringVar(value=str(self.item.get("name", "") or ""))
        initial_kind = "Aggregazione temporale" if str(self.item.get("kind", "") or "") == "period_aggregate" else "Formula composta"
        self.kind_var = tk.StringVar(value=initial_kind)
        self.formula_var = tk.StringVar(value=str(self.item.get("formula", "=") or "="))
        self.operation_var = tk.StringVar(value="Conta" if str(self.item.get("operation", "sum")) == "count" else "Somma")
        self.field_var = tk.StringVar(value=str(self.item.get("field", "") or ""))
        self.date_field_var = tk.StringVar(value=str(self.item.get("date_field", "") or ""))
        self.period_var = tk.StringVar(value="Ultimo anno" if str(self.item.get("period", "month")) == "year" else "Ultimo mese")
        self.period_mode_var = tk.StringVar(
            value="Periodo di calendario corrente"
            if str(self.item.get("period_mode", "rolling")) == "calendar"
            else "Finestra mobile"
        )

        fields = [
            str(f.get("name")) for f in profile.get("fields", []) or []
            if isinstance(f, dict) and f.get("name")
        ]
        date_fields = [
            str(f.get("name")) for f in profile.get("fields", []) or []
            if isinstance(f, dict) and f.get("name") and normalize_field_type(f.get("type")) == "date"
        ]

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Nome").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.name_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frm, text="Tipo").grid(row=1, column=0, sticky="w", pady=4)
        kind_combo = ttk.Combobox(
            frm, textvariable=self.kind_var,
            values=["Formula composta", "Aggregazione temporale"],
            state="readonly", width=28,
        )
        kind_combo.grid(row=1, column=1, sticky="w", pady=4)

        self.formula_frame = ttk.LabelFrame(frm, text="Formula composta", padding=8)
        self.formula_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(6, 4))
        self.formula_frame.columnconfigure(1, weight=1)
        ttk.Label(self.formula_frame, text="Formula").grid(row=0, column=0, sticky="w", pady=4)
        self.formula_entry = ttk.Entry(self.formula_frame, textvariable=self.formula_var)
        self.formula_entry.grid(row=0, column=1, sticky="ew", pady=4)

        aliases = sorted(final_formula_aliases(profile))
        alias_box = ttk.LabelFrame(self.formula_frame, text="Alias disponibili dai campi con Formula finale", padding=8)
        alias_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 4))
        alias_list = tk.Listbox(alias_box, height=6, exportselection=False)
        alias_scroll = ttk.Scrollbar(alias_box, orient="vertical", command=alias_list.yview)
        alias_list.configure(yscrollcommand=alias_scroll.set)
        alias_list.pack(side="left", fill="both", expand=True)
        alias_scroll.pack(side="right", fill="y")
        if aliases:
            for alias in aliases:
                alias_list.insert("end", alias)
            alias_list.bind("<ButtonRelease-1>", self._insert_alias_from_click)
        else:
            alias_list.insert("end", "Nessun alias disponibile. Imposta Somma o Conta in almeno un campo.")
            alias_list.configure(state="disabled")
        self.alias_list = alias_list

        self.aggregate_frame = ttk.LabelFrame(frm, text="Aggregazione temporale", padding=8)
        self.aggregate_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 4))
        ttk.Label(self.aggregate_frame, text="Operazione").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Combobox(
            self.aggregate_frame, textvariable=self.operation_var,
            values=["Somma", "Conta"], state="readonly", width=16
        ).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(self.aggregate_frame, text="Campo").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Combobox(
            self.aggregate_frame, textvariable=self.field_var,
            values=fields, state="readonly", width=32
        ).grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(self.aggregate_frame, text="Campo data").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Combobox(
            self.aggregate_frame, textvariable=self.date_field_var,
            values=date_fields, state="readonly", width=32
        ).grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(self.aggregate_frame, text="Periodo").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Combobox(
            self.aggregate_frame, textvariable=self.period_var,
            values=["Ultimo mese", "Ultimo anno"], state="readonly", width=20
        ).grid(row=3, column=1, sticky="w", pady=3)
        ttk.Label(self.aggregate_frame, text="Interpretazione").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Combobox(
            self.aggregate_frame, textvariable=self.period_mode_var,
            values=["Finestra mobile", "Periodo di calendario corrente"],
            state="readonly", width=30
        ).grid(row=4, column=1, sticky="w", pady=3)
        ttk.Label(
            self.aggregate_frame,
            text=(
                "Finestra mobile: ultimi 30 giorni / 365 giorni. "
                "Periodo di calendario: mese corrente / anno corrente."
            ),
            foreground="#555555", wraplength=650, justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(5, 0))

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Salva", command=self.save).pack(side="right")
        kind_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_mode())
        self._update_mode()
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def _update_mode(self):
        aggregate = self.kind_var.get() == "Aggregazione temporale"
        for child in self.formula_frame.winfo_children():
            try:
                child.configure(state="disabled" if aggregate else "normal")
            except Exception:
                pass
        for child in self.aggregate_frame.winfo_children():
            try:
                child.configure(state="normal" if aggregate else "disabled")
            except Exception:
                pass

    def _insert_alias_from_click(self, event=None):
        if not getattr(self, "alias_list", None):
            return
        try:
            index = self.alias_list.nearest(event.y) if event is not None else self.alias_list.curselection()[0]
            alias = str(self.alias_list.get(index) or "").strip()
        except Exception:
            return
        if not alias or alias.startswith("Nessun alias"):
            return
        try:
            self.formula_entry.focus_set()
            self.formula_entry.insert("insert", alias)
        except Exception:
            self.formula_var.set(self.formula_var.get() + alias)

    def save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Formula riepilogativa", "Inserisci il nome della formula.", parent=self)
            return

        if self.kind_var.get() == "Aggregazione temporale":
            field = self.field_var.get().strip()
            date_field = self.date_field_var.get().strip()
            if not field or not date_field:
                messagebox.showerror(
                    "Formula riepilogativa",
                    "Scegli il campo da aggregare e il campo data.",
                    parent=self,
                )
                return
            self.result = {
                "name": name,
                "kind": "period_aggregate",
                "operation": "count" if self.operation_var.get() == "Conta" else "sum",
                "field": field,
                "date_field": date_field,
                "period": "year" if self.period_var.get() == "Ultimo anno" else "month",
                "period_mode": "calendar" if self.period_mode_var.get() == "Periodo di calendario corrente" else "rolling",
            }
            self.destroy()
            return

        formula = self.formula_var.get().strip()
        if not formula:
            messagebox.showerror("Formula riepilogativa", "Inserisci la formula.", parent=self)
            return
        if not formula.startswith("="):
            formula = "=" + formula
        try:
            aliases = final_formula_aliases(self.profile)
            fake_cells = {alias: "A1" for alias in aliases}
            expand_summary_formula(formula, self.profile, fake_cells)
        except Exception as exc:
            messagebox.showerror("Formula riepilogativa", str(exc), parent=self)
            return
        self.result = {"name": name, "formula": formula}
        self.destroy()


class SummaryFormulasDialog(tk.Toplevel):
    """Gestione della lista di formule riepilogative del profilo."""

    def __init__(self, parent: tk.Misc, profile: dict, on_change=None):
        super().__init__(parent)
        self.profile = profile
        self.on_change = on_change
        self.title("Formule riepilogative")
        self.geometry("860x430")
        self.minsize(700, 360)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        cols = ("name", "formula")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("name", text="Nome")
        self.tree.heading("formula", text="Formula con alias")
        self.tree.column("name", width=220, stretch=False)
        self.tree.column("formula", width=560, stretch=True)
        self.tree.pack(fill="both", expand=True)

        aliases = sorted(final_formula_aliases(profile))
        aliases_text = ", ".join(aliases) if aliases else "nessuno"
        ttk.Label(
            frm,
            text=f"Alias disponibili: {aliases_text}",
            foreground="#555555",
            wraplength=800,
            justify="left",
        ).pack(fill="x", pady=(6, 0))

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Nuova", command=self.add).pack(side="left")
        ttk.Button(btns, text="Modifica", command=self.edit).pack(side="left", padx=5)
        ttk.Button(btns, text="Elimina", command=self.delete).pack(side="left", padx=5)
        ttk.Button(btns, text="Chiudi", command=self.destroy).pack(side="right")
        self.tree.bind("<Double-1>", lambda _e: self.edit())
        self.refresh()
        self.transient(parent)
        self.grab_set()
        center_dialog_later(self, parent)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, item in enumerate(self.profile.get("summary_formulas", []) or []):
            if not isinstance(item, dict):
                continue
            self.tree.insert("", "end", iid=str(i), values=(item.get("name", ""), summary_formula_display(item)))

    def _changed(self):
        if callable(self.on_change):
            self.on_change()

    def add(self):
        dlg = SummaryFormulaEditorDialog(self, self.profile)
        self.wait_window(dlg)
        if not dlg.result:
            return
        names = {str(x.get("name", "")).strip().casefold() for x in self.profile.get("summary_formulas", []) if isinstance(x, dict)}
        if dlg.result["name"].casefold() in names:
            messagebox.showerror("Formula riepilogativa", "Esiste già una formula con questo nome.", parent=self)
            return
        self.profile.setdefault("summary_formulas", []).append(dlg.result)
        self.refresh()
        self._changed()

    def edit(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        items = self.profile.setdefault("summary_formulas", [])
        dlg = SummaryFormulaEditorDialog(self, self.profile, items[i])
        self.wait_window(dlg)
        if not dlg.result:
            return
        names = {
            str(x.get("name", "")).strip().casefold()
            for j, x in enumerate(items) if j != i and isinstance(x, dict)
        }
        if dlg.result["name"].casefold() in names:
            messagebox.showerror("Formula riepilogativa", "Esiste già una formula con questo nome.", parent=self)
            return
        old_name = str(items[i].get("name", "") or "")
        items[i] = dlg.result
        new_name = str(dlg.result.get("name", "") or "")
        if old_name and new_name and old_name != new_name:
            self.profile["home_summary_fields"] = [
                f"summary:{new_name}" if x == f"summary:{old_name}" else x
                for x in self.profile.get("home_summary_fields", []) or []
            ][:6]
        self.refresh()
        self._changed()

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0])
        items = self.profile.setdefault("summary_formulas", [])
        item = items[i]
        if messagebox.askyesno(
            "Elimina formula",
            f"Eliminare la formula riepilogativa '{item.get('name', '')}'?",
            parent=self,
        ):
            old_name = str(item.get("name", "") or "")
            del items[i]
            if old_name:
                self.profile["home_summary_fields"] = [
                    x for x in self.profile.get("home_summary_fields", []) or []
                    if x != f"summary:{old_name}"
                ][:6]
            self.refresh()
            self._changed()


class MainWindow(tk.Tk):
    def __init__(self, config: ConfigManager, profile_path: Path):
        super().__init__()
        self.config_mgr = config
        set_regex_abbreviations(config.regex_abbreviations)
        self.profile_mgr = ProfileManager(config.profiles_dir)
        self.profile_path = profile_path
        self.profile = self.profile_mgr.load(profile_path)
        # La load normalizza anche i vecchi criteri dei tipi mail nelle nuove
        # regole di selezione v1.5. Salviamo subito la migrazione nel JSON.
        self.profile_mgr.save(profile_path, self.profile)
        self.imap = ImapService(config)
        self.title(f"EgoMailExtractor — {self.profile.get('name', profile_path.stem)}")
        self.resizable(True, True)
        self._apply_initial_geometry(config.data.get("window", {}).get("geometry", "1450x900"))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.stop_event = threading.Event()
        # Evento separato per interrompere l'estrazione senza interferire con la ricerca IMAP.
        self.extraction_stop_event = threading.Event()
        self.extraction_running = False
        self.ui_queue: queue.Queue = queue.Queue()
        self.headers: dict[int, MailHeader] = {}
        self.messages: dict[int, MailMessage] = {}
        self.row_matches: dict[int, set[str]] = {}
        self.search_order: list[int] = []
        self.filter_mail_types: set[str] = set()
        self.mail_type_columns: list[tuple[str, str]] = []
        self.selected_uids: set[int] = set()
        self.current_uid: int | None = None
        self.current_mail: MailMessage | None = None
        self.last_trigger = ""
        self.preview_source_offsets: dict[str, int] = {}
        self.hover_tip = HoverTip(self)

        self.subject_var = tk.StringVar()
        self.sender_var = tk.StringVar()
        self.folder_var = tk.StringVar(value=config.imap.folder)
        self.imap_profile_var = tk.StringVar(value=config.active_imap_profile_name or "<nessuna email>")
        self.status_var = tk.StringVar(value="Pronto")
        self.mail_type_var = tk.StringVar()
        self.type_filter_var = tk.StringVar(value="Filtro tipi mail: tutti")

        self._build_menu()
        self._build_ui()
        self._configure_tags()
        self._refresh_mail_types()
        self.after(120, self._restore_main_paned_sash)
        self.after(100, self._drain_queue)

    def _apply_initial_geometry(self, saved_geometry: str) -> None:
        """Adatta la finestra allo schermo senza nascondere i comandi inferiori."""
        self.update_idletasks()
        sw = max(800, self.winfo_screenwidth())
        sh = max(600, self.winfo_screenheight())

        # Margine per barra applicazioni/dock e bordi del window manager.
        max_w = max(760, sw - 80)
        max_h = max(500, sh - 120)
        min_w = min(900, max_w)
        min_h = min(540, max_h)
        self.minsize(min_w, min_h)

        match = re.match(r"^(\d+)x(\d+)", str(saved_geometry or ""))
        wanted_w = int(match.group(1)) if match else 1450
        wanted_h = int(match.group(2)) if match else 900
        width = max(min_w, min(wanted_w, max_w))
        height = max(min_h, min(wanted_h, max_h))
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _restore_main_paned_sash(self) -> None:
        """Rende l'anteprima inizialmente più piccola e ripristina la scelta dell'utente."""
        try:
            self.update_idletasks()
            total = max(1, self.main_paned.winfo_height())
            ratio = float(self.config_mgr.data.get("window", {}).get("main_sash_ratio", 0.68))
            ratio = max(0.48, min(0.82, ratio))
            y = int(total * ratio)
            self.main_paned.sash_place(0, 1, y)
        except Exception:
            pass

    def _on_close(self) -> None:
        try:
            window_cfg = self.config_mgr.data.setdefault("window", {})
            window_cfg["geometry"] = self.geometry()
            try:
                total = max(1, self.main_paned.winfo_height())
                _x, y = self.main_paned.sash_coord(0)
                window_cfg["main_sash_ratio"] = max(0.0, min(1.0, y / total))
            except Exception:
                pass
            self.config_mgr.save()
        except Exception:
            pass
        self.destroy()

    def show_about(self) -> None:
        # Riusa esattamente lo stesso splash dell'avvio, ma senza chiusura
        # automatica: l'utente può leggere i dati e cliccare il sito.
        show_splash(self, duration_ms=0, closable=True)

    def _build_menu(self):
        menu = tk.Menu(self)
        filem = tk.Menu(menu, tearoff=False)
        filem.add_command(label="Schermata iniziale", command=self.show_home_screen)
        openm = tk.Menu(filem, tearoff=False)
        openm.add_command(label="Excel generale del profilo di estrazione", command=self.open_general_excel)
        openm.add_command(label="Cartella degli Excel", command=self.open_work_folder)
        filem.add_cascade(label="Apri", menu=openm)
        filem.add_separator()
        filem.add_command(label="Preferenze…", command=self.preferences)
        filem.add_command(label="Apri cartella email…", command=self.open_imap_profiles_folder)
        filem.add_command(label="Apri cartella profili di estrazione…", command=self.open_extraction_profiles_folder)
        filem.add_command(label="Cambia password master…", command=self.change_master_password)
        filem.add_command(label="Riconnetti / prova email", command=self.reconnect_imap)
        filem.add_separator()
        filem.add_command(label="Esporta mail selezionate in Excel…", command=self.export_selected)
        filem.add_command(label="Estrai nuove mail / aggiorna Excel generale", command=self.extract_incremental)
        filem.add_command(label="Interrompi estrazione", command=self.stop_extraction)
        filem.add_command(label="Apri log estrazione…", command=self.open_extraction_log)
        filem.add_command(label="Apri cartella log…", command=self.open_extraction_log_folder)
        filem.add_command(label="Cancella log…", command=self.clear_extraction_log)
        filem.add_command(label="Esporta pacchetto debug ZIP…", command=self.export_debug_bundle)
        filem.add_separator()
        filem.add_command(label="Esci", command=self._on_close)
        menu.add_cascade(label="File", menu=filem)

        emailm = tk.Menu(menu, tearoff=False)
        emailm.add_command(label="Modifica email attiva…", command=self.open_imap_settings)
        if IS_PRO:
            emailm.add_command(label="Cambia email…", command=self.choose_imap_profile)
            emailm.add_command(label="Gestisci email…", command=self.manage_imap_profiles)
            emailm.add_separator()
            emailm.add_command(label="Profili di estrazione associati all'email…", command=self.manage_email_profile_associations)
        menu.add_cascade(label="Email", menu=emailm)

        profm = tk.Menu(menu, tearoff=False)
        if IS_PRO:
            profm.add_command(label="Cambia profilo…", command=self.change_profile)
        profm.add_command(label="Rinomina profilo…", command=self.rename_profile)
        profm.add_command(label="Data inizio estrazione…", command=self.set_extraction_start_date)
        profm.add_command(label="Riesegui estrazione generale…", command=self.rebuild_general_extraction)
        profm.add_command(label="Tabella totali…", command=lambda: self.show_pro_feature("Tabella totali"))
        profm.add_separator()
        profm.add_command(label="Export profilo di estrazione…", command=self.export_extraction_profiles)
        if IS_PRO:
            profm.add_command(label="Importa profilo di estrazione…", command=self.import_extraction_profiles)
        profm.add_command(label="Backup profilo di estrazione…", command=self.backup_extraction_profile)
        profm.add_command(label="Restore profilo di estrazione…", command=self.restore_extraction_profile)
        profm.add_separator()
        profm.add_command(label="Salva profilo", command=self.save_profile)
        profm.add_command(label="Nuovo tipo mail dal messaggio", command=self.create_mail_type)
        profm.add_command(label="Rinomina tipo mail…", command=self.rename_current_mail_type)
        profm.add_command(label="Elimina tipo mail…", command=self.delete_current_mail_type)
        profm.add_separator()
        profm.add_command(label="Regole selezione mail…", command=self.edit_mail_type_criteria)
        profm.add_command(label="Regole di estrazione…", command=self.manage_rules)
        profm.add_command(label="Gestisci campi e tipi Excel", command=self.manage_fields)
        menu.add_cascade(label="Profilo", menu=profm)

        wiz = tk.Menu(menu, tearoff=False)
        wiz.add_command(label="Configurazione guidata…", command=lambda: self.show_pro_feature("Wizard"))
        wiz.add_separator()
        wiz.add_command(label="Scegli icona profilo di estrazione…", command=lambda: self.show_pro_feature("Wizard icona profilo"))
        wiz.add_command(label="Campo riepilogo / calcolato (Somma o Conteggio)…", command=lambda: self.show_pro_feature("Wizard campi riepilogo"))
        wiz.add_command(label="Formule riepilogative…", command=lambda: self.show_pro_feature("Wizard formule riepilogative"))
        menu.add_cascade(label="Wizard", menu=wiz)

        helpm = tk.Menu(menu, tearoff=False)
        helpm.add_command(label="Regole di estrazione…", command=lambda: open_extraction_help(self))
        helpm.add_command(label="Richiedi assistenza via email…", command=self.request_support_email)
        helpm.add_command(label="Licenza…", command=self.open_license)
        helpm.add_separator()
        helpm.add_command(label="Informazioni su…", command=self.show_about)
        menu.add_cascade(label="Help", menu=helpm)
        self.config(menu=menu)

    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Oggetto contiene").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.subject_var, width=35).grid(row=0, column=1, padx=(4, 12), sticky="ew")
        ttk.Label(top, text="Mittente contiene").grid(row=0, column=2, sticky="w")
        ttk.Entry(top, textvariable=self.sender_var, width=32).grid(row=0, column=3, padx=(4, 12), sticky="ew")
        ttk.Label(top, text="Cartella").grid(row=0, column=4, sticky="w")
        ttk.Entry(top, textvariable=self.folder_var, width=18).grid(row=0, column=5, padx=(4, 12))
        ttk.Label(top, text="Email").grid(row=0, column=6, sticky="w")
        ttk.Label(top, textvariable=self.imap_profile_var, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=7, padx=(4, 12), sticky="w")
        ttk.Button(top, text="Cerca", command=self.start_search).grid(row=0, column=8, padx=3)
        ttk.Button(top, text="Interrompi", command=self.stop_search).grid(row=0, column=9, padx=3)
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        # Pannello messaggi applicazione: viene riservato prima del PanedWindow,
        # così non scompare quando la finestra viene ridotta. Conserva lo
        # storico dei messaggi e dispone di scorrimento verticale.
        bottom = ttk.LabelFrame(self, text="Messaggi applicazione", padding=(6, 4))
        bottom.pack(side="bottom", fill="x", padx=8, pady=(0, 8))
        status_holder = ttk.Frame(bottom)
        status_holder.pack(fill="both", expand=True)
        self.status_text = tk.Text(
            status_holder, height=4, wrap="word", undo=False, state="disabled",
            takefocus=True,
        )
        status_scroll = ttk.Scrollbar(status_holder, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=status_scroll.set)
        self.status_text.grid(row=0, column=0, sticky="nsew")
        status_scroll.grid(row=0, column=1, sticky="ns")
        status_holder.rowconfigure(0, weight=1)
        status_holder.columnconfigure(0, weight=1)
        status_buttons = ttk.Frame(bottom)
        status_buttons.pack(fill="x", pady=(3, 0))
        ttk.Button(status_buttons, text="Pulisci messaggi", command=self._clear_status_messages).pack(side="right")
        self.status_var.trace_add("write", self._on_status_var_changed)
        self._append_status_message(self.status_var.get())

        # Separatore ben visibile e trascinabile: lista mail e anteprima possono
        # essere ridimensionate liberamente dall'utente.
        paned = tk.PanedWindow(
            self, orient=tk.VERTICAL, sashwidth=9, sashrelief=tk.RAISED,
            showhandle=True, bd=0, relief=tk.FLAT,
        )
        paned.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        self.main_paned = paned

        list_frame = ttk.Frame(paned)
        preview_frame = ttk.Frame(paned)
        paned.add(list_frame, minsize=210, stretch="always")
        paned.add(preview_frame, minsize=135, stretch="always")

        # La tabella ha una colonna dinamica per ogni tipo mail del profilo.
        # La spunta indica che TUTTE le regole di selezione del tipo sono vere.
        tree_holder = ttk.Frame(list_frame)
        tree_holder.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_holder, columns=(), show="headings", selectmode="extended")
        y = ttk.Scrollbar(tree_holder, orient="vertical", command=self.tree.yview)
        x = ttk.Scrollbar(tree_holder, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        tree_holder.rowconfigure(0, weight=1)
        tree_holder.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.toggle_selected)

        filter_bar = ttk.Frame(list_frame)
        filter_bar.pack(fill="x", pady=(4, 0))
        ttk.Label(filter_bar, textvariable=self.type_filter_var).pack(side="left")
        ttk.Label(
            filter_bar,
            text=" — puoi selezionare più colonne; i filtri sono in OR",
            foreground="#666666"
        ).pack(side="left")
        ttk.Button(filter_bar, text="Mostra tutti", command=lambda: self.filter_by_mail_type(None)).pack(side="right")

        # Comandi operativi direttamente sotto la lista delle email.
        # In questo modo restano associati alla selezione delle righe e non
        # occupano spazio sotto l'anteprima.
        list_actions = ttk.Frame(list_frame)
        list_actions.pack(fill="x", pady=(5, 0))
        selection_row = ttk.Frame(list_actions)
        selection_row.pack(fill="x")
        ttk.Button(selection_row, text="Seleziona righe evidenziate", command=self.select_tree_rows).pack(side="left")
        ttk.Button(selection_row, text="Deseleziona tutto", command=self.clear_selected).pack(side="left", padx=5)

        extraction_row = ttk.Frame(list_actions)
        extraction_row.pack(fill="x", pady=(4, 0))
        ttk.Button(extraction_row, text="Estrai selezionate in Excel…", command=self.export_selected).pack(side="left")
        ttk.Button(
            extraction_row, text="Estrai nuove mail / aggiorna Excel generale",
            command=self.extract_incremental,
        ).pack(side="left", padx=6)
        ttk.Button(
            extraction_row, text="Interrompi estrazione",
            command=self.stop_extraction,
        ).pack(side="left", padx=2)

        # Barra comandi dell'anteprima.
        tools = ttk.Frame(preview_frame)
        tools.pack(fill="x", pady=(0, 5))

        tools_top = ttk.Frame(tools)
        tools_top.pack(fill="x")
        ttk.Label(tools_top, text="Tipo mail:").pack(side="left")
        self.mail_type_combo = ttk.Combobox(
            tools_top, textvariable=self.mail_type_var, state="readonly", width=28
        )
        self.mail_type_combo.pack(side="left", padx=(4, 10))
        self.mail_type_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_extraction_preview())
        ttk.Button(tools_top, text="Nuovo tipo", command=self.create_mail_type).pack(side="left", padx=2)
        ttk.Button(tools_top, text="Rinomina", command=self.rename_current_mail_type).pack(side="left", padx=2)
        ttk.Button(tools_top, text="Elimina", command=self.delete_current_mail_type).pack(side="left", padx=2)
        ttk.Button(tools_top, text="Selezione mail…", command=self.edit_mail_type_criteria).pack(side="left", padx=2)
        ttk.Button(tools_top, text="Regole estrazione…", command=self.manage_rules).pack(side="left", padx=2)
        ttk.Button(tools_top, text="Campi…", command=self.manage_fields).pack(side="left", padx=2)

        tools_bottom = ttk.Frame(tools)
        tools_bottom.pack(fill="x", pady=(4, 0))
        ttk.Label(tools_bottom, text="Crea regola da selezione:").pack(side="left", padx=(0, 4))
        ttk.Button(tools_bottom, text="Trigger", command=lambda: self.mark_selection("trigger")).pack(side="left", padx=2)
        ttk.Button(tools_bottom, text="Campo testo", command=lambda: self.mark_selection("text")).pack(side="left", padx=2)
        ttk.Button(tools_bottom, text="Numero", command=lambda: self.mark_selection("number")).pack(side="left", padx=2)
        ttk.Button(tools_bottom, text="Importo", command=lambda: self.mark_selection("money")).pack(side="left", padx=2)
        ttk.Button(tools_bottom, text="Data", command=lambda: self.mark_selection("date")).pack(side="left", padx=2)
        ttk.Button(tools_bottom, text="Chiave", command=lambda: self.mark_selection("key")).pack(side="left", padx=2)
        ttk.Button(tools_bottom, text="Salva profilo", command=self.save_profile).pack(side="right", padx=2)

        self.extraction_preview_var = tk.StringVar(value="Anteprima estrazione: apri una mail")
        preview_line = ttk.Frame(preview_frame)
        preview_line.pack(fill="x", pady=(0, 5))
        ttk.Label(preview_line, textvariable=self.extraction_preview_var, anchor="w").pack(side="left", fill="x", expand=True)

        self.text = tk.Text(preview_frame, wrap="word", undo=False)
        sy = ttk.Scrollbar(preview_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sy.set)
        self.text.pack(side="left", fill="both", expand=True)
        sy.pack(side="right", fill="y")

        # La schermata operativa completa viene costruita subito, ma all'avvio
        # è coperta da una home semplificata pensata per l'uso quotidiano.
        self._build_home_overlay()

    def _build_home_overlay(self) -> None:
        self.home_frame = ttk.Frame(self, padding=22)
        self.home_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._home_icon_refs = []

        ttk.Label(self.home_frame, text=DISPLAY_NAME, font=("TkDefaultFont", 22, "bold")).pack(pady=(4, 5))
        ttk.Label(self.home_frame, text="Scegli i dati da estrarre", font=("TkDefaultFont", 12), foreground="#555555").pack(pady=(0, 14))

        imap_box = ttk.Frame(self.home_frame, padding=(8,6))
        imap_box.pack(fill="x", padx=60, pady=(0, 10))
        ttk.Label(imap_box, text="Email:", width=22).pack(side="left")
        self.home_imap_var = tk.StringVar()
        ttk.Label(imap_box, textvariable=self.home_imap_var, font=("TkDefaultFont", 10, "bold")).pack(side="left", fill="x", expand=True)
        if IS_PRO:
            ttk.Button(imap_box, text="Cambia…", command=self._home_change_imap).pack(side="right")
        else:
            ttk.Button(imap_box, text="Modifica…", command=self.open_imap_settings).pack(side="right")

        outer = ttk.LabelFrame(self.home_frame, text="Profili di estrazione", padding=8)
        outer.pack(fill="both", expand=True, padx=60, pady=(0, 10))
        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.home_profiles_frame = ttk.Frame(canvas)
        self._home_canvas_window = canvas.create_window((0,0), window=self.home_profiles_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.home_profiles_frame.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(self._home_canvas_window, width=e.width))
        self.home_profiles_canvas = canvas

        self.home_status_var = tk.StringVar(value=self.status_var.get())
        ttk.Label(self.home_frame, textvariable=self.home_status_var, anchor="center", foreground="#555555").pack(fill="x", padx=60)
        self.status_var.trace_add("write", lambda *_: self.home_status_var.set(self.status_var.get()))

        # Stato dettagliato della sessione: profilo attivo + due fasi distinte.
        phase_box = ttk.Frame(self.home_frame)
        phase_box.pack(fill="x", padx=180, pady=(5, 0))
        self.home_extracting_profile_var = tk.StringVar(value="")
        ttk.Label(phase_box, textvariable=self.home_extracting_profile_var, font=("TkDefaultFont", 9, "bold"), anchor="w").grid(row=0,column=0,columnspan=3,sticky="ew",pady=(0,3))
        self.home_read_label_var = tk.StringVar(value="Lettura: 0%")
        self.home_extract_label_var = tk.StringVar(value="Estrazione: 0%")
        self.home_consolidate_label_var = tk.StringVar(value="Consolidamento: 0%")
        ttk.Label(phase_box,textvariable=self.home_read_label_var,width=30,anchor="w").grid(row=1,column=0,sticky="w")
        self.home_read_progress = ttk.Progressbar(phase_box,maximum=100,mode="determinate")
        self.home_read_progress.grid(row=1,column=1,sticky="ew",padx=(8,0))
        ttk.Label(phase_box,textvariable=self.home_extract_label_var,width=30,anchor="w").grid(row=2,column=0,sticky="w")
        self.home_extract_progress = ttk.Progressbar(phase_box,maximum=100,mode="determinate")
        self.home_extract_progress.grid(row=2,column=1,sticky="ew",padx=(8,0))
        ttk.Label(phase_box,textvariable=self.home_consolidate_label_var,width=30,anchor="w").grid(row=3,column=0,sticky="w")
        self.home_consolidate_progress = ttk.Progressbar(phase_box,maximum=100,mode="determinate")
        self.home_consolidate_progress.grid(row=3,column=1,sticky="ew",padx=(8,0))
        phase_box.columnconfigure(1,weight=1)
        self.home_stop_button = ttk.Button(
            self.home_frame, text="Interrompi estrazione", command=self.stop_extraction, state="disabled"
        )
        self.home_stop_button.pack(pady=(8, 2))

        bottom = ttk.Frame(self.home_frame, padding=(38, 8, 38, 2))
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Modifica regole / schermata completa", command=self._show_advanced_view).pack(side="left")
        if IS_PRO:
            ttk.Button(bottom, text="Crea nuova regola tramite Wizard", command=self.open_guided_wizard).pack(side="right")
        self._refresh_home_profiles()

    def _load_home_profile_icon(self, profile_path: Path, profile: dict):
        rel = str(profile.get("icon", "") or "").strip()
        if not rel:
            return None
        path = Path(rel)
        if not path.is_absolute():
            path = Path(self.config_mgr.profiles_dir) / path
        if not path.exists():
            return None
        try:
            if Image is not None and ImageTk is not None:
                img = Image.open(path).convert("RGBA")
                img.thumbnail((54,54))
                return ImageTk.PhotoImage(img)
            if path.suffix.lower() in {".png", ".gif"}:
                return tk.PhotoImage(file=str(path))
        except Exception:
            return None
        return None

    def _associated_profile_paths(self) -> list[Path]:
        if not IS_PRO:
            return [Path(self.profile_path)]
        paths = self.profile_mgr.list_profiles()
        names = self.config_mgr.associated_profile_names()
        if names is None:
            return paths
        allowed = set(names)
        return [Path(p) for p in paths if Path(p).name in allowed]

    def _general_excel_path(self, profile_path: Path | None = None, profile: dict | None = None) -> Path:
        p = Path(profile_path or self.profile_path)
        prof = profile or (self.profile if p == Path(self.profile_path) else self.profile_mgr.load(p))
        profile_name = safe_filename_component(str(prof.get("name", p.stem) or p.stem), p.stem)
        email_name = safe_filename_component(self.config_mgr.active_imap_profile_name or "Email", "Email")
        target = self.config_mgr.work_dir / f"{profile_name} - {email_name}.xlsx"
        # Migrazione trasparente del vecchio nome «profilo.xlsx» al nuovo
        # «profilo - email.xlsx», senza sovrascrivere file già esistenti.
        legacy = self.config_mgr.work_dir / f"{p.stem}.xlsx"
        if not target.exists() and legacy.exists() and legacy.is_file():
            try:
                legacy.replace(target)
            except Exception:
                pass
        return target

    @staticmethod
    def _home_last_extraction_label(raw_value: str) -> str:
        text = str(raw_value or "").strip()
        if not text:
            return "mai"
        try:
            dt = datetime.fromisoformat(text)
            target = dt.date()
            today = datetime.now().date()
            delta = (today - target).days
            if delta == 0:
                return "oggi"
            if delta == 1:
                return "ieri"
            if 1 < delta <= 30:
                return f"{delta} giorni fa"
            return f"{target.day}/{target.month}/{target.year}"
        except Exception:
            return text

    @staticmethod
    def _format_home_summary_number(value, spec: dict) -> str:
        if value in (None, ""):
            return "—"
        if isinstance(value, bool):
            value = int(value)
        if isinstance(value, (int, float)):
            if str(spec.get("format", "") or "") == "currency_eur":
                # Formato italiano leggibile nella home.
                txt = f"{float(value):,.2f}".replace(",", "§").replace(".", ",").replace("§", ".")
                return txt + " €"
            if isinstance(value, float) and not value.is_integer():
                return (f"{value:.2f}".rstrip("0").rstrip(".")).replace(".", ",")
            return str(int(value))
        return str(value)

    @staticmethod
    def _home_start_date_label(raw_value: str) -> str:
        text = str(raw_value or "").strip()
        if not text:
            return "non impostata"
        try:
            dt = datetime.fromisoformat(text).date()
            return f"{dt.day}/{dt.month}/{dt.year}"
        except Exception:
            return text

    def _home_profile_details(self, profile: dict, profile_path: Path | None = None) -> list[str]:
        pkey = Path(profile_path or self.profile_path).name
        state = self.config_mgr.extraction_state(pkey)
        lines = [
            f"Data inizio estrazione: {self._home_start_date_label(str(state.get('extraction_start_date', '') or ''))}",
            f"Ultima estrazione: {self._home_last_extraction_label(str(state.get('last_processed_at', '') or ''))}",
        ]

        selected = [str(x) for x in profile.get("home_summary_fields", []) or []][:6]
        cache = state.get("home_summary_values", []) or []
        by_field = {
            str(item.get("field", "")): item
            for item in cache if isinstance(item, dict) and item.get("field")
        }
        definitions = {
            str(f.get("name")): f
            for f in profile.get("fields", []) or []
            if isinstance(f, dict) and f.get("name")
        }
        summary_defs = {
            str(item.get("name", "")): item
            for item in profile.get("summary_formulas", []) or []
            if isinstance(item, dict) and item.get("name")
        }

        for ref in selected:
            source_kind, _, raw_name = ref.partition(":")
            if not raw_name:
                source_kind, raw_name = "field", ref
            cached = by_field.get(ref) or by_field.get(raw_name) or {}
            value = cached.get("value") if isinstance(cached, dict) else None
            fmt_spec = definitions.get(raw_name, {})
            if source_kind == "summary":
                label = raw_name
                if str(summary_defs.get(raw_name, {}).get("kind", "") or "") == "period_aggregate":
                    label = raw_name
            elif source_kind == "computed":
                label = raw_name.replace("_", " ")
            else:
                kind = str(fmt_spec.get("final_formula", "") or "").lower()
                prefix = "Somma" if kind == "sum" else "Conta" if kind == "count" else "Valore"
                label = f"{prefix} {raw_name.replace('_', ' ')}"
            lines.append(f"{label}: {self._format_home_summary_number(value, fmt_spec)}")
        return lines

    def _home_button_text(self, profile: dict, profile_path: Path) -> str:
        # Compatibilità con eventuali test/usi preesistenti: restituisce ancora
        # un testo completo, ma la home grafica usa ora titolo e dettagli separati.
        name = str(profile.get("name", profile_path.stem) or profile_path.stem)
        return "\n".join([name] + self._home_profile_details(profile, profile_path))

    def _update_active_home_summary_cache_from_excel(self, excel_path: Path) -> None:
        selected = [str(x) for x in self.profile.get("home_summary_fields", []) or []][:6]
        state = self._active_extraction_state()
        if not selected:
            state["home_summary_values"] = []
            self.config_mgr.save()
            return
        values = read_excel_home_summary_values(excel_path, self.profile, selected)
        state["home_summary_values"] = values
        state["home_summary_updated_at"] = datetime.now().isoformat(timespec="seconds")
        self.config_mgr.save()

    def _refresh_home_profile_buttons(self) -> None:
        if not hasattr(self, "home_profiles_frame"):
            return
        for child in self.home_profiles_frame.winfo_children():
            child.destroy()
        self._home_icon_refs = []
        profiles = self._associated_profile_paths()

        # Tutti i pulsanti usano la stessa griglia: icona, titolo e una colonna
        # informazioni che inizia sempre alla stessa ascissa. In questo modo i
        # valori restano allineati anche se un profilo non ha un'icona o ha un
        # nome più lungo/corto.
        card_bg = "#f7f9fc"
        active_bg = "#eef4fb"
        for path in profiles:
            try:
                prof = self.profile if Path(path).resolve() == Path(self.profile_path).resolve() else self.profile_mgr.load(path)
            except Exception:
                continue
            name = str(prof.get("name", path.stem) or path.stem)
            icon = self._load_home_profile_icon(Path(path), prof)
            if icon is not None:
                self._home_icon_refs.append(icon)

            bg = active_bg if Path(path).resolve() == Path(self.profile_path).resolve() else card_bg
            card = tk.Frame(
                self.home_profiles_frame, bg=bg, relief="raised", bd=1,
                highlightthickness=1, highlightbackground="#d7dce3", cursor="hand2",
                padx=16, pady=12,
            )
            card.pack(fill="x", padx=5, pady=6)
            card.grid_columnconfigure(0, minsize=70)
            card.grid_columnconfigure(1, minsize=390)
            card.grid_columnconfigure(2, minsize=300)
            card.grid_columnconfigure(3, weight=1)

            if icon is not None:
                icon_lbl = tk.Label(card, image=icon, bg=bg, cursor="hand2")
            else:
                # Colonna segnaposto: mantiene l'allineamento degli altri campi.
                icon_lbl = tk.Label(card, text="", width=7, bg=bg, cursor="hand2")
            icon_lbl.grid(row=0, column=0, sticky="w", padx=(0, 12))

            title = tk.Label(
                card, text=name, bg=bg, anchor="w", justify="left",
                font=("TkDefaultFont", 16, "bold"), cursor="hand2",
            )
            title.grid(row=0, column=1, sticky="w", padx=(0, 24))

            details = self._home_profile_details(prof, Path(path))
            # Data inizio, ultima estrazione e campi 1-2 restano nella prima
            # colonna informativa; i campi 3-6 sono in una seconda colonna.
            left_details = details[:4]
            right_details = details[4:8]
            details_left_lbl = tk.Label(
                card, text="\n".join(left_details), bg=bg, justify="left", anchor="w",
                font=("TkDefaultFont", 11), cursor="hand2",
            )
            details_left_lbl.grid(row=0, column=2, sticky="w", padx=(6, 24))
            details_right_lbl = tk.Label(
                card, text="\n".join(right_details), bg=bg, justify="left", anchor="w",
                font=("TkDefaultFont", 11), cursor="hand2",
            )
            details_right_lbl.grid(row=0, column=3, sticky="w", padx=(6, 8))

            run = lambda _e=None, p=Path(path): self._home_run_profile(p)
            for widget in (card, icon_lbl, title, details_left_lbl, details_right_lbl):
                widget.bind("<Button-1>", run)
                widget.bind("<Return>", run)

        if not profiles:
            ttk.Label(self.home_profiles_frame, text="Nessun profilo di estrazione disponibile.").pack(pady=20)

    def _refresh_home_profiles(self) -> None:
        if not hasattr(self, "home_imap_var"):
            return
        self.home_imap_var.set(self.config_mgr.active_imap_profile_name or "<nessuna email>")
        self._refresh_home_profile_buttons()

    def _home_change_imap(self) -> None:
        self.choose_imap_profile()
        self._refresh_home_profiles()

    def _home_run_profile(self, path: Path) -> None:
        self._load_profile_path(path)
        self.home_frame.lift()
        if hasattr(self, "home_stop_button"):
            self.home_stop_button.configure(state="normal")
        self.after(20, self.extract_incremental)

    def _show_advanced_view(self) -> None:
        if hasattr(self, "home_frame"):
            self.home_frame.place_forget()

    def show_home_screen(self) -> None:
        self._refresh_home_profiles()
        if hasattr(self, "home_frame"):
            self.home_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.home_frame.lift()


    def _append_status_message(self, message: str) -> None:
        text = str(message or "").strip()
        if not text or not hasattr(self, "status_text"):
            return
        stamp = datetime.now().strftime("%H:%M:%S")
        try:
            self.status_text.configure(state="normal")
            self.status_text.insert("end", f"[{stamp}] {text}\n")
            self.status_text.see("end")
            self.status_text.configure(state="disabled")
        except Exception:
            pass

    def _on_status_var_changed(self, *_args) -> None:
        self._append_status_message(self.status_var.get())

    def _clear_status_messages(self) -> None:
        try:
            self.status_text.configure(state="normal")
            self.status_text.delete("1.0", "end")
            self.status_text.configure(state="disabled")
        except Exception:
            pass

    def _configure_tags(self):
        self.text.tag_configure("trigger", background="#FFF59D")
        self.text.tag_configure("text", background="#C8E6C9")
        self.text.tag_configure("number", background="#BBDEFB")
        self.text.tag_configure("money", background="#B3E5FC", underline=True)
        self.text.tag_configure("date", background="#F8BBD0")
        self.text.tag_configure("key", background="#FFE0B2", underline=True)

    def _refresh_mail_types(self):
        names = [mt.get("name", "") for mt in self.profile.get("mail_types", [])]
        self.mail_type_combo["values"] = names
        if names and self.mail_type_var.get() not in names:
            self.mail_type_var.set(names[0])
        if not names:
            self.mail_type_var.set("")
        self.filter_mail_types.intersection_update(set(names))
        self._configure_tree_columns()
        self._update_type_filter_label()
        self._recompute_cached_matches()

    def _configure_tree_columns(self):
        base = ("sel", "uid", "date", "sender", "subject")
        self.mail_type_columns = [
            (f"mailtype_{i}", str(mt.get("name", "")))
            for i, mt in enumerate(self.profile.get("mail_types", []))
        ]
        cols = base + tuple(col_id for col_id, _name in self.mail_type_columns)
        self.tree.configure(columns=cols)
        for col, text, width, stretch in [
            ("sel", "Sel.", 45, False),
            ("uid", "UID", 70, False),
            ("date", "Data", 190, False),
            ("sender", "Mittente", 300, True),
            ("subject", "Oggetto", 560, True),
        ]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=stretch, anchor="w")
        for col_id, name in self.mail_type_columns:
            label = f"✓ {name}" if name in self.filter_mail_types else name
            self.tree.heading(
                col_id, text=label,
                command=lambda n=name: self.filter_by_mail_type(n)
            )
            self.tree.column(col_id, width=max(105, min(190, len(name) * 9 + 35)), stretch=False, anchor="center")

    def _tree_values(self, uid: int):
        h = self.headers.get(uid) or self.messages.get(uid)
        if h is None:
            return ()
        matches = self.row_matches.get(uid, set())
        mark = "☑" if uid in self.selected_uids else "☐"
        dynamic = tuple("✓" if name in matches else "" for _col, name in self.mail_type_columns)
        return (mark, uid, h.date, h.sender, h.subject) + dynamic

    def _row_visible(self, uid: int) -> bool:
        return matches_active_type_filters(self.row_matches.get(uid, set()), self.filter_mail_types)

    def _rebuild_tree_rows(self):
        selected_view = set(self.tree.selection())
        for item in self.tree.get_children():
            self.tree.delete(item)
        for uid in self.search_order:
            if not self._row_visible(uid):
                continue
            values = self._tree_values(uid)
            if values:
                self.tree.insert("", "end", iid=str(uid), values=values)
        keep = [iid for iid in selected_view if self.tree.exists(iid)]
        if keep:
            self.tree.selection_set(keep)

    def _recompute_cached_matches(self):
        if not hasattr(self, "tree"):
            return
        missing_body = False
        for uid in self.search_order:
            source = self.messages.get(uid) or self.headers.get(uid)
            if source is None:
                continue
            names: set[str] = set()
            for mt in self.profile.get("mail_types", []):
                body_needed = any(
                    isinstance(rule, dict)
                    and str(rule.get("trigger", "") or "").strip()
                    and rule.get("source") == "body"
                    for rule in mt.get("match_rules", [])
                )
                if body_needed and uid not in self.messages:
                    missing_body = True
                    continue
                if matches_mail_type(source, mt):
                    names.add(str(mt.get("name", "")))
            self.row_matches[uid] = names
        self._configure_tree_columns()
        self._rebuild_tree_rows()
        if missing_body:
            self.status_var.set(
                "Alcuni match sul corpo richiedono una nuova ricerca: premi Cerca per ricalcolare tutte le spunte."
            )

    def _active_filter_names(self) -> list[str]:
        """Tipi filtro attivi nell'ordine in cui compaiono nel profilo."""
        ordered = [name for _col, name in self.mail_type_columns]
        return [name for name in ordered if name in self.filter_mail_types]

    def _update_type_filter_label(self) -> None:
        names = self._active_filter_names() if self.mail_type_columns else sorted(self.filter_mail_types)
        if not names:
            self.type_filter_var.set("Filtro tipi mail: tutti")
        else:
            self.type_filter_var.set("Filtro tipi mail (OR): " + " | ".join(names))

    def filter_by_mail_type(self, name: str | None):
        # Ogni clic sull'intestazione aggiunge/rimuove quel tipo dal filtro.
        # Con più colonne selezionate la riga resta visibile se matcha ALMENO
        # UNO dei tipi scelti (OR). "Mostra tutti" azzera l'intero filtro.
        if name is None:
            self.filter_mail_types.clear()
        elif name in self.filter_mail_types:
            self.filter_mail_types.remove(name)
        else:
            self.filter_mail_types.add(name)
        self._update_type_filter_label()
        self._configure_tree_columns()
        self._rebuild_tree_rows()
        shown = len(self.tree.get_children())
        names = self._active_filter_names()
        desc = "tutti" if not names else " OR ".join(names)
        self.status_var.set(f"Filtro {desc} — {shown} mail visualizzate")

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "search_item":
                    item, match_names = payload
                    self.headers[item.uid] = item
                    if isinstance(item, MailMessage):
                        self.messages[item.uid] = item
                    self.row_matches[item.uid] = set(match_names)
                    if item.uid not in self.search_order:
                        self.search_order.append(item.uid)
                    if self._row_visible(item.uid):
                        self.tree.insert("", "end", iid=str(item.uid), values=self._tree_values(item.uid))
                elif kind == "progress":
                    i, total = payload
                    self.status_var.set(f"Ricerca: {i}/{total} — trovate {len(self.headers)}")
                elif kind == "extraction_phase_progress":
                    if len(payload) >= 5:
                        phase, current, total, profile_name, phase_extra = payload[:5]
                    else:
                        phase, current, total, profile_name = payload
                        phase_extra = ""
                    total=max(int(total or 0),1); current=max(0,min(int(current or 0),total))
                    pct=int(round(current*100/total))
                    if hasattr(self,"home_extracting_profile_var"):
                        self.home_extracting_profile_var.set(f"Profilo in estrazione: {profile_name}")
                    if phase=="read" and hasattr(self,"home_read_progress"):
                        self.home_read_progress["value"]=pct
                        self.home_read_label_var.set(f"Lettura: {pct}%  ({current}/{total})" + (f" — {phase_extra}" if phase_extra else ""))
                    elif phase=="extract" and hasattr(self,"home_extract_progress"):
                        self.home_extract_progress["value"]=pct
                        self.home_extract_label_var.set(f"Estrazione: {pct}%  ({current}/{total})")
                    elif phase=="consolidate" and hasattr(self,"home_consolidate_progress"):
                        self.home_consolidate_progress["value"]=pct
                        self.home_consolidate_label_var.set(f"Consolidamento: {pct}%  ({current}/{total})")
                elif kind == "done":
                    self.status_var.set(payload)
                elif kind == "error":
                    self.status_var.set("Errore")
                    messagebox.showerror("Errore", str(payload), parent=self)
                elif kind == "extraction_error":
                    self.extraction_running = False
                    if hasattr(self, "home_stop_button"):
                        self.home_stop_button.configure(state="disabled")
                    self.status_var.set("Errore estrazione")
                    messagebox.showerror("Errore estrazione", str(payload), parent=self)
                elif kind == "extraction_retry_prompt":
                    error, context, answer_event, answer_box = payload
                    self.status_var.set("Connessione email interrotta")
                    retry = messagebox.askyesno(
                        "Connessione email interrotta",
                        f"La connessione email si è interrotta durante {context}.\n\n"
                        f"Errore: {error}\n\n"
                        "EgoMailExtractor ha già tentato automaticamente di riconnettersi. "
                        "Vuoi riprovare e continuare dal punto raggiunto?",
                        parent=self,
                    )
                    answer_box["retry"] = bool(retry)
                    if not retry:
                        self.extraction_stop_event.set()
                    answer_event.set()
                elif kind == "message":
                    self._display_message(payload)
                elif kind == "extraction_done":
                    # payload: {message, path}. La domanda di apertura viene fatta
                    # sul thread GUI, mai dal worker.
                    self.extraction_running = False
                    if hasattr(self, "home_stop_button"):
                        self.home_stop_button.configure(state="disabled")
                    message = payload.get("message", "Estrazione completata")
                    path = payload.get("path", "")
                    self.status_var.set("Pronto")
                    if path:
                        open_now = messagebox.askyesno(
                            "Estrazione completata",
                            f"{message}\n\nVuoi aprire il file Excel?",
                            parent=self,
                        )
                        if open_now:
                            try:
                                self._open_file_cross_platform(path)
                            except Exception as exc:
                                messagebox.showerror(
                                    "Apertura Excel",
                                    f"Impossibile aprire il file:\n{path}\n\n{exc}",
                                    parent=self,
                                )
                    else:
                        messagebox.showinfo("Estrazione completata", message, parent=self)
                elif kind == "extraction_stopped":
                    self.extraction_running = False
                    if hasattr(self, "home_stop_button"):
                        self.home_stop_button.configure(state="disabled")
                    self.status_var.set("Estrazione interrotta")
                    message = str(payload)
                    excel_path = self._last_excel_path_for_profile()
                    if excel_path and excel_path.exists():
                        open_now = messagebox.askyesno(
                            "Estrazione interrotta",
                            message + "\n\nVuoi aprire il file Excel generale aggiornato fino a questo punto?",
                            parent=self,
                        )
                        if open_now:
                            try:
                                self._open_file_cross_platform(excel_path)
                            except Exception as exc:
                                messagebox.showerror("Apertura Excel", f"Impossibile aprire il file:\n{excel_path}\n\n{exc}", parent=self)
                    else:
                        messagebox.showinfo("Estrazione interrotta", message, parent=self)
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def start_search(self):
        if not self.config_mgr.imap.host or not self.config_mgr.imap.username or not self.config_mgr.get_password():
            dlg = ImapSettingsDialog(self, self.config_mgr)
            self.wait_window(dlg)
            if not self.config_mgr.get_password():
                return
        self.stop_event.clear()
        self.headers.clear()
        self.messages.clear()
        self.row_matches.clear()
        self.search_order.clear()
        self.selected_uids.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("Ricerca in corso…")
        needs_body = any(
            str(rule.get("trigger", "") or "").strip() and rule.get("source") == "body"
            for mt in self.profile.get("mail_types", [])
            for rule in mt.get("match_rules", [])
            if isinstance(rule, dict)
        )

        def worker():
            try:
                found = [0]
                def push_item(item):
                    found[0] += 1
                    names = [str(mt.get("name", "")) for mt in matching_mail_types(item, self.profile)]
                    self.ui_queue.put(("search_item", (item, names)))
                self.imap.search_progressive(
                    self.folder_var.get().strip() or "INBOX",
                    self.subject_var.get(), self.sender_var.get(), self.stop_event,
                    push_item,
                    lambda i, t: self.ui_queue.put(("progress", (i, t))),
                    include_body=needs_body,
                )
                msg = "Ricerca interrotta" if self.stop_event.is_set() else f"Ricerca completata: {found[0]} mail"
                self.ui_queue.put(("done", msg))
            except Exception as e:
                self.ui_queue.put(("error", e))
        threading.Thread(target=worker, daemon=True).start()

    def stop_search(self):
        self.stop_event.set()

    def on_tree_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        uid = int(sel[-1])
        self.current_uid = uid
        if uid in self.messages:
            self._display_message(self.messages[uid])
            return
        self.status_var.set(f"Lettura UID {uid}…")
        def worker():
            try:
                msg = self.imap.fetch_message(uid, self.folder_var.get().strip() or "INBOX")
                self.messages[uid] = msg
                self.ui_queue.put(("message", msg))
            except Exception as e:
                self.ui_queue.put(("error", e))
        threading.Thread(target=worker, daemon=True).start()

    def _display_message(self, msg: MailMessage):
        if self.current_uid is not None and msg.uid != self.current_uid:
            return
        self.current_mail = msg
        self.messages[msg.uid] = msg
        self.row_matches[msg.uid] = {str(mt.get("name", "")) for mt in matching_mail_types(msg, self.profile)}
        if self.tree.exists(str(msg.uid)):
            self.tree.item(str(msg.uid), values=self._tree_values(msg.uid))

        self.text.delete("1.0", "end")
        # Costruiamo il testo tenendo anche l'offset esatto delle tre possibili
        # sorgenti delle regole: serve per applicare i tag grafici in anteprima.
        parts: list[str] = []
        cursor = 0
        self.preview_source_offsets = {}

        def add(text: str):
            nonlocal cursor
            parts.append(text)
            cursor += len(text)

        add("Da: ")
        self.preview_source_offsets["sender"] = cursor
        add(msg.sender or "")
        add("\nDestinatario: ")
        self.preview_source_offsets["recipient"] = cursor
        add(getattr(msg, "recipient", "") or "")
        add("\nOggetto: ")
        self.preview_source_offsets["subject"] = cursor
        add(msg.subject or "")
        add("\nData mail: ")
        self.preview_source_offsets["date"] = cursor
        add(msg.date or "")
        add("\nID mail: ")
        self.preview_source_offsets["message_id"] = cursor
        add(msg.message_id or "")
        add(f"\nUID: {msg.uid}\n\n")
        self.preview_source_offsets["body"] = cursor
        add(msg.body or "")
        self.text.insert("1.0", "".join(parts))

        mt = self._match_type_name(msg)
        if mt:
            self.mail_type_var.set(mt)
        self._update_extraction_preview()
        self.status_var.set(f"UID {msg.uid} caricato")

    def _clear_rule_highlights(self):
        self.hover_tip.hide()
        for tag in tuple(self.text.tag_names()):
            if str(tag).startswith("extract_rule_"):
                try:
                    self.text.tag_delete(tag)
                except tk.TclError:
                    pass

    def _tag_source_span(self, source_name: str, span: tuple[int, int] | None, tag: str):
        if not span or not self.current_mail:
            return
        base = self.preview_source_offsets.get(source_name)
        if base is None:
            return
        start, end = span
        if start < 0 or end <= start:
            return
        try:
            a = self.text.index(f"1.0 + {base + start} chars")
            b = self.text.index(f"1.0 + {base + end} chars")
            self.text.tag_add(tag, a, b)
        except tk.TclError:
            pass

    def _bind_rule_tooltip(self, tag: str, field: str, value):
        label = f"Campo: {field}"
        if value not in (None, ""):
            label += f"\nValore memorizzato: {value}"
        self.text.tag_bind(
            tag, "<Enter>",
            lambda event, txt=label: self.hover_tip.show(txt, event.x_root, event.y_root)
        )
        self.text.tag_bind(tag, "<Leave>", lambda _event: self.hover_tip.hide())

    def _highlight_extraction_rules(self, mt: dict):
        """Colora trigger/match e sottolinea il valore estratto per ogni regola."""
        self._clear_rule_highlights()
        if not self.current_mail:
            return
        for i, rule in enumerate(mt.get("rules", [])):
            trace = extract_rule_trace(self.current_mail, rule)
            if trace.error or not trace.matched:
                continue
            color = RULE_COLORS[i % len(RULE_COLORS)]
            field = str(rule.get("field", ""))
            trigger_tag = f"extract_rule_trigger_{i}"
            value_tag = f"extract_rule_value_{i}"
            # Il trigger (o, per regex, l'intero match) è evidenziato col colore
            # assegnato alla regola. Il testo realmente prelevato è anche
            # sottolineato, così rimane distinguibile se ricade nello stesso span.
            self.text.tag_configure(trigger_tag, background=color)
            self.text.tag_configure(value_tag, background=color, underline=True)
            self._tag_source_span(trace.source_name, trace.trigger_span, trigger_tag)
            self._tag_source_span(trace.source_name, trace.value_span, value_tag)
            self._bind_rule_tooltip(trigger_tag, field, trace.value)
            self._bind_rule_tooltip(value_tag, field, trace.value)
            try:
                self.text.tag_raise(value_tag)
            except tk.TclError:
                pass

    def _update_extraction_preview(self):
        self._clear_rule_highlights()
        if not self.current_mail:
            self.extraction_preview_var.set("Anteprima estrazione: apri una mail")
            return
        try:
            mt = self._active_mail_type()
            if mt is None:
                self.extraction_preview_var.set("Anteprima estrazione: seleziona un tipo mail")
                return
            mt_name = mt.get("name", "")
            is_match = matches_mail_type(self.current_mail, mt)
            if not is_match:
                self.extraction_preview_var.set(
                    f"Anteprima estrazione [{mt_name}]: NON MATCH — le regole di estrazione non vengono applicate"
                )
                return
            record = {}
            for rule in mt.get("rules", []):
                value = extract_rule(self.current_mail, rule)
                if value not in (None, ""):
                    record[rule.get("field")] = value
            compute_fields(record, self.profile)
            self._highlight_extraction_rules(mt)
            if not record:
                self.extraction_preview_var.set(f"Anteprima estrazione [{mt_name}]: MATCH — nessun campo estratto")
                return
            pieces = [f"{k}={v}" for k, v in record.items()]
            text = f"Anteprima estrazione [{mt_name}]: MATCH — " + " | ".join(pieces)
            self.extraction_preview_var.set(text[:500] + ("…" if len(text) > 500 else ""))
        except Exception as exc:
            self.extraction_preview_var.set(f"Anteprima estrazione: errore {exc}")

    def _match_type_name(self, msg: MailMessage) -> str | None:
        matches = matching_mail_types(msg, self.profile)
        return matches[0].get("name") if matches else None

    def toggle_selected(self, event=None):
        item = self.tree.identify_row(event.y) if event else None
        if not item:
            sel = self.tree.selection()
            item = sel[-1] if sel else None
        if not item:
            return
        uid = int(item)
        if uid in self.selected_uids:
            self.selected_uids.remove(uid)
            mark = "☐"
        else:
            self.selected_uids.add(uid)
            mark = "☑"
        vals = list(self.tree.item(item, "values"))
        vals[0] = mark
        self.tree.item(item, values=vals)

    def select_tree_rows(self):
        for item in self.tree.selection():
            uid = int(item)
            self.selected_uids.add(uid)
            vals = list(self.tree.item(item, "values")); vals[0] = "☑"; self.tree.item(item, values=vals)

    def clear_selected(self):
        self.selected_uids.clear()
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values")); vals[0] = "☐"; self.tree.item(item, values=vals)

    def _selected_text(self) -> tuple[str, str, str] | None:
        try:
            start = self.text.index("sel.first")
            end = self.text.index("sel.last")
            return self.text.get(start, end).strip(), start, end
        except tk.TclError:
            messagebox.showwarning("Selezione richiesta", "Seleziona prima il testo nel contenuto della mail.", parent=self)
            return None

    def _active_mail_type(self) -> dict | None:
        name = self.mail_type_var.get()
        return next((x for x in self.profile.get("mail_types", []) if x.get("name") == name), None)

    def mark_selection(self, kind: str):
        if not self.current_mail:
            messagebox.showwarning("Mail", "Apri prima una mail.", parent=self)
            return
        data = self._selected_text()
        if not data:
            return
        selected, start, end = data
        if kind == "trigger":
            self.last_trigger = selected
            self.text.tag_add("trigger", start, end)
            self.status_var.set(f"Trigger attivo: {selected[:60]}")
            return
        mt = self._active_mail_type()
        if not mt:
            if not self.create_mail_type():
                return
            mt = self._active_mail_type()
        proposed = propose_field_name(self.last_trigger, selected, kind)
        field = simpledialog.askstring("Nome campo", "Nome del campo da popolare:", initialvalue=proposed, parent=self)
        if not field:
            return
        field = slugify(field, proposed)
        rule = infer_rule_from_selection(self.current_mail.body, self.last_trigger, selected, field, kind)
        field_type = {"number": "number", "money": "number", "date": "date"}.get(kind, "text")
        field_existed = get_field(self.profile, field) is not None
        ensure_field(
            self.profile, field, field_type,
            display_format="currency_eur" if kind == "money" else None,
        )
        key_was_present = field in self.profile.setdefault("reconcile_keys", [])
        if kind == "key" and field not in self.profile.setdefault("reconcile_keys", []):
            self.profile["reconcile_keys"].append(field)

        # Prima di aggiungerla al profilo, mostra l'editor completo: l'utente
        # può correggere nome, campo, tipo Excel, regex/trigger e conversione.
        dlg = RuleEditorDialog(self, self.profile, rule, title="Nuova regola di estrazione", current_mail=self.current_mail)
        self.wait_window(dlg)
        if not dlg.result:
            if not field_existed:
                self.profile["fields"] = [f for f in self.profile.get("fields", []) if not (isinstance(f, dict) and f.get("name") == field)]
            if not key_was_present:
                self.profile["reconcile_keys"] = [x for x in self.profile.get("reconcile_keys", []) if x != field]
            return
        rule = dlg.result
        mt.setdefault("rules", []).append(rule)
        self.text.tag_add(kind, start, end)
        self.status_var.set(f"Creata regola '{rule['name']}' per {rule.get('field', field)}")
        self._update_extraction_preview()

    def create_mail_type(self):
        if not self.current_mail:
            messagebox.showwarning("Tipo mail", "Apri prima una mail da usare come modello.", parent=self)
            return False
        subj = self.current_mail.subject
        proposed = slugify(re.sub(r"\d+", "", subj)[:55], "tipo_mail")
        name = simpledialog.askstring("Nuovo tipo mail", "Nome del tipo di mail:", initialvalue=proposed, parent=self)
        if not name:
            return False
        mt = {
            "name": slugify(name, "tipo_mail"),
            "description": "Creato dalla GUI",
            "match_rules": [
                {
                    "name": "match_oggetto",
                    "source": "subject",
                    "mode": "contains",
                    "trigger": subj,
                }
            ],
            "rules": []
        }
        self.profile.setdefault("mail_types", []).append(mt)
        self._refresh_mail_types()
        self.mail_type_var.set(mt["name"])
        self.edit_mail_type_criteria()
        return True

    def rename_current_mail_type(self):
        mt = self._active_mail_type()
        if not mt:
            messagebox.showwarning("Tipo mail", "Seleziona prima un tipo mail.", parent=self)
            return
        old_name = str(mt.get("name", ""))
        value = simpledialog.askstring(
            "Rinomina tipo mail", "Nuovo nome del tipo mail:",
            initialvalue=old_name, parent=self,
        )
        if not value:
            return
        try:
            new_name = rename_mail_type(self.profile, old_name, value)
        except Exception as exc:
            messagebox.showerror("Rinomina tipo mail", str(exc), parent=self)
            return
        if old_name in self.filter_mail_types:
            self.filter_mail_types.remove(old_name)
            self.filter_mail_types.add(new_name)
        self.mail_type_var.set(new_name)
        self._refresh_mail_types()
        self.mail_type_var.set(new_name)
        self._update_extraction_preview()
        self.status_var.set(f"Tipo mail rinominato: {old_name} → {new_name}")

    def delete_current_mail_type(self):
        mt = self._active_mail_type()
        if not mt:
            messagebox.showwarning("Tipo mail", "Seleziona prima un tipo mail.", parent=self)
            return
        name = str(mt.get("name", ""))
        n_match = len(mt.get("match_rules", []))
        n_extract = len(mt.get("rules", []))
        if not messagebox.askyesno(
            "Elimina tipo mail",
            f"Eliminare il tipo mail '{name}'?\n\n"
            f"Verranno eliminate anche {n_match} regole di selezione e {n_extract} regole di estrazione.",
            parent=self,
        ):
            return
        if not delete_mail_type(self.profile, name):
            return
        self.filter_mail_types.discard(name)
        self.mail_type_var.set("")
        self._refresh_mail_types()
        values = list(self.mail_type_combo["values"])
        if values:
            self.mail_type_var.set(values[0])
        self._update_extraction_preview()
        self.status_var.set(f"Tipo mail eliminato: {name}")

    def _on_match_rules_changed(self):
        # Le spunte nelle colonne dei tipi mail dipendono dai match: quando
        # l'utente modifica una regola, ricalcoliamo subito tutte le mail già
        # caricate dalla ricerca senza dover interrogare nuovamente IMAP.
        self._recompute_cached_matches()
        self._update_extraction_preview()

    def edit_mail_type_criteria(self):
        mt = self._active_mail_type()
        if not mt:
            messagebox.showwarning("Selezione mail", "Seleziona prima un tipo mail.", parent=self)
            return
        dlg = MatchRulesDialog(self, mt, on_change=self._on_match_rules_changed)
        self.wait_window(dlg)
        self._on_match_rules_changed()

    def manage_rules(self):
        mt = self._active_mail_type()
        if not mt:
            messagebox.showwarning("Regole", "Seleziona prima un tipo mail.", parent=self)
            return
        win = tk.Toplevel(self)
        win.title(f"Regole di estrazione — {mt.get('name', '')}")
        win.geometry("1020x470")
        frm = ttk.Frame(win, padding=10); frm.pack(fill="both", expand=True)
        cols = ("name", "field", "type", "reconcile", "policy", "source", "strategy", "transform", "trigger")
        tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="extended")
        for c, lab, w in [
            ("name","Regola",155),("field","Campo",130),("type","Tipo Excel",80),
            ("reconcile","Riconciliazione",100),("policy","Se già presente",110),
            ("source","Origine",70),("strategy","Strategia",95),("transform","Conversione",105),
            ("trigger","Trigger / pattern",240)
        ]:
            tree.heading(c, text=lab); tree.column(c, width=w, stretch=True)
        tree.pack(fill="both", expand=True)

        def refresh():
            for item in tree.get_children(): tree.delete(item)
            for i, rule in enumerate(mt.get("rules", [])):
                detail = rule.get("trigger") or rule.get("pattern") or ""
                spec = get_field(self.profile, rule.get("field", "")) or {"type": "text"}
                typ = FIELD_TYPE_LABELS[normalize_field_type(spec.get("type"))]
                is_key = "Sì" if rule.get("field", "") in self.profile.get("reconcile_keys", []) else ""
                tree.insert("", "end", iid=str(i), values=(
                    rule.get("name",""), rule.get("field",""), typ, is_key,
                    UPDATE_POLICY_LABELS.get(rule.get("existing_value_policy", "replace"), "Sostituisci"),
                    rule.get("source","body"), rule.get("strategy",""), rule.get("transform",""), detail
                ))
        def edit_rule():
            sel = tree.selection()
            if not sel: return
            i = int(sel[0])
            dlg = RuleEditorDialog(win, self.profile, mt.get("rules", [])[i], current_mail=self.current_mail)
            win.wait_window(dlg)
            if dlg.result:
                mt["rules"][i] = dlg.result
                refresh()
                self._update_extraction_preview()
        def add_rule():
            default_field = self.profile.get("fields", [{}])[0].get("name", "campo") if self.profile.get("fields") else "campo"
            initial = {
                "name": "nuova_regola",
                "field": default_field,
                "source": "body",
                "strategy": "regex",
                "pattern": "(.+)",
                "group": 1,
                "transform": "text",
                "existing_value_policy": "replace",
            }
            dlg = RuleEditorDialog(win, self.profile, initial, title="Nuova regola", current_mail=self.current_mail)
            win.wait_window(dlg)
            if dlg.result:
                mt.setdefault("rules", []).append(dlg.result)
                refresh()
                self._update_extraction_preview()
        def delete_rule():
            sel = tree.selection()
            if not sel: return
            i = int(sel[0])
            rule = mt.get("rules", [])[i]
            if messagebox.askyesno("Elimina regola", f"Eliminare la regola '{rule.get('name','')}'?", parent=win):
                del mt["rules"][i]
                refresh()
                self._update_extraction_preview()
        def copy_rules_to_target():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(
                    "Copia regole",
                    "Seleziona una o più regole di estrazione da copiare.",
                    parent=win,
                )
                return
            selected_indexes = sorted(int(item) for item in sel)

            # Costruiamo l'elenco dei profili disponibili. Il profilo corrente
            # usa l'oggetto in memoria, così vengono considerate anche eventuali
            # modifiche non ancora salvate dall'utente.
            profile_choices: list[tuple[str, Path, dict | None]] = []
            for path in self.profile_mgr.list_profiles():
                try:
                    if Path(path).resolve() == Path(self.profile_path).resolve():
                        prof = self.profile
                    else:
                        prof = self.profile_mgr.load(path)
                    label = f"{prof.get('name', path.stem)}  [{path.name}]"
                    profile_choices.append((label, Path(path), prof))
                except Exception:
                    continue
            if not any(path.resolve() == Path(self.profile_path).resolve() for _, path, _ in profile_choices):
                label = f"{self.profile.get('name', self.profile_path.stem)}  [{self.profile_path.name}]"
                profile_choices.insert(0, (label, Path(self.profile_path), self.profile))

            if not profile_choices:
                messagebox.showerror("Copia regole", "Nessun profilo disponibile.", parent=win)
                return

            dlg = tk.Toplevel(win)
            dlg.title(f"Copia {len(selected_indexes)} regola/e di estrazione")
            dlg.geometry("650x285")
            dlg.transient(win)
            dlg.grab_set()
            center_dialog_later(dlg, win)
            body = ttk.Frame(dlg, padding=12)
            body.pack(fill="both", expand=True)

            ttk.Label(
                body,
                text=(
                    "Scegli il profilo e il Tipo mail di destinazione.\n"
                    "Le regole vengono copiate, non spostate. Se manca un campo usato dalla regola, "
                    "viene copiata anche la sua definizione e l'eventuale stato di riconciliazione."
                ),
                wraplength=610,
                justify="left",
            ).pack(fill="x", pady=(0, 12))

            grid = ttk.Frame(body)
            grid.pack(fill="x")
            grid.columnconfigure(1, weight=1)
            ttk.Label(grid, text="Profilo destinazione:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
            profile_var = tk.StringVar()
            profile_combo = ttk.Combobox(
                grid, textvariable=profile_var, state="readonly",
                values=[item[0] for item in profile_choices],
            )
            profile_combo.grid(row=0, column=1, sticky="ew", pady=5)

            ttk.Label(grid, text="Tipo mail destinazione:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
            mail_type_var = tk.StringVar()
            mail_type_combo = ttk.Combobox(grid, textvariable=mail_type_var, state="readonly")
            mail_type_combo.grid(row=1, column=1, sticky="ew", pady=5)

            summary_var = tk.StringVar(value="")
            ttk.Label(body, textvariable=summary_var, foreground="#555555", wraplength=610).pack(fill="x", pady=(10, 0))

            current_destination: dict[str, object] = {}

            def update_mail_types(_event=None):
                label = profile_var.get()
                choice = next((item for item in profile_choices if item[0] == label), None)
                if choice is None:
                    mail_type_combo["values"] = ()
                    mail_type_var.set("")
                    return
                _, path, prof = choice
                if prof is None:
                    try:
                        prof = self.profile_mgr.load(path)
                    except Exception as exc:
                        messagebox.showerror("Copia regole", f"Impossibile leggere il profilo:\n{exc}", parent=dlg)
                        return
                current_destination.clear()
                current_destination.update({"path": path, "profile": prof})
                types = [
                    str(item.get("name", "")) for item in prof.get("mail_types", [])
                    if isinstance(item, dict) and item.get("name")
                ]
                mail_type_combo["values"] = types
                preferred = mt.get("name", "") if mt.get("name", "") in types else (types[0] if types else "")
                mail_type_var.set(preferred)
                if types:
                    summary_var.set(f"{len(selected_indexes)} regola/e verranno copiate in '{preferred}'.")
                else:
                    summary_var.set("Il profilo scelto non contiene Tipi mail.")

            def do_copy():
                dest_path = current_destination.get("path")
                dest_profile = current_destination.get("profile")
                dest_mail_type = mail_type_var.get().strip()
                if not isinstance(dest_path, Path) or not isinstance(dest_profile, dict) or not dest_mail_type:
                    messagebox.showwarning("Copia regole", "Seleziona un profilo e un Tipo mail di destinazione.", parent=dlg)
                    return
                try:
                    result = copy_extraction_rules(
                        self.profile,
                        str(mt.get("name", "")),
                        selected_indexes,
                        dest_profile,
                        dest_mail_type,
                    )
                    self.profile_mgr.save(dest_path, dest_profile)
                except Exception as exc:
                    messagebox.showerror("Copia regole", f"Copia non riuscita:\n{exc}", parent=dlg)
                    return

                # Se la destinazione è il profilo corrente, l'oggetto è già
                # stato aggiornato in memoria; aggiorniamo anche l'interfaccia.
                same_profile = Path(dest_path).resolve() == Path(self.profile_path).resolve()
                if same_profile:
                    self.profile = dest_profile
                    if dest_mail_type == mt.get("name"):
                        refresh()
                    self._refresh_mail_types()
                    self._update_extraction_preview()

                copied = ", ".join(result.get("copied_names", []))
                extra = []
                if result.get("added_fields"):
                    extra.append("Campi creati: " + ", ".join(result["added_fields"]))
                if result.get("added_reconcile_keys"):
                    extra.append("Chiavi di riconciliazione preservate: " + ", ".join(result["added_reconcile_keys"]))
                detail = "\n".join(extra)
                messagebox.showinfo(
                    "Copia completata",
                    f"Copiate {result.get('copied_count', 0)} regole in:\n"
                    f"{dest_profile.get('name', dest_path.stem)} → {dest_mail_type}\n\n"
                    f"Regole: {copied}" + (f"\n\n{detail}" if detail else ""),
                    parent=dlg,
                )
                self.status_var.set(
                    f"Copiate {result.get('copied_count', 0)} regole in {dest_profile.get('name', dest_path.stem)} / {dest_mail_type}"
                )
                dlg.destroy()

            profile_combo.bind("<<ComboboxSelected>>", update_mail_types)
            # Preseleziona il profilo corrente.
            current_label = next((
                label for label, path, _ in profile_choices
                if path.resolve() == Path(self.profile_path).resolve()
            ), profile_choices[0][0])
            profile_var.set(current_label)
            update_mail_types()

            buttons = ttk.Frame(body)
            buttons.pack(fill="x", side="bottom", pady=(14, 0))
            ttk.Button(buttons, text="Copia", command=do_copy).pack(side="right")
            ttk.Button(buttons, text="Annulla", command=dlg.destroy).pack(side="right", padx=(0, 6))
            dlg.wait_window()

        def show_json():
            sel = tree.selection()
            if not sel: return
            if len(sel) > 1:
                messagebox.showinfo("Dettaglio regola", "Seleziona una sola regola per visualizzarne il JSON.", parent=win)
                return
            rule = mt.get("rules", [])[int(sel[0])]
            messagebox.showinfo("Dettaglio regola", json.dumps(rule, indent=2, ensure_ascii=False), parent=win)
        btns = ttk.Frame(frm); btns.pack(fill="x", pady=(8,0))
        ttk.Button(btns, text="Nuova regola", command=add_rule).pack(side="left")
        ttk.Button(btns, text="Modifica", command=edit_rule).pack(side="left", padx=5)
        ttk.Button(btns, text="Copia selezionate…", command=copy_rules_to_target).pack(side="left", padx=5)
        ttk.Button(btns, text="Dettaglio JSON", command=show_json).pack(side="left", padx=5)
        ttk.Button(btns, text="Elimina", command=delete_rule).pack(side="left", padx=5)
        ttk.Button(btns, text="Chiudi", command=win.destroy).pack(side="right")
        tree.bind("<Double-1>", lambda _e: edit_rule())
        refresh()
        win.transient(self); win.grab_set(); center_dialog_later(win, self)

    def manage_fields(self):
        win = tk.Toplevel(self)
        win.title("Campi del profilo e tipi Excel")
        win.geometry("1320x520")
        frm = ttk.Frame(win, padding=10); frm.pack(fill="both", expand=True)
        cols = ("order", "name", "type", "format", "mail_counter", "formula", "computed", "key")
        tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")
        for c, lab, w in [
            ("order", "Ordine Excel", 85), ("name", "Campo", 220), ("type", "Tipo cella Excel", 115),
            ("format", "Formato", 90), ("mail_counter", "Conta Tipo mail", 170),
            ("formula", "Formula finale", 105), ("computed", "Calcolo automatico", 260),
            ("key", "Riconciliazione", 110)
        ]:
            tree.heading(c, text=lab); tree.column(c, width=w, stretch=True)
        tree.pack(fill="both", expand=True)

        def refresh():
            for item in tree.get_children(): tree.delete(item)
            for i, field in enumerate(self.profile.get("fields", [])):
                if not isinstance(field, dict):
                    continue
                typ = FIELD_TYPE_LABELS[normalize_field_type(field.get("type"))]
                fmt = "Valuta €" if field.get("format") == "currency_eur" else ""
                mail_counter = str(field.get("mail_type_counter", "") or "")
                formula = FINAL_FORMULA_LABELS.get(str(field.get("final_formula", "") or "").lower(), "")
                computed = ""
                for comp in self.profile.get("computed_fields", []) or []:
                    if isinstance(comp, dict) and comp.get("field") == field.get("name"):
                        if comp.get("kind") == "days_between":
                            priority_short = "regola prioritaria" if str(comp.get("priority", "rule_wins")) != "always" else "ricalcola sempre"
                            computed = f"Giorni: {comp.get('start_field', '')} → {comp.get('end_field', '')} ({priority_short})"
                        else:
                            computed = str(comp.get("kind", ""))
                        break
                if not computed and str(field.get("final_formula", "") or "").lower() in {"sum", "count"}:
                    kind = str(field.get("final_formula", "") or "").lower()
                    source = str(field.get("summary_source_field", "") or field.get("name", ""))
                    computed = f"Riepilogo: {'SOMMA' if kind == 'sum' else 'CONTA'} di {source}"
                key = "Sì" if field.get("name") in self.profile.get("reconcile_keys", []) else ""
                tree.insert("", "end", iid=str(i), values=(
                    i + 1, field.get("name", ""), typ, fmt, mail_counter, formula, computed, key
                ))

        def edit_field():
            sel = tree.selection()
            if not sel: return
            i = int(sel[0])
            old = self.profile["fields"][i]
            old_name = old.get("name", "")
            dlg = FieldEditorDialog(win, self.profile, old)
            win.wait_window(dlg)
            if not dlg.result:
                return
            new = dlg.result
            new_name = new["name"]

            # Se il nuovo nome appartiene già a un altro campo, non creiamo
            # due colonne omonime: chiediamo esplicitamente se l'utente intende
            # dichiarare che si tratta dello stesso campo e, in tal caso,
            # fondiamo le due definizioni e reindirizziamo tutte le regole.
            target_index = next((
                j for j, field in enumerate(self.profile.get("fields", []))
                if j != i and isinstance(field, dict) and field.get("name") == new_name
            ), None)
            if target_index is not None:
                if not messagebox.askyesno(
                    "Unisci campi",
                    f"Esiste già il campo '{new_name}'.\n\n"
                    f"Vuoi indicare che '{old_name}' e '{new_name}' sono lo stesso campo?\n\n"
                    "Scegliendo Sì, tutte le regole che scrivono nel primo campo "
                    "saranno reindirizzate al campo esistente e la riga duplicata "
                    "verrà eliminata.",
                    parent=win,
                ):
                    return

                # Applichiamo al campo sorgente le opzioni appena scelte nel
                # dialogo, mantenendo temporaneamente il vecchio nome per poter
                # aggiornare correttamente tutti i riferimenti.
                source_for_merge = dict(new)
                source_for_merge["name"] = old_name
                self.profile["fields"][i] = source_for_merge
                keys = self.profile.setdefault("reconcile_keys", [])
                if dlg.key_var.get() and old_name not in keys:
                    keys.append(old_name)
                if not dlg.key_var.get() and old_name in keys:
                    keys.remove(old_name)

                merge_field_into(self.profile, old_name, new_name, source_index=i)
                self.status_var.set(f"Campi uniti: {old_name} → {new_name}")
                messagebox.showinfo(
                    "Campi uniti",
                    f"'{old_name}' ora è lo stesso campo di '{new_name}'.\n"
                    "Le regole e le chiavi di riconciliazione sono state aggiornate.",
                    parent=win,
                )
                refresh(); self._update_extraction_preview()
                return

            self.profile["fields"][i] = new
            # Se il campo viene rinominato verso un nome nuovo, aggiorna tutte
            # le regole, le chiavi, i campi calcolati e gli alias riepilogativi.
            if new_name != old_name:
                for mt2 in self.profile.get("mail_types", []):
                    for rule in mt2.get("rules", []):
                        if rule.get("field") == old_name:
                            rule["field"] = new_name
                for comp in self.profile.get("computed_fields", []):
                    for attr in ("field", "start_field", "end_field"):
                        if comp.get(attr) == old_name:
                            comp[attr] = new_name
                self.profile["reconcile_keys"] = [new_name if x == old_name else x for x in self.profile.get("reconcile_keys", [])]
                # Aggiorna anche gli alias usati nelle formule riepilogative nominate.
                for sf in self.profile.get("summary_formulas", []) or []:
                    if not isinstance(sf, dict):
                        continue
                    formula_text = str(sf.get("formula", "") or "")
                    for prefix in ("somma_", "conta_"):
                        formula_text = re.sub(
                            rf"\b{re.escape(prefix + old_name)}\b",
                            prefix + new_name, formula_text, flags=re.I,
                        )
                    sf["formula"] = formula_text
                    if str(sf.get("kind", "") or "") == "period_aggregate":
                        if sf.get("field") == old_name:
                            sf["field"] = new_name
                        if sf.get("date_field") == old_name:
                            sf["date_field"] = new_name
            keys = self.profile.setdefault("reconcile_keys", [])
            if dlg.key_var.get() and new_name not in keys:
                keys.append(new_name)
            if not dlg.key_var.get() and new_name in keys:
                keys.remove(new_name)
            refresh(); self._update_extraction_preview()

        def add_field():
            dlg = FieldEditorDialog(win, self.profile)
            win.wait_window(dlg)
            if not dlg.result: return
            if get_field(self.profile, dlg.result["name"]):
                messagebox.showerror("Campo", "Esiste già un campo con questo nome.", parent=win)
                return
            self.profile.setdefault("fields", []).append(dlg.result)
            if dlg.key_var.get():
                self.profile.setdefault("reconcile_keys", []).append(dlg.result["name"])
            refresh()

        def delete_field():
            sel = tree.selection()
            if not sel: return
            i = int(sel[0])
            field = self.profile["fields"][i]
            name = field.get("name", "")
            used_rules = sum(
                1 for mt2 in self.profile.get("mail_types", [])
                for rule in mt2.get("rules", []) if rule.get("field") == name
            )
            if used_rules:
                messagebox.showwarning(
                    "Campo in uso",
                    f"Il campo '{name}' è usato da {used_rules} regole. Modifica o elimina prima quelle regole.",
                    parent=win,
                )
                return
            computed_users = [
                str(comp.get("field", "")) for comp in self.profile.get("computed_fields", []) or []
                if isinstance(comp, dict) and name in {
                    str(comp.get("field", "") or ""),
                    str(comp.get("start_field", "") or ""),
                    str(comp.get("end_field", "") or ""),
                }
            ]
            if computed_users:
                messagebox.showwarning(
                    "Campo usato da campi calcolati",
                    f"Il campo '{name}' è usato dai calcoli: {', '.join(computed_users)}. "
                    "Modifica o elimina prima quei campi calcolati.",
                    parent=win,
                )
                return
            aliases = {f"somma_{name}", f"conta_{name}"}
            formula_users = [
                str(sf.get("name", "")) for sf in self.profile.get("summary_formulas", []) or []
                if isinstance(sf, dict) and any(
                    re.search(rf"\b{re.escape(alias)}\b", str(sf.get("formula", "") or ""), flags=re.I)
                    for alias in aliases
                )
            ]
            if formula_users:
                messagebox.showwarning(
                    "Campo usato da formule riepilogative",
                    f"Il campo '{name}' è richiamato dalle formule: {', '.join(formula_users)}. "
                    "Modifica o elimina prima quelle formule.",
                    parent=win,
                )
                return
            if messagebox.askyesno("Elimina campo", f"Eliminare il campo '{name}'?", parent=win):
                del self.profile["fields"][i]
                self.profile["reconcile_keys"] = [x for x in self.profile.get("reconcile_keys", []) if x != name]
                refresh()

        def move_field(delta: int):
            sel = tree.selection()
            if not sel:
                return
            i = int(sel[0])
            fields = self.profile.get("fields", [])
            j = i + delta
            if j < 0 or j >= len(fields):
                return
            fields[i], fields[j] = fields[j], fields[i]
            refresh()
            tree.selection_set(str(j))
            tree.focus(str(j))
            tree.see(str(j))
            self.status_var.set("Ordine colonne Excel aggiornato")

        def show_field_generation_rules():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo(
                    "Regole del campo",
                    "Seleziona prima un campo nella tabella.",
                    parent=win,
                )
                return
            i = int(sel[0])
            fields = self.profile.get("fields", [])
            if i < 0 or i >= len(fields):
                return
            field_name = str(fields[i].get("name", "") or "")
            FieldGenerationRulesDialog(win, self.profile, field_name)

        def computed_fields_changed():
            refresh()
            self.status_var.set("Campi calcolati aggiornati")

        btns = ttk.Frame(frm); btns.pack(fill="x", pady=(8,0))
        ttk.Button(btns, text="Nuovo campo", command=add_field).pack(side="left")
        ttk.Button(btns, text="Modifica", command=edit_field).pack(side="left", padx=5)
        ttk.Button(btns, text="Elimina", command=delete_field).pack(side="left", padx=5)
        ttk.Button(
            btns, text="Regole che generano…", command=show_field_generation_rules
        ).pack(side="left", padx=5)
        ttk.Separator(btns, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(btns, text="↑ Sposta su", command=lambda: move_field(-1)).pack(side="left", padx=(0, 4))
        ttk.Button(btns, text="↓ Sposta giù", command=lambda: move_field(1)).pack(side="left")
        ttk.Separator(btns, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(
            btns, text="Campi calcolati…",
            command=lambda: ComputedFieldsDialog(
                win, self.profile, on_change=computed_fields_changed
            ),
        ).pack(side="left")
        ttk.Button(
            btns, text="Formule riepilogative…",
            command=lambda: SummaryFormulasDialog(
                win, self.profile,
                on_change=lambda: self._save_profile_immediately("Formule riepilogative aggiornate e salvate")
            ),
        ).pack(side="left", padx=(5, 0))
        ttk.Button(
            btns, text="Valori schermata iniziale…",
            command=lambda: HomeSummaryFieldsDialog(
                win, self.profile,
                on_change=lambda: self._save_profile_immediately(
                    "Valori schermata iniziale aggiornati e salvati"
                ),
            ),
        ).pack(side="left", padx=(5, 0))
        ttk.Button(btns, text="Chiudi", command=win.destroy).pack(side="right")
        tree.bind("<Double-1>", lambda _e: edit_field())
        refresh(); win.transient(self); win.grab_set()

    def _save_profile_immediately(self, status_message: str = "Profilo aggiornato e salvato") -> None:
        """Salva immediatamente le modifiche effettuate dalle finestre secondarie."""
        self.profile_mgr.save(self.profile_path, self.profile)
        self.status_var.set(status_message)
        self._refresh_home_profile_buttons()

    def save_profile(self):
        self.profile_mgr.save(self.profile_path, self.profile)
        self.status_var.set(f"Profilo salvato: {self.profile_path.name}")

    @staticmethod
    def _open_file_cross_platform(path: str | Path) -> None:
        """Apre il file con l'applicazione associata su Windows, macOS o Linux."""
        target = str(Path(path).resolve())
        if sys.platform.startswith("win"):
            os.startfile(target)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target])
        else:
            subprocess.Popen(["xdg-open", target])

    def open_imap_profiles_folder(self):
        """Apre la cartella che contiene i profili/credenziali IMAP cifrati."""
        path = self.config_mgr.email_dir
        try:
            path.mkdir(parents=True, exist_ok=True)
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror(
                "Cartella email",
                f"Impossibile aprire la cartella delle email:\n{path}\n\n{exc}",
                parent=self,
            )

    def open_extraction_profiles_folder(self):
        """Apre la cartella che contiene i profili JSON di estrazione."""
        path = Path(self.config_mgr.profiles_dir)
        try:
            path.mkdir(parents=True, exist_ok=True)
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror(
                "Cartella profili di estrazione",
                f"Impossibile aprire la cartella dei profili di estrazione:\n{path}\n\n{exc}",
                parent=self,
            )

    def open_general_excel(self):
        """Apre l'Excel generale del profilo attivo dalla cartella di lavoro."""
        path = self._general_excel_path()
        if not path.exists():
            remembered = self._last_excel_path_for_profile()
            if remembered is not None:
                path = remembered
        if not path.exists():
            messagebox.showinfo(
                "Excel generale",
                "Non esiste ancora un Excel generale per questo profilo di estrazione.",
                parent=self,
            )
            return
        try:
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror("Excel generale", f"Impossibile aprire:\n{path}\n\n{exc}", parent=self)

    def open_work_folder(self):
        """Apre la cartella degli Excel configurata nelle Preferenze."""
        path = self.config_mgr.work_dir
        try:
            path.mkdir(parents=True, exist_ok=True)
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror("Cartella degli Excel", f"Impossibile aprire:\n{path}\n\n{exc}", parent=self)

    def open_workspace_folder(self):
        """Apre il workspace dell'applicazione (configurazione, log e dati di servizio)."""
        path = Path(self.config_mgr.app_dir)
        try:
            path.mkdir(parents=True, exist_ok=True)
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror("Cartella workspace", f"Impossibile aprire:\n{path}\n\n{exc}", parent=self)

    def _community_pro_feature_notice(self):
        messagebox.showinfo(
            "EgoMailExtractor Community",
            "Questa funzione guidata è disponibile in EgoMailExtractor Pro.\n\n"
            "Nella Community puoi configurare manualmente campi, regole e formule dalla schermata completa.",
            parent=self,
        )

    def wizard_choose_profile_icon(self):
        self._community_pro_feature_notice()

    def wizard_summary_fields(self):
        self._community_pro_feature_notice()

    def wizard_summary_formulas(self):
        self._community_pro_feature_notice()

    def request_support_email(self):
        """Configura esplicitamente SMTP e invia una richiesta assistenza solo dopo conferma."""
        imap = self.config_mgr.imap
        password = self.config_mgr.get_password()
        if not imap.host or not imap.username or not password:
            messagebox.showwarning(
                "Richiedi assistenza",
                "Per proporre la configurazione SMTP è necessario avere prima una email configurata e sbloccata.",
                parent=self,
            )
            return

        smtp_host, smtp_port, smtp_security = infer_smtp_from_imap(imap.host, imap.port, imap.ssl)
        dlg = tk.Toplevel(self)
        dlg.title("Richiedi assistenza - configurazione SMTP")
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("700x720")
        dlg.minsize(650, 630)
        center_dialog_later(dlg, self)

        frame = ttk.Frame(dlg, padding=14)
        frame.pack(fill="both", expand=True)
        notice = (
            "Per inviare la richiesta di assistenza EgoMailExtractor deve configurare una connessione SMTP.\n"
            "I dati proposti derivano dal email attiva: controllali e confermali prima dell'invio.\n"
            "La password non viene inviata all'assistenza e non viene salvata in una nuova configurazione separata."
        )
        ttk.Label(frame, text=notice, wraplength=650, justify="left").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        host_v = tk.StringVar(value=smtp_host)
        port_v = tk.StringVar(value=str(smtp_port))
        sec_v = tk.StringVar(value=smtp_security)
        user_v = tk.StringVar(value=imap.username)
        pwd_v = tk.StringVar(value=password)
        to_v = tk.StringVar(value="ennio@domoticachepassione.it")
        subject_v = tk.StringVar(value=f"{DISPLAY_NAME} {APP_VERSION} - Richiesta assistenza")
        confirm_v = tk.BooleanVar(value=False)
        show_pwd_v = tk.BooleanVar(value=False)
        attach_log_v = tk.BooleanVar(value=False)
        attach_profiles_v = tk.BooleanVar(value=False)
        attach_preferences_v = tk.BooleanVar(value=False)

        fields = [
            ("Server SMTP", host_v),
            ("Porta SMTP", port_v),
            ("Utente / mittente", user_v),
            ("Email assistenza (destinatario)", to_v),
            ("Oggetto", subject_v),
        ]
        row = 1
        entries = {}
        for label, var in fields:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ent = ttk.Entry(frame, textvariable=var, width=56)
            ent.grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
            entries[label] = ent
            row += 1

        ttk.Label(frame, text="Sicurezza").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=sec_v, values=("SSL", "STARTTLS", "NONE"), state="readonly", width=18).grid(row=row, column=1, sticky="w", pady=4)
        row += 1

        ttk.Label(frame, text="Password SMTP").grid(row=row, column=0, sticky="w", pady=4)
        pwd_entry = ttk.Entry(frame, textvariable=pwd_v, show="•", width=42)
        pwd_entry.grid(row=row, column=1, sticky="ew", pady=4)
        def toggle_password():
            pwd_entry.configure(show="" if show_pwd_v.get() else "•")
        ttk.Checkbutton(frame, text="Mostra", variable=show_pwd_v, command=toggle_password).grid(row=row, column=2, sticky="w", padx=(8,0))
        row += 1

        ttk.Label(frame, text="Messaggio").grid(row=row, column=0, sticky="nw", pady=(8,4))
        body = tk.Text(frame, height=12, wrap="word")
        body.grid(row=row, column=1, columnspan=2, sticky="nsew", pady=(8,4))
        body.insert("1.0", f"Descrivi qui il problema.\n\nVersione: {DISPLAY_NAME} {APP_VERSION}\nEmail: {self.config_mgr.active_imap_profile_name}\n")
        frame.rowconfigure(row, weight=1)
        row += 1

        ttk.Label(frame, text="Allegati assistenza (facoltativi)").grid(row=row, column=0, sticky="nw", pady=(10,4))
        attachments_frame = ttk.Frame(frame)
        attachments_frame.grid(row=row, column=1, columnspan=2, sticky="w", pady=(10,4))
        ttk.Checkbutton(attachments_frame, text="Allega log di estrazione in uno ZIP", variable=attach_log_v).pack(anchor="w")
        ttk.Checkbutton(attachments_frame, text="Allega profili di estrazione in uno ZIP", variable=attach_profiles_v).pack(anchor="w")
        ttk.Checkbutton(attachments_frame, text="Allega preferenze (settings.json) in uno ZIP", variable=attach_preferences_v).pack(anchor="w")
        ttk.Label(attachments_frame, text="Nessun allegato viene inviato se non lo selezioni esplicitamente. Le preferenze non contengono password email/SMTP.", wraplength=480).pack(anchor="w", pady=(3,0))
        row += 1

        ttk.Checkbutton(
            frame,
            text="Confermo di aver verificato i dati SMTP sopra indicati e autorizzo EgoMailExtractor a usarli per questa richiesta di assistenza.",
            variable=confirm_v,
        ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(10,8))
        row += 1

        buttons = ttk.Frame(frame)
        buttons.grid(row=row, column=0, columnspan=3, sticky="e")

        def do_send():
            if not confirm_v.get():
                messagebox.showwarning("Conferma richiesta", "Devi prima confermare esplicitamente i dati SMTP.", parent=dlg)
                return
            try:
                port = int(port_v.get().strip())
            except ValueError:
                messagebox.showerror("Dati SMTP", "La porta SMTP deve essere un numero.", parent=dlg)
                return
            if not to_v.get().strip():
                messagebox.showerror("Dati SMTP", "Inserisci l'indirizzo email dell'assistenza.", parent=dlg)
                return
            summary = (
                "Confermi la configurazione SMTP e l'invio?\n\n"
                f"Server: {host_v.get().strip()}\n"
                f"Porta: {port}\n"
                f"Sicurezza: {sec_v.get()}\n"
                f"Mittente: {user_v.get().strip()}\n"
                f"Destinatario: {to_v.get().strip()}\n"
                f"Oggetto: {subject_v.get().strip()}\n\n"
                "La password è configurata ma non verrà inclusa nel messaggio.\n\n"
                f"Allega log di estrazione: {'SÌ' if attach_log_v.get() else 'NO'}\n"
                f"Allega profili di estrazione: {'SÌ' if attach_profiles_v.get() else 'NO'}\n"
                f"Allega preferenze: {'SÌ' if attach_preferences_v.get() else 'NO'}"
            )
            if not messagebox.askyesno("Conferma invio assistenza", summary, parent=dlg):
                return
            settings = SmtpSettings(
                host=host_v.get().strip(), port=port, security=sec_v.get(),
                username=user_v.get().strip(), password=pwd_v.get(),
            )
            try:
                attachments = []
                temp_dir = None
                if attach_log_v.get() or attach_profiles_v.get() or attach_preferences_v.get():
                    temp_dir = tempfile.TemporaryDirectory(prefix="egomail_support_")
                    zip_path = Path(temp_dir.name) / f"EgoMailExtractor_assistenza_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        if attach_log_v.get():
                            log_path = self._extraction_log_path()
                            if log_path.exists() and log_path.is_file(): zf.write(log_path, arcname=f"log/{log_path.name}")
                        if attach_profiles_v.get():
                            profiles_dir = Path(self.config_mgr.profiles_dir)
                            if profiles_dir.exists():
                                for profile_file in profiles_dir.glob("*.json"):
                                    if profile_file.is_file(): zf.write(profile_file, arcname=f"profili_estrazione/{profile_file.name}")
                        if attach_preferences_v.get():
                            pref_path = Path(self.config_mgr.path)
                            if pref_path.exists() and pref_path.is_file():
                                zf.write(pref_path, arcname="preferenze/settings.json")
                    attachments.append(zip_path)
                send_support_email(settings, to_v.get(), subject_v.get(), body.get("1.0", "end-1c"), attachments=attachments)
                if temp_dir is not None: temp_dir.cleanup()
            except Exception as exc:
                messagebox.showerror("Invio assistenza", f"Invio non riuscito.\n\n{exc}", parent=dlg)
                return
            messagebox.showinfo("Invio assistenza", "Richiesta di assistenza inviata correttamente.", parent=dlg)
            dlg.destroy()

        ttk.Button(buttons, text="Annulla", command=dlg.destroy).pack(side="right", padx=(8,0))
        ttk.Button(buttons, text="Conferma e invia", command=do_send).pack(side="right")
        frame.columnconfigure(1, weight=1)
        dlg.wait_window()

    def open_license(self):
        """Apre la licenza dell'edizione installata."""
        license_name = "LICENSE_PRO.txt" if IS_PRO else "LICENSE.txt"
        path = _resource_path(license_name)
        if not path.exists():
            messagebox.showerror("Licenza", f"File di licenza non trovato:\n{path}", parent=self)
            return
        try:
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror("Licenza", f"Impossibile aprire la licenza:\n{path}\n\n{exc}", parent=self)

    def open_extraction_log(self):
        path = self._extraction_log_path()
        if not path.exists():
            messagebox.showinfo(
                "Log estrazione",
                f"Il log non è ancora stato creato.\nVerrà salvato qui dopo la prima estrazione completata:\n{path}",
                parent=self,
            )
            return
        try:
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror("Log estrazione", f"Impossibile aprire il log:\n{path}\n\n{exc}", parent=self)

    def open_extraction_log_folder(self):
        path = self._extraction_log_path().parent
        try:
            path.mkdir(parents=True, exist_ok=True)
            self._open_file_cross_platform(path)
        except Exception as exc:
            messagebox.showerror(
                "Cartella log",
                f"Impossibile aprire la cartella del log:\n{path}\n\n{exc}",
                parent=self,
            )

    def clear_extraction_log(self):
        path = self._extraction_log_path()
        if not path.exists():
            messagebox.showinfo(
                "Cancella log",
                f"Il log non esiste ancora:\n{path}",
                parent=self,
            )
            return
        if not messagebox.askyesno(
            "Cancella log",
            "Cancellare definitivamente tutto il contenuto del log di estrazione?\n\n"
            f"{path}",
            parent=self,
        ):
            return
        try:
            path.unlink()
            self.status_var.set(f"Log cancellato: {path.name}")
        except Exception as exc:
            messagebox.showerror(
                "Cancella log",
                f"Impossibile cancellare il log:\n{path}\n\n{exc}",
                parent=self,
            )

    def _profile_settings_key(self) -> str:
        try:
            base = str(Path(self.profile_path).resolve())
        except Exception:
            base = str(self.profile_path)
        return f"{base}::imap={self.config_mgr.active_imap_profile_id or 'none'}"

    def _active_extraction_state(self) -> dict:
        """Stato operativo della coppia email + profilo, salvato nelle Preferenze email."""
        state = self.config_mgr.extraction_state(self.profile_path.name)
        # Migrazione una tantum dalle vecchie versioni, nelle quali lo stato era
        # contenuto nel JSON del profilo di estrazione.
        if not state.get("_migrated_from_profile"):
            pid = self.config_mgr.active_imap_profile_id or "default"
            old_states = self.profile.get("state_by_email", {}) or {}
            legacy = old_states.get(pid, {}) if isinstance(old_states, dict) else {}
            if not legacy and not state.get("last_processed_uid"):
                legacy = self.profile.get("state", {}) or {}
            if isinstance(legacy, dict):
                for key in ("folder","uidvalidity","last_processed_uid","last_processed_at","first_processed_uid","first_processed_at","home_summary_values","home_summary_updated_at"):
                    if key in legacy and not state.get(key): state[key] = legacy[key]
            if self.profile.get("extraction_start_date") and not state.get("extraction_start_date"):
                state["extraction_start_date"] = self.profile.get("extraction_start_date")
            state["_migrated_from_profile"] = True
            self.config_mgr.save()
        return state

    def _touch_active_extraction_state(self, folder: str, uidvalidity: int, last_uid: int) -> None:
        state = self._active_extraction_state()
        now = datetime.now().isoformat(timespec="seconds")
        if not int(state.get("first_processed_uid", 0) or 0) and int(last_uid or 0):
            state["first_processed_uid"] = int(last_uid or 0)
            state["first_processed_at"] = now
        state.update({
            "folder": folder, "uidvalidity": int(uidvalidity or 0),
            "last_processed_uid": int(last_uid or 0),
            "last_processed_at": now,
        })
        self.config_mgr.save()

    def _remember_last_excel_path(self, path: str | Path) -> None:
        """Memorizza per profilo l'ultimo Excel scritto con successo."""
        mapping = self.config_mgr.data.setdefault("last_excel_by_profile", {})
        if not isinstance(mapping, dict):
            mapping = {}
            self.config_mgr.data["last_excel_by_profile"] = mapping
        mapping[self._profile_settings_key()] = str(Path(path).resolve())
        try:
            self.config_mgr.save()
        except Exception:
            # Il file Excel è già stato scritto: un eventuale problema nel
            # salvataggio della preferenza non deve invalidare l'estrazione.
            pass

    def _last_excel_path_for_profile(self) -> Path | None:
        mapping = self.config_mgr.data.get("last_excel_by_profile", {}) or {}
        if isinstance(mapping, dict):
            value = str(mapping.get(self._profile_settings_key(), "") or "").strip()
            if value:
                path = Path(value)
                if path.exists() and path.is_file():
                    return path
        # Compatibilità con estrazioni effettuate prima della v1.16: il log
        # registra sempre il percorso Excel. Usiamo l'ultima occorrenza valida.
        log_path = self._extraction_log_path()
        if log_path.exists():
            try:
                last_from_log = ""
                for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines():
                    if line.startswith("File Excel:"):
                        candidate = line.split(":", 1)[1].strip()
                        if candidate:
                            last_from_log = candidate
                if last_from_log:
                    legacy_path = Path(last_from_log)
                    if legacy_path.exists() and legacy_path.is_file():
                        return legacy_path
            except Exception:
                pass
        # Fallback naturale per il file generale del profilo.
        general = self._general_excel_path()
        if general.exists() and general.is_file():
            return general
        return None

    def _debug_parameters(self, excel_path: Path | None) -> dict:
        """Snapshot dei parametri utili al debug, senza password o segreti."""
        imap = self.config_mgr.imap
        return {
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "application": {
                "name": APP_TITLE,
                "version": APP_VERSION,
                "date": APP_DATE,
            },
            "profile": {
                "name": self.profile.get("name", ""),
                "path": str(Path(self.profile_path).resolve()),
                "schema_version": self.profile.get("schema_version"),
                "state": dict(self._active_extraction_state()),
            },
            "imap": {
                "profile_name": self.config_mgr.active_imap_profile_name,
                "host": imap.host,
                "port": imap.port,
                "ssl": imap.ssl,
                "username": imap.username,
                "folder": self.folder_var.get().strip() or imap.folder or "INBOX",
                "password": "<NON ESPORTATA>",
            },
            "search_runtime": {
                "subject_filter": self.subject_var.get(),
                "sender_filter": self.sender_var.get(),
                "active_mail_type_filters": sorted(self.filter_mail_types),
            },
            "settings": {
                "profiles_dir": str(self.config_mgr.profiles_dir),
                "work_dir": str(self.config_mgr.work_dir),
                "regex_abbreviations": list(self.config_mgr.regex_abbreviations),
            },
            "files": {
                "extraction_log": str(self._extraction_log_path()),
                "last_excel": str(excel_path) if excel_path else "",
            },
            "security_note": (
                "Password master, password email e file imap_credentials.enc non sono inclusi nel pacchetto."
            ),
        }

    def export_debug_bundle(self):
        """Esporta log, ultimo Excel, profilo e parametri in un unico ZIP."""
        default_name = f"{self.profile_path.stem}_debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Esporta pacchetto debug",
            defaultextension=".zip",
            filetypes=[("Archivio ZIP", "*.zip")],
            initialfile=default_name,
        )
        if not path:
            return
        log_path = self._extraction_log_path()
        excel_path = self._last_excel_path_for_profile()
        try:
            manifest = build_debug_bundle(
                path,
                profile=self.profile,
                profile_path=self.profile_path,
                log_path=log_path,
                excel_path=excel_path,
                parameters=self._debug_parameters(excel_path),
                preferences_path=self.config_mgr.path,
            )
            missing = manifest.get("missing", []) or []
            extra = ""
            if missing:
                extra = "\n\nNon disponibili e quindi non inclusi: " + ", ".join(missing) + "."
            messagebox.showinfo(
                "Pacchetto debug creato",
                f"Creato:\n{Path(path).resolve()}\n\n"
                "Il pacchetto non contiene password né il file cifrato delle credenziali."
                + extra,
                parent=self,
            )
            self.status_var.set(f"Pacchetto debug creato: {Path(path).name}")
        except Exception as exc:
            messagebox.showerror(
                "Pacchetto debug",
                f"Impossibile creare il pacchetto debug:\n{exc}",
                parent=self,
            )

    def stop_extraction(self):
        """Richiede l'arresto dell'estrazione in corso.

        L'operazione corrente di lettura IMAP può terminare, poi il ciclo si ferma.
        I file JSON di lavoro già completati restano su disco per consentire la ripresa.
        """
        if not self.extraction_running:
            self.status_var.set("Nessuna estrazione in corso")
            return
        self.extraction_stop_event.set()
        self.status_var.set("Interruzione estrazione richiesta…")

    @staticmethod
    def _mail_chronological_key(mail: MailHeader | MailMessage) -> tuple[float, int]:
        """Ordina realmente per data della mail; UID usato come spareggio/fallback."""
        try:
            raw_date = getattr(mail, "raw_date", "") or mail.date
            dt = parsedate_to_datetime(raw_date)
            if dt is None:
                raise ValueError("data assente")
            if dt.tzinfo is None:
                ts = dt.timestamp()
            else:
                ts = dt.timestamp()
            return (float(ts), int(mail.uid))
        except Exception:
            # Una data non interpretabile viene posta in fondo, mantenendo
            # comunque l'ordine crescente degli UID.
            return (float("inf"), int(mail.uid))

    def _ask_retry_imap_from_worker(self, error: Exception, context: str) -> bool:
        answer_event = threading.Event()
        answer_box: dict[str, bool] = {}
        self.ui_queue.put(("extraction_retry_prompt", (error, context, answer_event, answer_box)))
        while not answer_event.wait(0.2):
            if self.extraction_stop_event.is_set():
                return False
        return bool(answer_box.get("retry", False))

    def _imap_call_with_reconnect(self, func, context: str):
        """Esegue una chiamata IMAP con riconnessione automatica e retry interattivo."""
        while True:
            try:
                return func()
            except Exception as first_error:
                if self.extraction_stop_event.is_set():
                    raise ExtractionRetryCancelled(str(first_error)) from first_error
                self.ui_queue.put((
                    "done",
                    f"Connessione email interrotta durante {context}. Riconnessione automatica…",
                ))
                time.sleep(1.0)
                try:
                    # I metodi ImapService aprono una nuova connessione ad ogni chiamata:
                    # ripetere la chiamata equivale quindi a riconnettersi.
                    return func()
                except Exception as retry_error:
                    if not self._ask_retry_imap_from_worker(retry_error, context):
                        raise ExtractionRetryCancelled(str(retry_error)) from retry_error
                    self.ui_queue.put(("done", f"Nuovo tentativo email durante {context}…"))
                    time.sleep(0.5)

    def _extraction_checkpoint_path(self) -> Path:
        config = getattr(self, "config_mgr", None)
        email_name = getattr(config, "active_imap_profile_name", "") if config is not None else ""
        if not email_name:
            return self.profile_path.with_name(self.profile_path.stem + "_estrazione_checkpoint.json")
        email = safe_filename_component(email_name, "email")
        return self.profile_path.with_name(self.profile_path.stem + f"_{email}_estrazione_checkpoint.json")

    def _load_extraction_checkpoint(self, folder: str, uidvalidity: int, base_last_uid: int, excel_path: Path) -> set[int]:
        path = self._extraction_checkpoint_path()
        if not path.exists() or not excel_path.exists():
            return set()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if str(data.get("folder", "")) != str(folder):
                return set()
            if int(data.get("uidvalidity", 0) or 0) != int(uidvalidity or 0):
                return set()
            if int(data.get("base_last_uid", 0) or 0) != int(base_last_uid or 0):
                return set()
            return {int(x) for x in data.get("processed_uids", [])}
        except Exception:
            return set()

    def _save_extraction_checkpoint(self, folder: str, uidvalidity: int, base_last_uid: int, processed_uids: set[int], excel_path: Path) -> None:
        path = self._extraction_checkpoint_path()
        data = {
            "profile": self.profile_path.name,
            "folder": folder,
            "uidvalidity": int(uidvalidity or 0),
            "base_last_uid": int(base_last_uid or 0),
            "processed_uids": sorted(int(x) for x in processed_uids),
            "excel_path": str(excel_path),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)

    def _clear_extraction_checkpoint(self) -> None:
        try:
            self._extraction_checkpoint_path().unlink(missing_ok=True)
        except Exception:
            pass


    def _pipeline_cache_dir(self) -> Path:
        path = Path(self.config_mgr.work_dir) / ".egomailextractor"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _pipeline_key(self) -> str:
        profile_name = safe_filename_component(Path(self.profile_path).stem, "profilo")
        email_name = safe_filename_component(self.config_mgr.active_imap_profile_name or "email", "email")
        return f"{profile_name}__{email_name}"

    def _reading_cache_path(self) -> Path:
        return self._pipeline_cache_dir() / f"{self._pipeline_key()}__lettura.json"

    def _operations_cache_path(self) -> Path:
        return self._pipeline_cache_dir() / f"{self._pipeline_key()}__estrazione.json"

    def _profile_pipeline_signature(self) -> str:
        # Firma solo la configurazione funzionale. Lo stato email/profilo è
        # memorizzato separatamente e non deve invalidare la cache.
        import hashlib
        payload = json.dumps(self.profile, sort_keys=True, ensure_ascii=False, default=str)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def _pipeline_base_metadata(self, folder: str, base_last_uid: int) -> dict:
        return {
            "pipeline_version": 1,
            "profile_file": Path(self.profile_path).name,
            "profile_name": str(self.profile.get("name", "") or Path(self.profile_path).stem),
            "profile_signature": self._profile_pipeline_signature(),
            "email": self.config_mgr.active_imap_profile_name or "",
            "folder": folder,
            "base_last_uid": int(base_last_uid or 0),
            "start_date": str(self._active_extraction_state().get("extraction_start_date", "") or ""),
        }

    def _pipeline_cache_matches(self, data: dict | None, folder: str, base_last_uid: int) -> bool:
        if not isinstance(data, dict):
            return False
        expected = self._pipeline_base_metadata(folder, base_last_uid)
        for key in ("pipeline_version", "profile_file", "profile_signature", "email", "folder", "base_last_uid", "start_date"):
            if str(data.get(key, "")) != str(expected.get(key, "")):
                return False
        return True

    def _restore_original_if_pipeline_stale(self, final_path: Path) -> None:
        original, temporary = transactional_paths(final_path)
        safe_unlink(temporary)
        if original.exists() and not final_path.exists():
            try:
                shutil.copy2(original, final_path)
            except Exception:
                pass
        elif original.exists() and final_path.exists():
            # Il file definitivo è già stato committato e lo stato avanzato:
            # il residuo "originale" può essere rimosso.
            safe_unlink(original)

    @staticmethod
    def _event_stats_add(total: dict[str, int], stats: dict[str, int]) -> None:
        for key in ("mails", "matched_mails", "type_matches", "field_updates"):
            total[key] = int(total.get(key, 0) or 0) + int(stats.get(key, 0) or 0)

    def _extraction_log_path(self) -> Path:
        config = getattr(self, "config_mgr", None)
        email_name = getattr(config, "active_imap_profile_name", "") if config is not None else ""
        if not email_name:
            return self.profile_path.with_name(self.profile_path.stem + "_estrazione.log")
        email = safe_filename_component(email_name, "email")
        return self.profile_path.with_name(self.profile_path.stem + f"_{email}_estrazione.log")

    def _fetch_and_extract(self, uids: list[int], folder: str) -> tuple[list[dict], int]:
        # Determiniamo prima l'ordine cronologico usando solo gli header.
        # Solo dopo scarichiamo/applichiamo il corpo, iniziando quindi davvero
        # dalla mail più vecchia.
        unique_uids = sorted({int(uid) for uid in uids})
        max_uid = max(unique_uids, default=0)
        header_map: dict[int, MailHeader | MailMessage] = {}
        missing: list[int] = []
        for uid in unique_uids:
            if uid in self.messages:
                header_map[uid] = self.messages[uid]
            elif uid in self.headers:
                header_map[uid] = self.headers[uid]
            else:
                missing.append(uid)

        if missing:
            fetched_headers = self._imap_call_with_reconnect(
                lambda: self.imap.fetch_headers(
                    missing, folder, self.extraction_stop_event,
                    on_progress=lambda i, total: self.ui_queue.put((
                        "done", f"Lettura date mail {i}/{total} per ordinamento cronologico…"
                    )),
                ),
                "la lettura delle date delle email",
            )
            for header in fetched_headers:
                self.headers[header.uid] = header
                header_map[header.uid] = header
        if self.extraction_stop_event.is_set():
            return [], max_uid

        ordered_uids = sorted(
            unique_uids,
            key=lambda uid: self._mail_chronological_key(
                header_map.get(uid) or MailHeader(uid=uid, sender="", subject="", date="")
            ),
        )

        events: list[dict] = []
        for i, uid in enumerate(ordered_uids, 1):
            if self.extraction_stop_event.is_set():
                break
            mail = self.messages.get(uid) or self._imap_call_with_reconnect(
                lambda uid=uid: self.imap.fetch_message(uid, folder),
                f"la lettura della mail UID {uid}",
            )
            self.messages[uid] = mail
            event = extract_mail_detailed(mail, self.profile)
            events.append(event)
            type_names = [str(t.get("mail_type", "")) for t in event.get("types", [])]
            label = ", ".join(type_names) if type_names else "mail non riconosciuta"
            self.ui_queue.put((
                "done",
                f"Estrazione {i}/{len(ordered_uids)} dalla più vecchia — {label}",
            ))
        return events, max_uid

    def export_selected(self):
        uids = sorted(self.selected_uids)
        if not uids:
            # Se non sono state spuntate righe, usa la selezione corrente del Treeview.
            uids = sorted(int(x) for x in self.tree.selection())
        if not uids:
            messagebox.showwarning("Estrazione", "Seleziona almeno una mail.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Salva estrazione Excel", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")], initialfile=f"{self.profile_path.stem}_selezione.xlsx"
        )
        if not path:
            return
        if self.extraction_running:
            messagebox.showwarning("Estrazione", "È già in corso un'estrazione.", parent=self)
            return
        self.extraction_stop_event.clear()
        self.extraction_running = True
        folder = self.folder_var.get().strip() or "INBOX"
        self.status_var.set("Estrazione selezionate…")
        def worker():
            try:
                events, _ = self._fetch_and_extract(uids, folder)
                if self.extraction_stop_event.is_set():
                    self.ui_queue.put((
                        "extraction_stopped",
                        "L'estrazione è stata interrotta. Il file Excel non è stato scritto.",
                    ))
                    return
                records, stats, log_text = apply_detailed_extractions(
                    [], events, self.profile,
                    mode="estrazione mail selezionate", excel_path=path,
                )
                write_excel(path, records, self.profile)
                self._remember_last_excel_path(path)
                log_path = self._extraction_log_path()
                append_extraction_log(log_path, log_text)
                self.ui_queue.put(("extraction_done", {
                    "path": str(path),
                    "message": (
                        f"Creato:\n{path}\n\nRecord estratti: {len(records)}\n"
                        f"Log: {log_path}"
                    ),
                }))
            except ExtractionRetryCancelled as e:
                self.ui_queue.put((
                    "extraction_stopped",
                    f"Estrazione sospesa perché la connessione email non è disponibile.\n\nUltimo errore: {e}",
                ))
            except Exception as e:
                self.ui_queue.put(("extraction_error", e))
        threading.Thread(target=worker, daemon=True).start()


    def extract_incremental(self):
        """Esegue la pipeline persistente Lettura → Estrazione → Consolidamento.

        Lettura:
          - recupera gli header;
          - applica già i criteri Tipo mail valutabili dagli header;
          - ordina le sole candidate e salva la sequenza in JSON.

        Estrazione:
          - scarica il corpo solo delle candidate;
          - salva dopo ogni mail un JSON di operazioni/eventi, senza modificare
            l'Excel generale.

        Consolidamento:
          - preserva l'Excel corrente come "originale";
          - applica in memoria tutte le operazioni a una copia temporanea;
          - salva una sola volta al termine e sostituisce atomicamente il file
            definitivo. In caso di interruzione il temporaneo viene scartato.
        """
        folder = self.folder_var.get().strip() or "INBOX"
        if hasattr(self, "home_stop_button"):
            self.home_stop_button.configure(state="normal")
        state = self._active_extraction_state()
        old_uidv = int(state.get("uidvalidity", 0) or 0)
        last_uid = int(state.get("last_processed_uid", 0) or 0)
        general_path = self._general_excel_path()
        if self.extraction_running:
            messagebox.showwarning("Estrazione", "È già in corso un'estrazione.", parent=self)
            return

        self.extraction_stop_event.clear()
        self.extraction_running = True
        pname = str(self.profile.get("name") or Path(self.profile_path).stem)
        self.status_var.set("Avvio pipeline di estrazione…")
        self.ui_queue.put(("extraction_phase_progress", ("read", 0, 1, pname, "")))
        self.ui_queue.put(("extraction_phase_progress", ("extract", 0, 1, pname, "")))
        self.ui_queue.put(("extraction_phase_progress", ("consolidate", 0, 1, pname, "")))

        def worker():
            read_path = self._reading_cache_path()
            operations_path = self._operations_cache_path()
            log_path = self._extraction_log_path()
            stats_total = {"mails": 0, "matched_mails": 0, "type_matches": 0, "field_updates": 0}
            uidv = old_uidv
            scan_max_uid = last_uid
            effective_last = last_uid
            try:
                # -----------------------------------------------------------------
                # 0. Ripresa: se l'Estrazione era già terminata, si passa
                # direttamente al Consolidamento.
                # -----------------------------------------------------------------
                operations_data = read_json(operations_path)
                if operations_data and not self._pipeline_cache_matches(operations_data, folder, last_uid):
                    self._restore_original_if_pipeline_stale(general_path)
                    safe_unlink(operations_path)
                    safe_unlink(read_path)
                    operations_data = None

                extraction_complete = bool(operations_data and operations_data.get("extraction_complete"))
                if extraction_complete:
                    uidv = int(operations_data.get("uidvalidity", old_uidv) or old_uidv)
                    scan_max_uid = int(operations_data.get("scan_max_uid", last_uid) or last_uid)
                    self.ui_queue.put(("extraction_phase_progress", ("read", 1, 1, pname, "cache già completata")))
                    done_count = len(operations_data.get("events", []) or [])
                    self.ui_queue.put(("extraction_phase_progress", ("extract", max(done_count, 1), max(done_count, 1), pname, "")))
                    self.ui_queue.put(("done", f"Profilo {pname} — estrazione già completata: riprendo dal Consolidamento."))
                else:
                    # -------------------------------------------------------------
                    # 1. LETTURA: riusa il JSON completo se presente e valido.
                    # -------------------------------------------------------------
                    read_data = read_json(read_path)
                    if read_data and not self._pipeline_cache_matches(read_data, folder, last_uid):
                        safe_unlink(read_path)
                        read_data = None

                    def _reading_header_is_candidate(header):
                        # Limite temporale aggiuntivo del profilo.
                        start_date_raw_local = str(state.get("extraction_start_date", "") or "").strip()
                        if start_date_raw_local:
                            try:
                                start_dt_local = datetime.fromisoformat(start_date_raw_local).date()
                                try:
                                    d = parsedate_to_datetime(getattr(header, "raw_date", "") or header.date)
                                    if d is not None and d.date() < start_dt_local:
                                        return False
                                except Exception:
                                    pass
                            except Exception:
                                pass
                        return header_prefilter_matches_profile(
                            MailMessage(
                                uid=header.uid, sender=header.sender, subject=header.subject,
                                date=header.date, message_id=header.message_id,
                                recipient=getattr(header, "recipient", "") or "",
                                body="", raw_date=getattr(header, "raw_date", "") or "",
                            ),
                            self.profile,
                        )

                    if read_data and read_data.get("reading_complete"):
                        candidate_rows = list(read_data.get("candidates", []) or [])
                        uidv = int(read_data.get("uidvalidity", old_uidv) or old_uidv)
                        scan_max_uid = int(read_data.get("scan_max_uid", last_uid) or last_uid)
                        candidate_headers = [dict_to_header(x) for x in candidate_rows if isinstance(x, dict)]
                        candidate_uids = [int(x.uid) for x in candidate_headers]
                        total_scanned = int(read_data.get("scanned_count", len(candidate_uids)) or len(candidate_uids))
                        self.ui_queue.put((
                            "extraction_phase_progress",
                            ("read", max(total_scanned, 1), max(total_scanned, 1), pname, f"candidate: {len(candidate_uids)} — cache"),
                        ))
                        self.ui_queue.put((
                            "done",
                            f"Profilo {pname} — Lettura già completata: uso {len(candidate_uids)} candidate salvate su disco.",
                        ))
                    else:
                        # Se esiste un checkpoint parziale valido, congela la stessa
                        # sequenza UID della sessione e riprende dagli UID non letti.
                        resume_partial = bool(read_data and not read_data.get("reading_complete") and read_data.get("scan_uids"))
                        if resume_partial:
                            uidv = int(read_data.get("uidvalidity", old_uidv) or old_uidv)
                            unique_uids = [int(x) for x in read_data.get("scan_uids", []) or []]
                            scan_max_uid = int(read_data.get("scan_max_uid", max(unique_uids, default=last_uid)) or last_uid)
                            processed_read_uids = {int(x) for x in read_data.get("processed_uids", []) or []}
                            candidate_by_uid = {}
                            for row in read_data.get("candidates", []) or []:
                                if isinstance(row, dict):
                                    try:
                                        candidate_by_uid[int(row.get("uid", 0) or 0)] = row
                                    except Exception:
                                        pass
                            self.ui_queue.put((
                                "done",
                                f"Profilo {pname} — riprendo la Lettura dal checkpoint: "
                                f"{len(processed_read_uids)}/{len(unique_uids)} header già letti, "
                                f"{len(candidate_by_uid)} candidate già trovate.",
                            ))
                        else:
                            start_date_raw = str(state.get("extraction_start_date", "") or "").strip()
                            first_historical_scan = (last_uid <= 0)
                            if first_historical_scan and start_date_raw:
                                try:
                                    start_date = datetime.fromisoformat(start_date_raw).date()
                                    uidv, uids = self._imap_call_with_reconnect(
                                        lambda: self.imap.uids_since_date(folder, start_date),
                                        f"la ricerca storica delle email dal {start_date.strftime('%d/%m/%Y')}",
                                    )
                                except Exception:
                                    uidv, uids = self._imap_call_with_reconnect(
                                        lambda: self.imap.incremental_uids(folder, 0),
                                        "la ricerca storica delle email",
                                    )
                            else:
                                uidv, uids = self._imap_call_with_reconnect(
                                    lambda: self.imap.incremental_uids(folder, last_uid),
                                    "la ricerca delle nuove email",
                                )

                            effective_last = last_uid
                            if old_uidv and uidv and old_uidv != uidv:
                                effective_last = 0
                                uidv, uids = self._imap_call_with_reconnect(
                                    lambda: self.imap.incremental_uids(folder, 0),
                                    "la nuova scansione della casella dopo il cambio UIDVALIDITY",
                                )

                            unique_uids = sorted({int(uid) for uid in uids})
                            scan_max_uid = max(unique_uids, default=effective_last)
                            processed_read_uids = set()
                            candidate_by_uid = {}
                            read_data = {
                                **self._pipeline_base_metadata(folder, last_uid),
                                "uidvalidity": int(uidv or 0),
                                "scan_max_uid": int(scan_max_uid or 0),
                                "scan_uids": unique_uids,
                                "scanned_count": 0,
                                "processed_uids": [],
                                "candidate_count": 0,
                                "reading_complete": False,
                                "created_at": datetime.now().isoformat(timespec="seconds"),
                                "saved_at": datetime.now().isoformat(timespec="seconds"),
                                "candidates": [],
                            }
                            # Primo checkpoint: viene scritto PRIMA di iniziare a
                            # scaricare gli header, così anche un crash immediato
                            # conserva lo snapshot degli UID della sessione.
                            atomic_write_json(read_path, read_data)

                        read_total = max(len(unique_uids), 1)
                        self.ui_queue.put((
                            "extraction_phase_progress",
                            ("read", len(processed_read_uids), read_total, pname, f"candidate: {len(candidate_by_uid)}"),
                        ))

                        remaining_read_uids = [uid for uid in unique_uids if uid not in processed_read_uids]
                        batch_size = 50
                        for batch_start in range(0, len(remaining_read_uids), batch_size):
                            if self.extraction_stop_event.is_set():
                                break
                            batch = remaining_read_uids[batch_start:batch_start + batch_size]
                            header_map: dict[int, MailHeader | MailMessage] = {}
                            missing: list[int] = []
                            for uid in batch:
                                if uid in self.messages:
                                    header_map[uid] = self.messages[uid]
                                elif uid in self.headers:
                                    header_map[uid] = self.headers[uid]
                                else:
                                    missing.append(uid)

                            if missing:
                                fetched_headers = self._imap_call_with_reconnect(
                                    lambda missing=missing: self.imap.fetch_headers(
                                        missing, folder, self.extraction_stop_event,
                                        on_progress=None,
                                    ),
                                    "la lettura degli header delle email",
                                )
                                for header in fetched_headers:
                                    self.headers[header.uid] = header
                                    header_map[header.uid] = header

                            # Il blocco viene considerato letto dopo il tentativo IMAP.
                            # Gli header effettivamente recuperati vengono prefiltrati
                            # subito; gli UID senza header non diventano candidate.
                            for uid in batch:
                                header = header_map.get(uid)
                                if header is not None:
                                    if _reading_header_is_candidate(header):
                                        candidate_by_uid[int(uid)] = header_to_dict(header)
                                    processed_read_uids.add(int(uid))
                                elif not self.extraction_stop_event.is_set():
                                    # Se il server ha completato il blocco ma un UID non
                                    # restituisce header, lo registriamo come letto/non candidato.
                                    # Se invece l'utente ha interrotto, l'UID resta pendente
                                    # e verrà ritentato alla ripresa.
                                    processed_read_uids.add(int(uid))

                            read_data.update({
                                "uidvalidity": int(uidv or 0),
                                "scan_max_uid": int(scan_max_uid or 0),
                                "scan_uids": unique_uids,
                                "scanned_count": len(processed_read_uids),
                                "processed_uids": sorted(processed_read_uids),
                                "candidate_count": len(candidate_by_uid),
                                "reading_complete": False,
                                "saved_at": datetime.now().isoformat(timespec="seconds"),
                                "candidates": list(candidate_by_uid.values()),
                            })
                            # Checkpoint progressivo su disco ad ogni blocco.
                            atomic_write_json(read_path, read_data)
                            self.ui_queue.put((
                                "extraction_phase_progress",
                                ("read", len(processed_read_uids), read_total, pname, f"candidate: {len(candidate_by_uid)}"),
                            ))
                            self.ui_queue.put((
                                "done",
                                f"Profilo {pname} — Lettura {len(processed_read_uids)}/{len(unique_uids)} — "
                                f"candidate: {len(candidate_by_uid)} — checkpoint salvato",
                            ))

                        if self.extraction_stop_event.is_set():
                            self.ui_queue.put((
                                "extraction_stopped",
                                f"Interruzione durante la Lettura. Checkpoint salvato: "
                                f"{len(processed_read_uids)}/{len(unique_uids)} header, "
                                f"{len(candidate_by_uid)} candidate. Alla prossima esecuzione la Lettura riprenderà da qui.",
                            ))
                            return

                        candidate_headers = [dict_to_header(row) for row in candidate_by_uid.values()]
                        candidate_headers.sort(key=self._mail_chronological_key)
                        candidate_uids = [int(h.uid) for h in candidate_headers]
                        read_data.update({
                            "scanned_count": len(unique_uids),
                            "processed_uids": sorted(processed_read_uids),
                            "candidate_count": len(candidate_headers),
                            "reading_complete": True,
                            "saved_at": datetime.now().isoformat(timespec="seconds"),
                            "candidates": [header_to_dict(h) for h in candidate_headers],
                        })
                        atomic_write_json(read_path, read_data)
                        self.ui_queue.put((
                            "extraction_phase_progress",
                            ("read", read_total, read_total, pname, f"candidate: {len(candidate_headers)}"),
                        ))
                        self.ui_queue.put((
                            "done",
                            f"Profilo {pname} — Lettura completata: {len(unique_uids)} header, "
                            f"{len(candidate_headers)} mail candidate in base ai Tipi mail.",
                        ))

                    # -------------------------------------------------------------
                    # 2. ESTRAZIONE: file JSON persistente aggiornato dopo ogni mail.
                    # -------------------------------------------------------------
                    operations_data = read_json(operations_path)
                    if operations_data and not self._pipeline_cache_matches(operations_data, folder, last_uid):
                        safe_unlink(operations_path)
                        operations_data = None
                    if not operations_data:
                        operations_data = {
                            **self._pipeline_base_metadata(folder, last_uid),
                            "uidvalidity": int(uidv or 0),
                            "scan_max_uid": int(scan_max_uid or 0),
                            "created_at": datetime.now().isoformat(timespec="seconds"),
                            "updated_at": datetime.now().isoformat(timespec="seconds"),
                            "extraction_complete": False,
                            "processed_uids": [],
                            "operations": [],
                            "events": [],
                        }
                        atomic_write_json(operations_path, operations_data)

                    processed = {int(x) for x in operations_data.get("processed_uids", []) or []}
                    remaining_uids = [uid for uid in candidate_uids if uid not in processed]
                    extract_total = max(len(candidate_uids), 1)
                    self.ui_queue.put((
                        "extraction_phase_progress",
                        ("extract", len(processed), extract_total, pname, ""),
                    ))

                    for uid in remaining_uids:
                        if self.extraction_stop_event.is_set():
                            self.ui_queue.put((
                                "extraction_stopped",
                                f"Interruzione durante l'Estrazione. Operazioni salvate: {len(processed)}/{len(candidate_uids)}. "
                                "Alla prossima esecuzione verranno riutilizzati sia la Lettura sia il file operazioni.",
                            ))
                            return

                        mail = self.messages.get(uid) or self._imap_call_with_reconnect(
                            lambda uid=uid: self.imap.fetch_message(uid, folder),
                            f"la lettura della mail UID {uid}",
                        )
                        self.messages[uid] = mail
                        if not int(state.get("first_processed_uid", 0) or 0):
                            state["first_processed_uid"] = int(uid)
                            state["first_processed_at"] = datetime.now().isoformat(timespec="seconds")
                            state["first_mail_date"] = str(getattr(mail, "date", "") or "")
                            self.config_mgr.save()

                        event = extract_mail_detailed(mail, self.profile)
                        serialized = event_to_dict(event, self.profile)

                        # Journal esplicito delle modifiche Excel. Ogni singola
                        # operazione viene persistita atomicamente; in caso di
                        # crash la mail non marcata come processata verrà
                        # rieseguita e il journal di quell'UID viene ricostruito.
                        journal = operations_data.setdefault("operations", [])
                        journal[:] = [x for x in journal if int(x.get("uid", -1) or -1) != int(uid)]
                        for op in serialized.get("operations", []) or []:
                            journal.append({"uid": int(uid), **dict(op)})
                            operations_data["updated_at"] = datetime.now().isoformat(timespec="seconds")
                            atomic_write_json(operations_path, operations_data)

                        operations_data.setdefault("events", []).append(serialized)
                        processed.add(uid)
                        operations_data["processed_uids"] = sorted(processed)
                        operations_data["updated_at"] = datetime.now().isoformat(timespec="seconds")
                        atomic_write_json(operations_path, operations_data)

                        type_names = [str(t.get("mail_type", "")) for t in event.get("types", [])]
                        label = ", ".join(type_names) if type_names else "mail candidata ma non riconosciuta dopo il controllo del body"
                        self.ui_queue.put((
                            "extraction_phase_progress",
                            ("extract", len(processed), extract_total, pname, ""),
                        ))
                        self.ui_queue.put((
                            "done",
                            f"Profilo {pname} — Estrazione {len(processed)}/{len(candidate_uids)} — {label} — operazioni JSON salvate",
                        ))

                    operations_data["extraction_complete"] = True
                    operations_data["completed_at"] = datetime.now().isoformat(timespec="seconds")
                    atomic_write_json(operations_path, operations_data)
                    # A fine Estrazione la sequenza di Lettura non serve più.
                    safe_unlink(read_path)
                    self.ui_queue.put((
                        "extraction_phase_progress",
                        ("extract", extract_total, extract_total, pname, ""),
                    ))

                # -----------------------------------------------------------------
                # 3. CONSOLIDAMENTO transazionale.
                # -----------------------------------------------------------------
                operations_data = read_json(operations_path) or operations_data or {}
                events_data = list(operations_data.get("events", []) or [])
                consolidate_total = max(len(events_data), 1)
                self.ui_queue.put(("extraction_phase_progress", ("consolidate", 0, consolidate_total, pname, "")))

                original_path, temporary_path = prepare_transaction_source(
                    general_path,
                    lambda p: write_excel(p, [], self.profile),
                )
                # Ad ogni ripresa il temporaneo viene sempre ricreato dall'originale.
                safe_unlink(temporary_path)
                shutil.copy2(original_path, temporary_path)

                records = load_existing_records(original_path, self.profile)
                log_lines = [
                    "",
                    "=" * 100,
                    f"INIZIO CONSOLIDAMENTO - {datetime.now().astimezone().isoformat(timespec='seconds')}",
                    f"Profilo: {pname}",
                    f"Originale: {original_path}",
                    f"Temporaneo: {temporary_path}",
                    f"Operazioni/mail: {len(events_data)}",
                ]

                for idx, event_data in enumerate(events_data, 1):
                    if self.extraction_stop_event.is_set():
                        safe_unlink(temporary_path)
                        append_extraction_log(log_path, "\n".join(log_lines + [
                            "CONSOLIDAMENTO INTERROTTO: file temporaneo eliminato; originale preservato.",
                            "=" * 100,
                        ]))
                        self.ui_queue.put((
                            "extraction_stopped",
                            f"Interruzione durante il Consolidamento ({idx-1}/{len(events_data)}). "
                            "Il file temporaneo è stato eliminato. Alla prossima esecuzione il consolidamento "
                            "ripartirà dall'Excel originale.",
                        ))
                        return
                    event = dict_to_event(event_data)
                    records, stats, log_text = apply_detailed_extractions(
                        records, [event], self.profile,
                        mode="consolidamento", excel_path=general_path,
                        include_header=False, include_footer=False,
                    )
                    self._event_stats_add(stats_total, stats)
                    log_lines.append(log_text)
                    self.ui_queue.put((
                        "extraction_phase_progress",
                        ("consolidate", idx, consolidate_total, pname, ""),
                    ))

                # Un solo salvataggio Excel al termine di tutte le operazioni.
                write_excel(temporary_path, records, self.profile)
                # La PivotTable viene ricalcolata una sola volta per sessione, al termine del Consolidamento.
                pivot_ok, pivot_msg = refresh_native_excel_pivot(temporary_path, self.profile)
                if not pivot_ok:
                    raise RuntimeError(pivot_msg)

                # Commit del file: il temporaneo diventa definitivo solo ora.
                temporary_path.replace(general_path)
                # Manteniamo ancora l'originale fino a quando checkpoint/stato e
                # file di lavoro non sono stati committati: così un crash in
                # questa finestra può rifare il Consolidamento senza duplicare dati.
                self._remember_last_excel_path(general_path)

                uidv = int(operations_data.get("uidvalidity", uidv) or uidv)
                scan_max_uid = int(operations_data.get("scan_max_uid", scan_max_uid) or scan_max_uid)
                self._touch_active_extraction_state(folder, uidv, scan_max_uid)
                self._update_active_home_summary_cache_from_excel(general_path)
                self.profile_mgr.save(self.profile_path, self.profile)

                safe_unlink(operations_path)
                safe_unlink(read_path)
                self._clear_extraction_checkpoint()
                safe_unlink(original_path)

                append_extraction_log(log_path, "\n".join(log_lines + [
                    f"Tabella totali: {pivot_msg}",
                    f"FINE CONSOLIDAMENTO - {datetime.now().astimezone().isoformat(timespec='seconds')}",
                    "=" * 100,
                ]))
                self.ui_queue.put(("extraction_phase_progress", ("consolidate", consolidate_total, consolidate_total, pname, "")))
                self.ui_queue.put(("extraction_done", {
                    "path": str(general_path),
                    "message": (
                        f"File generale consolidato:\\n{general_path}\\n\\n"
                        f"Mail candidate estratte: {len(events_data)}\\n"
                        f"Mail con match: {stats_total['matched_mails']}\\n"
                        f"Campi aggiornati: {stats_total['field_updates']}\\n"
                        f"Ultimo UID di scansione memorizzato: {scan_max_uid}\\n"
                        f"Log: {log_path}"
                    ),
                }))
            except ExtractionRetryCancelled as exc:
                self.ui_queue.put((
                    "extraction_stopped",
                    f"Connessione email non disponibile. La pipeline è stata sospesa senza perdere i file di lavoro.\\n\\n{exc}",
                ))
            except Exception as exc:
                # Se il consolidamento aveva creato un temporaneo, non deve mai
                # diventare il file definitivo in caso di errore.
                try:
                    _orig, _tmp = transactional_paths(general_path)
                    safe_unlink(_tmp)
                except Exception:
                    pass
                self.ui_queue.put(("extraction_error", exc))

        threading.Thread(target=worker, daemon=True).start()

    def _on_active_imap_profile_changed(self, *, clear_search: bool = True):
        """Aggiorna la GUI dopo un cambio account IMAP.

        Le cache vengono svuotate perché UID identici su account differenti non
        identificano la stessa mail e non devono mai essere mescolati.
        """
        self.folder_var.set(self.config_mgr.imap.folder or "INBOX")
        name = self.config_mgr.active_imap_profile_name or "<nessuna email>"
        self.imap_profile_var.set(name)
        if clear_search:
            self.stop_event.set()
            self.headers.clear()
            self.messages.clear()
            self.row_matches.clear()
            self.search_order.clear()
            self.selected_uids.clear()
            self.current_uid = None
            self.current_mail = None
            self.filter_mail_types.clear()
            try:
                self._rebuild_tree_rows()
                self.text.delete("1.0", "end")
                self.extraction_preview_var.set("Anteprima estrazione: apri una mail")
            except Exception:
                pass
        self.status_var.set(f"Email attivo: {name}")
        self._refresh_home_profiles()

    def open_imap_settings(self):
        creating = not bool(self.config_mgr.active_imap_profile_id)
        dlg = ImapSettingsDialog(
            self, self.config_mgr,
            self.config_mgr.active_imap_profile_id or None,
            creating=creating,
        )
        self.wait_window(dlg)
        if dlg.result:
            self._on_active_imap_profile_changed(clear_search=True)

    def choose_imap_profile(self):
        profiles = self.config_mgr.list_imap_profiles()
        if not IS_PRO and profiles:
            self.config_mgr.set_active_imap_profile(str(profiles[0]["id"]))
            self._on_active_imap_profile_changed(clear_search=True)
            return
        if not profiles:
            dlg = ImapSettingsDialog(self, self.config_mgr, creating=True)
            self.wait_window(dlg)
            if dlg.result:
                self._on_active_imap_profile_changed(clear_search=True)
            return
        if len(profiles) == 1:
            self.config_mgr.set_active_imap_profile(str(profiles[0]["id"]))
            self._on_active_imap_profile_changed(clear_search=True)
            return
        chooser = ImapProfileChooser(self, self.config_mgr, title="Cambia email")
        self.wait_window(chooser)
        if chooser.result:
            self.config_mgr.set_active_imap_profile(chooser.result)
            self._on_active_imap_profile_changed(clear_search=True)

    def manage_imap_profiles(self):
        if not IS_PRO:
            messagebox.showinfo("EgoMailExtractor Community", "La versione Community utilizza una sola email.", parent=self)
            return
        dlg = ImapProfilesDialog(
            self, self.config_mgr,
            on_active_changed=lambda: self._on_active_imap_profile_changed(clear_search=True),
        )
        self.wait_window(dlg)
        self._on_active_imap_profile_changed(clear_search=False)

    def reconnect_imap(self):
        self.status_var.set("Verifica connessione email…")
        folder = self.folder_var.get().strip() or "INBOX"
        def worker():
            ok, msg = self.imap.test_connection(folder)
            self.ui_queue.put(("done", msg if ok else f"Errore email: {msg}"))
            if not ok:
                self.ui_queue.put(("error", msg))
        threading.Thread(target=worker, daemon=True).start()

    def change_master_password(self):
        dlg = MasterPasswordDialog(self, creating=True)
        dlg.title("Cambia password master")
        self.wait_window(dlg)
        if not dlg.result:
            return
        try:
            self.config_mgr.change_master_password(dlg.result)
            messagebox.showinfo(
                "Password master",
                "Password master aggiornata. Il file delle credenziali email è stato ricifrato.",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror("Errore", str(e), parent=self)

    def logout_change_account(self):
        """Compatibilità col vecchio comando: apre ora la gestione multi-account."""
        self.manage_imap_profiles()

    def preferences(self):
        win = tk.Toplevel(self)
        win.title("Preferenze")
        win.resizable(True, True)
        win.geometry("860x540")
        win.minsize(700, 430)

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)
        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        # --- Generali ---
        general = ttk.Frame(notebook, padding=12)
        notebook.add(general, text="Generali")
        general.columnconfigure(0, weight=1)
        pvar = tk.StringVar(value=str(self.config_mgr.profiles_dir))
        ttk.Label(general, text="Directory profili").grid(row=0, column=0, sticky="w")
        path_row = ttk.Frame(general)
        path_row.grid(row=1, column=0, sticky="ew", pady=5)
        path_row.columnconfigure(0, weight=1)
        ttk.Entry(path_row, textvariable=pvar).grid(row=0, column=0, sticky="ew")

        def browse():
            p = filedialog.askdirectory(parent=win, initialdir=pvar.get())
            if p:
                pvar.set(p)

        ttk.Button(path_row, text="Sfoglia…", command=browse).grid(row=0, column=1, padx=(6, 0))

        wvar = tk.StringVar(value=str(self.config_mgr.work_dir))
        ttk.Label(general, text="Cartella degli Excel").grid(row=2, column=0, sticky="w", pady=(10,0))
        work_row = ttk.Frame(general)
        work_row.grid(row=3, column=0, sticky="ew", pady=5)
        work_row.columnconfigure(0, weight=1)
        ttk.Entry(work_row, textvariable=wvar).grid(row=0, column=0, sticky="ew")
        def browse_work():
            p = filedialog.askdirectory(parent=win, initialdir=wvar.get())
            if p:
                wvar.set(p)
        ttk.Button(work_row, text="Sfoglia…", command=browse_work).grid(row=0, column=1, padx=(6, 0))
        ttk.Label(
            general,
            text="In questa cartella vengono creati e aggiornati gli Excel generali dei profili di estrazione.",
            foreground="#666666", wraplength=760,
        ).grid(row=4, column=0, sticky="w", pady=(0,6))

        evar = tk.StringVar(value=str(self.config_mgr.email_dir))
        ttk.Label(general, text="Cartella email").grid(row=5, column=0, sticky="w", pady=(10,0))
        email_row = ttk.Frame(general)
        email_row.grid(row=6, column=0, sticky="ew", pady=5)
        email_row.columnconfigure(0, weight=1)
        ttk.Entry(email_row, textvariable=evar).grid(row=0, column=0, sticky="ew")
        def browse_email_dir():
            p = filedialog.askdirectory(parent=win, initialdir=evar.get())
            if p: evar.set(p)
        ttk.Button(email_row, text="Sfoglia…", command=browse_email_dir).grid(row=0, column=1, padx=(6, 0))
        ttk.Label(general, text="In questa cartella viene conservata la configurazione cifrata delle email.", foreground="#666666").grid(row=7,column=0,sticky="w")

        imap_pref = ttk.Frame(general)
        imap_pref.grid(row=8, column=0, sticky="ew", pady=10)
        ttk.Label(imap_pref, text="Email attivo:").pack(side="left")
        ttk.Label(imap_pref, textvariable=self.imap_profile_var, font=("TkDefaultFont", 9, "bold")).pack(side="left", padx=(5, 12))
        ttk.Button(
            imap_pref, text="Gestisci email…",
            command=self.manage_imap_profiles,
        ).pack(side="left")

        # --- Abbreviazioni Regex ---
        regex_tab = ttk.Frame(notebook, padding=12)
        notebook.add(regex_tab, text="Abbreviazioni Regex")
        regex_tab.rowconfigure(1, weight=1)
        regex_tab.columnconfigure(0, weight=1)
        ttk.Label(
            regex_tab,
            text=(
                "Le abbreviazioni vengono sostituite letteralmente prima che una regex sia compilata. "
                "Esempio: --ALFANUMERICO--  →  [A-Za-z0-9]+. "
                "Sono applicate alle regex di selezione, alle regex di estrazione e alle regex sul valore."
            ),
            wraplength=780, justify="left",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))

        table_holder = ttk.Frame(regex_tab)
        table_holder.grid(row=1, column=0, sticky="nsew")
        table_holder.rowconfigure(0, weight=1)
        table_holder.columnconfigure(0, weight=1)
        abbr_tree = ttk.Treeview(
            table_holder, columns=("token", "replacement"), show="headings", selectmode="browse"
        )
        abbr_tree.heading("token", text="Abbreviazione")
        abbr_tree.heading("replacement", text="Sostituzione Regex")
        abbr_tree.column("token", width=220, stretch=False)
        abbr_tree.column("replacement", width=500, stretch=True)
        abbr_y = ttk.Scrollbar(table_holder, orient="vertical", command=abbr_tree.yview)
        abbr_x = ttk.Scrollbar(table_holder, orient="horizontal", command=abbr_tree.xview)
        abbr_tree.configure(yscrollcommand=abbr_y.set, xscrollcommand=abbr_x.set)
        abbr_tree.grid(row=0, column=0, sticky="nsew")
        abbr_y.grid(row=0, column=1, sticky="ns")
        abbr_x.grid(row=1, column=0, sticky="ew")

        abbreviations = [dict(x) for x in self.config_mgr.regex_abbreviations]

        def refresh_abbr_tree():
            for iid in abbr_tree.get_children():
                abbr_tree.delete(iid)
            for idx, item in enumerate(abbreviations):
                abbr_tree.insert(
                    "", "end", iid=str(idx),
                    values=(item.get("token", ""), item.get("replacement", "")),
                )

        def selected_index() -> int | None:
            sel = abbr_tree.selection()
            if not sel:
                return None
            try:
                return int(sel[0])
            except Exception:
                return None

        def edit_abbreviation(index: int | None = None):
            old = abbreviations[index] if index is not None and 0 <= index < len(abbreviations) else {}
            token = simpledialog.askstring(
                "Abbreviazione Regex", "Stringa abbreviazione:",
                initialvalue=str(old.get("token", "--ALFANUMERICO--")), parent=win,
            )
            if token is None:
                return
            token = token.strip()
            if not token:
                messagebox.showwarning("Abbreviazione Regex", "L'abbreviazione non può essere vuota.", parent=win)
                return
            duplicate = next((i for i, x in enumerate(abbreviations) if x.get("token") == token and i != index), None)
            if duplicate is not None:
                messagebox.showwarning("Abbreviazione Regex", "Esiste già una riga con questa abbreviazione.", parent=win)
                return
            replacement = simpledialog.askstring(
                "Abbreviazione Regex", "Stringa Regex che deve sostituirla:",
                initialvalue=str(old.get("replacement", "[A-Za-z0-9]+")), parent=win,
            )
            if replacement is None:
                return
            item = {"token": token, "replacement": replacement}
            if index is None:
                abbreviations.append(item)
            else:
                abbreviations[index] = item
            refresh_abbr_tree()

        def delete_abbreviation():
            idx = selected_index()
            if idx is None:
                return
            token = abbreviations[idx].get("token", "")
            if messagebox.askyesno("Abbreviazioni Regex", f"Eliminare «{token}»?", parent=win):
                del abbreviations[idx]
                refresh_abbr_tree()

        buttons = ttk.Frame(regex_tab)
        buttons.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(buttons, text="Aggiungi…", command=lambda: edit_abbreviation(None)).pack(side="left")
        ttk.Button(buttons, text="Modifica…", command=lambda: edit_abbreviation(selected_index())).pack(side="left", padx=5)
        ttk.Button(buttons, text="Elimina", command=delete_abbreviation).pack(side="left")
        abbr_tree.bind("<Double-1>", lambda _e: edit_abbreviation(selected_index()))
        refresh_abbr_tree()

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10, 0))

        def save():
            self.config_mgr.data["profiles_dir"] = str(Path(pvar.get()))
            self.config_mgr.data["work_dir"] = str(Path(wvar.get()))
            Path(wvar.get()).mkdir(parents=True, exist_ok=True)
            self.config_mgr.set_email_dir(Path(evar.get()))
            self.config_mgr.data["regex_abbreviations"] = normalize_regex_abbreviations(abbreviations)
            self.config_mgr.ensure_profiles_dir()
            self.config_mgr.save()
            set_regex_abbreviations(self.config_mgr.regex_abbreviations)
            self._recompute_cached_matches()
            self._update_extraction_preview()
            self.status_var.set(
                f"Preferenze salvate — {len(self.config_mgr.regex_abbreviations)} abbreviazioni Regex attive"
            )
            messagebox.showinfo(
                "Preferenze",
                "Preferenze salvate. Le abbreviazioni Regex sono già attive. "
                "La nuova directory profili sarà usata al prossimo cambio profilo/avvio. "
                "La cartella degli Excel e la cartella email sono già attive.",
                parent=win,
            )
            win.destroy()

        ttk.Button(footer, text="Annulla", command=win.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(footer, text="Salva", command=save).pack(side="right")
        win.transient(self)
        win.grab_set()
        center_dialog_later(win, self)

    def _load_profile_path(self, path: Path | str):
        """Attiva un profilo e riallinea tutte le cache della schermata principale."""
        mgr = ProfileManager(self.config_mgr.profiles_dir)
        path = Path(path)
        self.profile_path = path
        self.profile_mgr = mgr
        self.profile = mgr.load(path)
        self.profile_mgr.save(path, self.profile)
        self.title(f"EgoMailExtractor — {self.profile.get('name', path.stem)}")
        self.headers.clear()
        self.messages.clear()
        self.row_matches.clear()
        self.search_order.clear()
        self.selected_uids.clear()
        self.current_uid = None
        self.current_mail = None
        self.filter_mail_types.clear()
        self.type_filter_var.set("Filtro tipi mail: tutti")
        self.text.delete("1.0", "end")
        self.extraction_preview_var.set("Anteprima estrazione: apri una mail")
        self._refresh_mail_types()
        self._refresh_home_profiles()

    def show_pro_feature(self, feature_name="Funzionalità Pro"):
        win = tk.Toplevel(self)
        win.title(feature_name)
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=18); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Per questa funzionalità scaricare la versione Pro", font=("TkDefaultFont", 11, "bold")).pack(pady=(0,10))
        link = tk.Label(frm, text="Vai alla pagina di EgoMailExtractor Pro", fg="#0563C1", cursor="hand2", font=("TkDefaultFont", 10, "underline"))
        link.pack(pady=4)
        link.bind("<Button-1>", lambda _e: webbrowser.open("https://www.domoticachepassione.it/wp/acquista-egomailextractor-pro/"))
        ttk.Button(frm, text="Chiudi", command=win.destroy).pack(pady=(14,0))
        win.transient(self); win.grab_set(); center_dialog_later(win, self)

    def configure_totals_table(self):
        if not IS_PRO:
            return self.show_pro_feature("Tabella totali")
        fields = [f if isinstance(f,str) else str(f.get("name", "")) for f in self.profile.get("fields", [])]
        fields = [f for f in fields if f]
        fmap = {str(f.get("name")): f for f in self.profile.get("fields", []) if isinstance(f,dict) and f.get("name")}
        numeric = [n for n in fields if str(fmap.get(n,{}).get("type","")).lower() == "number"]
        dates = [n for n in fields if str(fmap.get(n,{}).get("type","")).lower() == "date"]
        cfg = dict(self.profile.get("totals_table", {}) or {})
        win=tk.Toplevel(self); win.title("Tabella totali"); win.geometry("720x650")
        frm=ttk.Frame(win,padding=12); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Seleziona i campi della tabella totali",font=("TkDefaultFont",11,"bold")).pack(anchor="w")
        panes=ttk.Frame(frm); panes.pack(fill="both",expand=True,pady=8)
        row_vars={}; col_vars={}; val_vars={}
        def box(parent,title,names,selected,store):
            lf=ttk.LabelFrame(parent,text=title,padding=8); lf.pack(side="left",fill="both",expand=True,padx=4)
            canvas=tk.Canvas(lf,highlightthickness=0); sb=ttk.Scrollbar(lf,orient="vertical",command=canvas.yview); inner=ttk.Frame(canvas)
            inner.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all"))); canvas.create_window((0,0),window=inner,anchor="nw"); canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")
            for n in names:
                v=tk.BooleanVar(value=n in selected); store[n]=v; ttk.Checkbutton(inner,text=n,variable=v).pack(anchor="w")
        box(panes,"Etichette di riga",fields,set(cfg.get("row_fields",[])),row_vars)
        box(panes,"Etichette di colonna",fields,set(cfg.get("column_fields",[])),col_vars)
        box(panes,"Valori da sommare",numeric,set(cfg.get("value_fields",[])),val_vars)
        datefrm=ttk.LabelFrame(frm,text="Raggruppamento data",padding=8); datefrm.pack(fill="x",pady=6)
        date_var=tk.StringVar(value=str(cfg.get("date_field","") or "")); year_var=tk.BooleanVar(value=bool(cfg.get("group_year"))); month_var=tk.BooleanVar(value=bool(cfg.get("group_month")))
        ttk.Label(datefrm,text="Campo data:").pack(side="left"); ttk.Combobox(datefrm,textvariable=date_var,values=[""]+dates,state="readonly",width=28).pack(side="left",padx=6)
        ttk.Checkbutton(datefrm,text="Separa per anno",variable=year_var).pack(side="left",padx=8); ttk.Checkbutton(datefrm,text="Separa per mese",variable=month_var).pack(side="left",padx=8)
        def save():
            self.profile["totals_table"]={"row_fields":[n for n,v in row_vars.items() if v.get()],"column_fields":[n for n,v in col_vars.items() if v.get()],"value_fields":[n for n,v in val_vars.items() if v.get()],"date_field":date_var.get().strip(),"group_year":year_var.get(),"group_month":month_var.get()}
            self.save_profile(); win.destroy()
        btn=ttk.Frame(frm); btn.pack(fill="x",pady=(8,0)); ttk.Button(btn,text="Salva",command=save).pack(side="right",padx=4); ttk.Button(btn,text="Annulla",command=win.destroy).pack(side="right",padx=4)
        win.transient(self); win.grab_set(); center_dialog_later(win,self)

    def open_guided_wizard(self):
        if not IS_PRO:
            messagebox.showinfo(
                "EgoMailExtractor Community",
                "Il Wizard visuale è disponibile in EgoMailExtractor Pro.\n\n"
                "Nella Community puoi creare e modificare liberamente regole, campi e tipi mail dalla schermata completa.",
                parent=self,
            )
            return
        from .wizard import GuidedExtractionWizard
        GuidedExtractionWizard(self)

    def change_profile(self):
        if not IS_PRO:
            messagebox.showinfo(
                "EgoMailExtractor Community",
                "La versione Community utilizza un solo profilo di estrazione.",
                parent=self,
            )
            return
        mgr = ProfileManager(self.config_mgr.profiles_dir)
        paths = self._associated_profile_paths()
        chooser = ProfileChooser(self, mgr, paths=paths, associated_only=True, config=self.config_mgr)
        self.wait_window(chooser)
        if chooser.result:
            self._load_profile_path(chooser.result)

    def manage_email_profile_associations(self):
        if not self.config_mgr.active_imap_profile_id:
            messagebox.showwarning("Profili associati", "Seleziona prima una email attiva.", parent=self)
            return
        dlg = EmailExtractionAssociationsDialog(self, self.config_mgr, ProfileManager(self.config_mgr.profiles_dir))
        self.wait_window(dlg)
        self._refresh_home_profiles()

    def export_extraction_profiles(self):
        profiles = [Path(self.profile_path)] if not IS_PRO else self.profile_mgr.list_profiles()
        if not profiles:
            messagebox.showinfo("Export profilo di estrazione", "Non ci sono profili di estrazione da esportare.", parent=self)
            return

        win = tk.Toplevel(self)
        win.title("Export profilo di estrazione")
        win.transient(self)
        win.grab_set()
        win.geometry("560x500")
        center_dialog_later(win, self)
        ttk.Label(win, text="Seleziona i profili da esportare", font=("TkDefaultFont", 11, "bold")).pack(anchor="w", padx=14, pady=(14, 6))
        ttk.Label(win, text="Verrà creato uno ZIP separato per ogni profilo selezionato.").pack(anchor="w", padx=14, pady=(0, 8))

        outer = ttk.Frame(win)
        outer.pack(fill="both", expand=True, padx=14, pady=4)
        canvas = tk.Canvas(outer, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas)
        body.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        vars_by_path = []
        for path in profiles:
            try:
                prof = self.profile_mgr.load(path)
                name = str(prof.get("name") or path.stem)
            except Exception:
                name = path.stem
            var = tk.BooleanVar(value=(Path(path) == Path(self.profile_path)))
            vars_by_path.append((path, var))
            ttk.Checkbutton(body, text=name, variable=var).pack(anchor="w", fill="x", padx=4, pady=3)

        result = {"ok": False}
        def do_export():
            selected = [path for path, var in vars_by_path if var.get()]
            if not selected:
                messagebox.showwarning("Export profilo di estrazione", "Seleziona almeno un profilo.", parent=win)
                return
            folder = filedialog.askdirectory(title="Cartella in cui salvare i profili esportati", parent=win)
            if not folder:
                return
            exported = []
            errors = []
            for path in selected:
                try:
                    exported.append(export_profile_package(path, folder, self.config_mgr.regex_abbreviations))
                except Exception as exc:
                    errors.append(f"{Path(path).name}: {exc}")
            if exported:
                msg = f"Creati {len(exported)} pacchetti ZIP in:\n{folder}"
                if errors:
                    msg += "\n\nErrori:\n" + "\n".join(errors)
                messagebox.showinfo("Export completato", msg, parent=win)
                result["ok"] = True
                win.destroy()
            elif errors:
                messagebox.showerror("Export profilo di estrazione", "\n".join(errors), parent=win)

        buttons = ttk.Frame(win)
        buttons.pack(fill="x", padx=14, pady=14)
        ttk.Button(buttons, text="Esporta selezionati…", command=do_export).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="Annulla", command=win.destroy).pack(side="right")

    def import_extraction_profiles(self):
        paths = filedialog.askopenfilenames(
            title="Importa profilo di estrazione", parent=self,
            filetypes=[("Pacchetti profilo EgoMailExtractor", "*.zip"), ("Tutti i file", "*.*")],
        )
        if not paths: return
        imported_paths=[]; imported_abbreviations=[]; errors=[]
        for zip_path in paths:
            try:
                profile, _abbr, _member = inspect_profile_package(zip_path)
                requested_name = str(profile.get("name") or "profilo").strip()
                existing = None
                for pp in self.profile_mgr.list_profiles():
                    try:
                        if str(self.profile_mgr.load(pp).get("name") or "").casefold() == requested_name.casefold():
                            existing = Path(pp); break
                    except Exception: pass
                target_name = None
                if existing is not None:
                    if not messagebox.askyesno(
                        "Profilo già esistente",
                        f"Esiste già il profilo «{requested_name}».\n\nVuoi importarlo con un nome differente?",
                        parent=self,
                    ):
                        continue
                    while True:
                        target_name = simpledialog.askstring(
                            "Nuovo nome profilo", "Nome da assegnare al profilo importato:",
                            initialvalue=requested_name + " - importato", parent=self,
                        )
                        if target_name is None: break
                        target_name = target_name.strip()
                        if not target_name: continue
                        candidate = self.profile_mgr.profiles_dir / f"{slugify(target_name,'profilo')}.json"
                        if candidate.exists():
                            messagebox.showwarning("Import profilo", "Esiste già un profilo con questo nome.", parent=self); continue
                        break
                    if target_name is None: continue
                dest, abbreviations = import_profile_package(zip_path, self.profile_mgr, target_name=target_name)
                imported_paths.append(dest); imported_abbreviations.extend(abbreviations)
            except Exception as exc:
                errors.append(f"{Path(zip_path).name}: {exc}")
        if imported_abbreviations:
            merged=merge_regex_abbreviations(self.config_mgr.regex_abbreviations, imported_abbreviations)
            self.config_mgr.set_regex_abbreviations(merged); self.config_mgr.save(); set_regex_abbreviations(self.config_mgr.regex_abbreviations)
        if imported_paths:
            names=[]
            for pp in imported_paths:
                try: names.append(str(self.profile_mgr.load(pp).get("name") or pp.stem))
                except Exception: names.append(pp.stem)
            msg=f"Importati {len(imported_paths)} profili di estrazione:\n\n" + "\n".join(f"• {n}" for n in names)
            if errors: msg += "\n\nErrori:\n" + "\n".join(errors)
            messagebox.showinfo("Import profilo di estrazione", msg, parent=self)
            self._refresh_home_profile_buttons()
        elif errors:
            messagebox.showerror("Import profilo di estrazione", "Nessun profilo importato.\n\n"+"\n".join(errors), parent=self)

    def _profile_excel_files_for_backup(self, profile_path: Path, profile: dict) -> list[tuple[str, Path]]:
        out=[]
        pname=safe_filename_component(str(profile.get("name", profile_path.stem) or profile_path.stem), profile_path.stem)
        for email in self.config_mgr.list_imap_profiles():
            ename=safe_filename_component(str(email.get("name") or "Email"), "Email")
            path=self.config_mgr.work_dir / f"{pname} - {ename}.xlsx"
            if path.exists(): out.append((path.name,path))
        return out

    def _choose_profiles_for_action(self, title: str, prompt: str, preselected: set[str] | None = None):
        profiles=self.profile_mgr.list_profiles()
        if not profiles:
            messagebox.showinfo(title,"Non ci sono profili di estrazione disponibili.",parent=self); return None
        win=tk.Toplevel(self); win.title(title); win.transient(self); win.grab_set(); win.geometry("580x500"); center_dialog_later(win,self)
        ttk.Label(win,text=prompt,font=("TkDefaultFont",11,"bold")).pack(anchor="w",padx=14,pady=(14,8))
        outer=ttk.Frame(win); outer.pack(fill="both",expand=True,padx=14,pady=4)
        canvas=tk.Canvas(outer,highlightthickness=0); sb=ttk.Scrollbar(outer,orient="vertical",command=canvas.yview); body=ttk.Frame(canvas)
        body.bind("<Configure>",lambda _e:canvas.configure(scrollregion=canvas.bbox("all"))); canvas.create_window((0,0),window=body,anchor="nw"); canvas.configure(yscrollcommand=sb.set); canvas.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")
        items=[]; preselected=preselected or set()
        for pp in profiles:
            try: name=str(self.profile_mgr.load(pp).get("name") or pp.stem)
            except Exception: name=pp.stem
            v=tk.BooleanVar(value=(name in preselected or Path(pp)==Path(self.profile_path)))
            items.append((Path(pp),name,v)); ttk.Checkbutton(body,text=name,variable=v).pack(anchor="w",fill="x",padx=4,pady=3)
        result={"paths":None}
        def ok():
            sel=[p for p,n,v in items if v.get()]
            if not sel: messagebox.showwarning(title,"Seleziona almeno un profilo.",parent=win); return
            result["paths"]=sel; win.destroy()
        bar=ttk.Frame(win); bar.pack(fill="x",padx=14,pady=14); ttk.Button(bar,text="Continua",command=ok).pack(side="right",padx=(8,0)); ttk.Button(bar,text="Annulla",command=win.destroy).pack(side="right")
        self.wait_window(win); return result["paths"]

    def backup_extraction_profile(self):
        selected = self._choose_profiles_for_action(
            "Backup profili di estrazione",
            "Seleziona i profili da includere nel backup",
        )
        if not selected:
            return
        folder = filedialog.askdirectory(title="Cartella in cui salvare il backup", parent=self)
        if not folder:
            return
        try:
            payload = []
            for pp in selected:
                prof = self.profile_mgr.load(pp)
                states = self.config_mgr.all_extraction_states_for_profile(Path(pp).name)
                payload.append((pp, states, self._profile_excel_files_for_backup(Path(pp), prof)))
            out = export_profiles_backup(payload, folder, self.config_mgr.regex_abbreviations)
            messagebox.showinfo(
                "Backup profili",
                f"Backup creato con {len(selected)} profili:\n{out}",
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Backup profili", str(exc), parent=self)

    def restore_extraction_profile(self):
        zp = filedialog.askopenfilename(
            title="Restore profili di estrazione",
            parent=self,
            filetypes=[("Backup EgoMailExtractor", "*.zip"), ("Tutti i file", "*.*")],
        )
        if not zp:
            return
        try:
            entries, abbreviations = inspect_profiles_backup(zp)
            if not entries:
                raise ValueError("Il backup non contiene profili ripristinabili")

            win = tk.Toplevel(self)
            win.title("Restore profili di estrazione")
            win.transient(self)
            win.grab_set()
            win.geometry("600x500")
            center_dialog_later(win, self)
            ttk.Label(
                win,
                text="Seleziona i profili da ripristinare",
                font=("TkDefaultFont", 11, "bold"),
            ).pack(anchor="w", padx=14, pady=(14, 5))
            ttk.Label(win, text=f"Backup: {Path(zp).name}").pack(anchor="w", padx=14, pady=(0, 8))

            outer = ttk.Frame(win)
            outer.pack(fill="both", expand=True, padx=14)
            canvas = tk.Canvas(outer, highlightthickness=0)
            sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            body = ttk.Frame(canvas)
            body.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=body, anchor="nw")
            canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            variables = []
            for entry in entries:
                v = tk.BooleanVar(value=True)
                variables.append((entry, v))
                ttk.Checkbutton(body, text=str(entry.get("name") or "profilo"), variable=v).pack(
                    anchor="w", fill="x", padx=4, pady=3
                )

            result = {"entries": None}

            def continue_restore():
                selected_entries = [entry for entry, var in variables if var.get()]
                if not selected_entries:
                    messagebox.showwarning("Restore", "Seleziona almeno un profilo.", parent=win)
                    return
                result["entries"] = selected_entries
                win.destroy()

            buttons = ttk.Frame(win)
            buttons.pack(fill="x", padx=14, pady=14)
            ttk.Button(buttons, text="Ripristina selezionati…", command=continue_restore).pack(
                side="right", padx=(8, 0)
            )
            ttk.Button(buttons, text="Annulla", command=win.destroy).pack(side="right")
            self.wait_window(win)
            selected_entries = result["entries"]
            if not selected_entries:
                return

            names = ", ".join(str(e.get("name") or "profilo") for e in selected_entries)
            warning = (
                "Il restore sovrascriverà, per i profili selezionati, la definizione del profilo, "
                "l'icona, le abbreviazioni Regex, gli stati di estrazione associati alle email "
                "e gli Excel generali presenti nel backup.\n\n"
                f"Profili: {names}\n\n"
                "I dati sostituiti non saranno recuperabili automaticamente. Continuare?"
            )
            if not messagebox.askyesno("Conferma restore", warning, parent=self):
                return

            import zipfile

            restored = []
            with zipfile.ZipFile(zp, "r") as zf:
                for entry in selected_entries:
                    profile, states = read_profile_from_backup(zp, entry)
                    name = str(profile.get("name") or "profilo")
                    dest = None
                    for pp in self.profile_mgr.list_profiles():
                        try:
                            if str(self.profile_mgr.load(pp).get("name") or "").casefold() == name.casefold():
                                dest = Path(pp)
                                break
                        except Exception:
                            pass
                    if dest is None:
                        dest = self.profile_mgr.profiles_dir / f"{slugify(name, 'profilo')}.json"

                    icon_member = str(entry.get("icon_member") or "")
                    if icon_member and icon_member in zf.namelist():
                        suffix = Path(icon_member).suffix.lower()
                        icon_dest = dest.with_name(dest.stem + "_icon" + suffix)
                        icon_dest.write_bytes(zf.read(icon_member))
                        profile["icon"] = icon_dest.name

                    self.profile_mgr.save(dest, profile)
                    self.config_mgr.set_all_extraction_states_for_profile(dest.name, states)
                    for member in entry.get("excel_members", []):
                        if member in zf.namelist():
                            target = self.config_mgr.work_dir / Path(member).name
                            target.parent.mkdir(parents=True, exist_ok=True)
                            target.write_bytes(zf.read(member))
                    restored.append(name)

            self.config_mgr.data["regex_abbreviations"] = normalize_regex_abbreviations(abbreviations)
            self.config_mgr.save()
            set_regex_abbreviations(self.config_mgr.regex_abbreviations)
            self._refresh_home_profile_buttons()
            messagebox.showinfo(
                "Restore profili",
                f"Ripristinati {len(restored)} profili:\n" + "\n".join(restored),
                parent=self,
            )
        except Exception as exc:
            messagebox.showerror("Restore profili", f"Restore non riuscito:\n{exc}", parent=self)

    def _rename_profile_excel_files(self, old_display: str, new_display: str) -> None:
        for email in self.config_mgr.list_imap_profiles():
            ename=safe_filename_component(str(email.get("name") or "Email"), "Email")
            old=self.config_mgr.work_dir / f"{safe_filename_component(old_display,'profilo')} - {ename}.xlsx"
            new=self.config_mgr.work_dir / f"{safe_filename_component(new_display,'profilo')} - {ename}.xlsx"
            if old.exists() and old.resolve()!=new.resolve() and not new.exists():
                try: old.replace(new)
                except Exception: pass

    def rename_profile(self):
        old_path = Path(self.profile_path)
        old_name = str(self.profile.get("name", old_path.stem) or old_path.stem)
        new_name = simpledialog.askstring("Rinomina profilo", "Nuovo nome del profilo:", initialvalue=old_name, parent=self)
        if not new_name or not new_name.strip() or new_name.strip() == old_name:
            return
        new_name = new_name.strip()
        new_stem = slugify(new_name, old_path.stem)
        new_path = old_path.with_name(new_stem + old_path.suffix)
        if new_path != old_path and new_path.exists():
            messagebox.showerror("Rinomina profilo", f"Esiste già un profilo con file {new_path.name}.", parent=self)
            return
        self.profile["name"] = new_name
        self.profile_mgr.save(old_path, self.profile)
        if new_path != old_path:
            old_path.replace(new_path)
            self.config_mgr.rename_associated_profile(old_path.name, new_path.name)
            self.config_mgr.rename_extraction_profile_state(old_path.name, new_path.name)
        self._rename_profile_excel_files(old_name, new_name)
        self.profile_path = new_path
        self.profile = self.profile_mgr.load(new_path)
        self.title(f"EgoMailExtractor — {new_name}")
        self._refresh_home_profiles()
        self.status_var.set(f"Profilo rinominato: {new_name}")

    def set_extraction_start_date(self):
        state = self._active_extraction_state()
        current = str(state.get("extraction_start_date", "") or "")
        initial = ""
        if current:
            try: initial = datetime.fromisoformat(current).strftime("%d/%m/%Y")
            except Exception: initial = current
        value = simpledialog.askstring(
            "Data inizio estrazione",
            "Inserisci la data da cui iniziare l'estrazione generale (GG/MM/AAAA).\nLascia vuoto per rimuovere il limite.",
            initialvalue=initial, parent=self,
        )
        if value is None:
            return
        value=value.strip()
        if not value:
            state["extraction_start_date"] = ""
        else:
            try:
                dt=datetime.strptime(value, "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Data inizio estrazione", "Data non valida. Usa GG/MM/AAAA.", parent=self); return
            state["extraction_start_date"] = dt.date().isoformat()
        state = self._active_extraction_state()
        state.update({"uidvalidity": 0, "last_processed_uid": 0, "last_processed_at": ""})
        self.config_mgr.save()
        self._clear_extraction_checkpoint()
        self.status_var.set("Data inizio estrazione aggiornata; alla prossima esecuzione la scansione ripartirà dalla data indicata")


    def rebuild_general_extraction(self):
        """Azzera lo stato della sola coppia email+profilo e ricostruisce l'Excel generale."""
        email_name = self.config_mgr.active_imap_profile_name or "Email attiva"
        profile_name = str(self.profile.get("name", Path(self.profile_path).stem) or Path(self.profile_path).stem)
        start_raw = str(self._active_extraction_state().get("extraction_start_date", "") or "").strip()
        start_label = "inizio casella"
        if start_raw:
            try: start_label = datetime.fromisoformat(start_raw).strftime("%d/%m/%Y")
            except Exception: start_label = start_raw
        if not messagebox.askyesno(
            "Riesegui estrazione generale",
            f"Verrà rieseguita l'estrazione del profilo\n{profile_name}\n\ndalla email\n{email_name}\n\na partire dal\n{start_label}\n\nL'Excel generale verrà ricostruito. Continuare?",
            parent=self,
        ):
            return
        if self.extraction_running:
            messagebox.showwarning("Riesegui estrazione generale", "È già in corso un'estrazione.", parent=self)
            return
        general_path = self._general_excel_path()
        try:
            if general_path.exists(): general_path.unlink()
        except Exception as exc:
            messagebox.showerror("Riesegui estrazione generale", f"Impossibile eliminare l'Excel generale esistente:\n{general_path}\n\n{exc}", parent=self)
            return
        state = self._active_extraction_state()
        state.update({"uidvalidity": 0, "last_processed_uid": 0, "last_processed_at": ""})
        self.config_mgr.save()
        self._clear_extraction_checkpoint()
        self.status_var.set(f"Ricostruzione completa da {start_label}…")
        self.extract_incremental()



def choose_profile(root: tk.Tk, config: ConfigManager) -> Path | None:
    mgr = ProfileManager(config.profiles_dir)
    chooser = ProfileChooser(root, mgr)
    root.wait_window(chooser)
    return chooser.result
