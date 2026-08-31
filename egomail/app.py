from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import messagebox

from .config import ConfigManager
from .crypto_store import EncryptedStoreError, InvalidMasterPassword
from .profile import ProfileManager
from .ui import ImapProfileChooser, ImapSettingsDialog, MainWindow, MasterPasswordDialog, ProfileChooser, show_splash, apply_app_icon


def _unlock_at_startup(root: tk.Tk, config: ConfigManager) -> bool:
    """Richiede la password master dopo la visualizzazione dello splash."""
    creating = not config.credentials_initialized
    while True:
        dlg = MasterPasswordDialog(root, creating=creating)
        root.wait_window(dlg)
        if not dlg.result:
            return False
        try:
            config.unlock_credentials(dlg.result, create_if_missing=creating)
            if creating:
                migrated = config.migrate_legacy_credentials()
                if migrated:
                    messagebox.showinfo(
                        "Credenziali protette",
                        "Le impostazioni email della versione precedente sono state migrate nel file cifrato.",
                        parent=root,
                    )
            return True
        except InvalidMasterPassword:
            messagebox.showerror("Password master", "Password non corretta.", parent=root)
        except EncryptedStoreError as exc:
            messagebox.showerror("Credenziali", str(exc), parent=root)
            return False
        except Exception as exc:
            messagebox.showerror("Errore", str(exc), parent=root)
            return False


def main():
    hidden = tk.Tk()
    apply_app_icon(hidden)
    hidden.withdraw()

    show_splash(hidden)

    config = ConfigManager()
    if not _unlock_at_startup(hidden, config):
        hidden.destroy()
        return

    # Community: una sola email. Se una configurazione importata contiene
    # più account viene usato quello predefinito/attivo, senza esporre
    # nell'interfaccia la gestione multiprofilo.
    email_profiles = config.list_imap_profiles()
    if email_profiles:
        selected = config.default_imap_profile_id or config.active_imap_profile_id or str(email_profiles[0]["id"])
        config.set_active_imap_profile(selected)

    # Community: un solo profilo di estrazione. Non installiamo i profili
    # builtin multipli; al primo avvio creiamo un profilo vuoto, modificabile
    # integralmente con gli editor manuali.
    manager = ProfileManager(config.profiles_dir)
    profiles = manager.list_profiles()
    if profiles:
        profile_path = profiles[0]
    else:
        profile_path = manager.create("Profilo Community")

    hidden.destroy()
    app = MainWindow(config, profile_path)
    apply_app_icon(app)
    if not config.imap.host or not config.imap.username or not config.get_password():
        dlg = ImapSettingsDialog(app, config)
        app.wait_window(dlg)
    app.mainloop()
    config.lock_credentials()
