import json
import tempfile
import unittest
from pathlib import Path

from egomail.config import ConfigManager, ImapSettings
from egomail.crypto_store import EncryptedJsonStore, InvalidMasterPassword


class CryptoStoreTests(unittest.TestCase):
    def test_encrypt_decrypt_and_wrong_password(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "imap_credentials.enc"
            store = EncryptedJsonStore(path)
            secret = {"imap": {"host": "imap.example.org", "username": "user@example.org", "password": "SuperSecret!"}}
            store.save("master-pass", secret)
            self.assertEqual(store.load("master-pass"), secret)
            disk = path.read_text(encoding="utf-8")
            self.assertNotIn("SuperSecret!", disk)
            self.assertNotIn("user@example.org", disk)
            with self.assertRaises(InvalidMasterPassword):
                store.load("wrong")

    def test_config_keeps_imap_out_of_settings_json(self):
        with tempfile.TemporaryDirectory() as d:
            config = ConfigManager(Path(d))
            config.unlock_credentials("master", create_if_missing=True)
            config.set_imap_credentials(
                ImapSettings(host="imap.example.org", port=993, ssl=True, username="u@example.org", folder="INBOX"),
                "mail-password",
            )
            config.save()
            settings = json.loads(config.path.read_text(encoding="utf-8"))
            self.assertNotIn("imap", settings)
            encrypted = config.credentials_path.read_text(encoding="utf-8")
            self.assertNotIn("imap.example.org", encrypted)
            self.assertNotIn("mail-password", encrypted)

            config2 = ConfigManager(Path(d))
            config2.unlock_credentials("master")
            self.assertEqual(config2.imap.username, "u@example.org")
            self.assertEqual(config2.get_password(), "mail-password")

    def test_change_master_password(self):
        with tempfile.TemporaryDirectory() as d:
            config = ConfigManager(Path(d))
            config.unlock_credentials("old", create_if_missing=True)
            config.set_imap_credentials(ImapSettings(host="imap.test", username="user"), "pwd")
            config.change_master_password("new")
            config.lock_credentials()

            again = ConfigManager(Path(d))
            with self.assertRaises(InvalidMasterPassword):
                again.unlock_credentials("old")
            again.unlock_credentials("new")
            self.assertEqual(again.get_password(), "pwd")


if __name__ == "__main__":
    unittest.main()

class MultiImapProfileTests(unittest.TestCase):
    def test_multiple_profiles_are_encrypted_and_switchable(self):
        with tempfile.TemporaryDirectory() as d:
            config = ConfigManager(Path(d))
            config.unlock_credentials("master", create_if_missing=True)
            p1 = config.create_imap_profile(
                "Casa", ImapSettings(host="imap.one.test", username="one@example.org", folder="INBOX"),
                "pwd-one", make_default=True, make_active=True,
            )
            p2 = config.create_imap_profile(
                "Lavoro", ImapSettings(host="imap.two.test", username="two@example.org", folder="Archive"),
                "pwd-two", make_active=True,
            )
            self.assertEqual(config.imap_profile_count, 2)
            self.assertEqual(config.active_imap_profile_id, p2)
            self.assertEqual(config.default_imap_profile_id, p1)
            self.assertEqual(config.imap.username, "two@example.org")
            self.assertEqual(config.get_password(), "pwd-two")

            config.set_active_imap_profile(p1)
            self.assertEqual(config.active_imap_profile_name, "Casa")
            self.assertEqual(config.imap.host, "imap.one.test")
            self.assertEqual(config.get_password(), "pwd-one")

            disk = config.credentials_path.read_text(encoding="utf-8")
            for secret in ("imap.one.test", "imap.two.test", "one@example.org", "two@example.org", "pwd-one", "pwd-two"):
                self.assertNotIn(secret, disk)

            again = ConfigManager(Path(d))
            again.unlock_credentials("master")
            self.assertEqual(again.imap_profile_count, 2)
            again.set_active_imap_profile(p2)
            self.assertEqual(again.imap.folder, "Archive")
            self.assertEqual(again.get_password(), "pwd-two")

    def test_single_profile_becomes_default(self):
        with tempfile.TemporaryDirectory() as d:
            config = ConfigManager(Path(d))
            config.unlock_credentials("master", create_if_missing=True)
            pid = config.create_imap_profile(
                "Solo", ImapSettings(host="imap.test", username="solo@example.org"), "pwd"
            )
            self.assertEqual(config.default_imap_profile_id, pid)
            self.assertEqual(config.active_imap_profile_id, pid)
            self.assertTrue(config.list_imap_profiles()[0]["is_default"])
