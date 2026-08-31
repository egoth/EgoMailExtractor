from __future__ import annotations

import base64
import json
import os
from pathlib import Path
from typing import Any

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

FORMAT_NAME = "EgoEncryptedStore"
FORMAT_VERSION = 1
KDF_ITERATIONS = 600_000
SALT_SIZE = 16
NONCE_SIZE = 12
KEY_SIZE = 32
ASSOCIATED_DATA = b"EgoMailExtractor.IMAP.v1"


class InvalidMasterPassword(ValueError):
    """La password master non consente di autenticare/decriptare il file."""


class EncryptedStoreError(RuntimeError):
    """File cifrato mancante, danneggiato o non compatibile."""


def _b64e(value: bytes) -> str:
    return base64.b64encode(value).decode("ascii")


def _b64d(value: str) -> bytes:
    return base64.b64decode(value.encode("ascii"), validate=True)


def _derive_key(password: str, salt: bytes, iterations: int) -> bytes:
    if not password:
        raise ValueError("La password master non può essere vuota")
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=KEY_SIZE,
        salt=salt,
        iterations=iterations,
    )
    return kdf.derive(password.encode("utf-8"))


class EncryptedJsonStore:
    """Archivio JSON cifrato autenticato con AES-256-GCM.

    Sul disco viene scritto esclusivamente il contenitore cifrato. Il salvataggio
    atomico usa un file temporaneo anch'esso cifrato: non viene mai creato un
    file intermedio contenente le credenziali in chiaro.
    """

    def __init__(self, path: str | Path):
        self.path = Path(path)

    @property
    def exists(self) -> bool:
        return self.path.exists()

    def save(self, password: str, data: dict[str, Any]) -> None:
        salt = os.urandom(SALT_SIZE)
        nonce = os.urandom(NONCE_SIZE)
        key = _derive_key(password, salt, KDF_ITERATIONS)
        plaintext = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        ciphertext = AESGCM(key).encrypt(nonce, plaintext, ASSOCIATED_DATA)
        envelope = {
            "format": FORMAT_NAME,
            "version": FORMAT_VERSION,
            "cipher": "AES-256-GCM",
            "kdf": "PBKDF2-HMAC-SHA256",
            "iterations": KDF_ITERATIONS,
            "salt": _b64e(salt),
            "nonce": _b64e(nonce),
            "ciphertext": _b64e(ciphertext),
        }
        payload = json.dumps(envelope, indent=2, ensure_ascii=False).encode("utf-8")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_name(self.path.name + ".tmp")
        try:
            with open(tmp, "wb") as f:
                f.write(payload)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp, self.path)
        finally:
            try:
                if tmp.exists():
                    tmp.unlink()
            except OSError:
                pass

    def load(self, password: str) -> dict[str, Any]:
        try:
            envelope = json.loads(self.path.read_text(encoding="utf-8"))
            if envelope.get("format") != FORMAT_NAME or int(envelope.get("version", 0)) != FORMAT_VERSION:
                raise EncryptedStoreError("Formato del file credenziali non riconosciuto")
            iterations = int(envelope["iterations"])
            salt = _b64d(envelope["salt"])
            nonce = _b64d(envelope["nonce"])
            ciphertext = _b64d(envelope["ciphertext"])
        except EncryptedStoreError:
            raise
        except Exception as exc:
            raise EncryptedStoreError("Il file delle credenziali è danneggiato o incompleto") from exc

        try:
            key = _derive_key(password, salt, iterations)
            plaintext = AESGCM(key).decrypt(nonce, ciphertext, ASSOCIATED_DATA)
        except InvalidTag as exc:
            raise InvalidMasterPassword("Password master non corretta") from exc
        except Exception as exc:
            raise EncryptedStoreError("Impossibile decriptare il file delle credenziali") from exc

        try:
            data = json.loads(plaintext.decode("utf-8"))
            if not isinstance(data, dict):
                raise ValueError("root non oggetto")
            return data
        except Exception as exc:
            raise EncryptedStoreError("Il contenuto decriptato non è valido") from exc
