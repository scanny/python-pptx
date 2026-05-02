"""Optional password-protection (ECMA-376 Agile Encryption) support.

Reading and writing password-protected `.pptx` files is delegated to the optional
``msoffcrypto-tool`` third-party package so that python-pptx does not need to carry
its own implementation of AES key derivation and CFBF compound-document parsing.

This module is a thin adapter:

* :func:`is_encrypted_stream` sniffs the OLE2 compound-document magic signature so a
  caller can detect an encrypted package without loading ``msoffcrypto`` at all.
* :func:`decrypt_stream` decrypts an encrypted OOXML stream to bytes.
* :func:`encrypt_bytes` encrypts plain OOXML bytes to an encrypted bytestring.

Each function raises :class:`pptx.exc.EncryptedPackageError` with an actionable message
when ``msoffcrypto-tool`` is not installed or the password is wrong.
"""

# pyright: reportMissingTypeStubs=false, reportUnknownMemberType=false
# pyright: reportUnknownVariableType=false, reportUnknownArgumentType=false
# pyright: reportCallIssue=false

from __future__ import annotations

import io
from typing import IO

from pptx.exc import EncryptedPackageError

# -- OLE2 Compound File Binary Format signature; every ECMA-376 encrypted OOXML file
# -- (Agile or Standard) begins with this magic because it is stored as a CFBF container
# -- with an ``EncryptionInfo`` and ``EncryptedPackage`` stream.
_OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

_MISSING_DEP_MSG = (
    "password-protected .pptx files require the optional 'msoffcrypto-tool' "
    "package. Install it with `pip install msoffcrypto-tool`."
)


def is_encrypted_stream(stream: IO[bytes]) -> bool:
    """Return True if the first 8 bytes of `stream` are the OLE2 magic signature.

    The stream's position is restored before return. A `.pptx` that is stored as a plain
    ZIP archive starts with ``PK\\x03\\x04``; an encrypted one is wrapped in a CFBF
    container and starts with the OLE2 magic.
    """
    pos = stream.tell()
    try:
        header = stream.read(len(_OLE_SIGNATURE))
    finally:
        stream.seek(pos)
    return header == _OLE_SIGNATURE


def decrypt_stream(stream: IO[bytes], password: str) -> bytes:
    """Return the plaintext OOXML bytes from encrypted `stream`.

    Raises :class:`pptx.exc.EncryptedPackageError` if ``msoffcrypto-tool`` is not
    installed, if the file is not a supported encrypted OOXML file, or if `password`
    is wrong.
    """
    try:
        import msoffcrypto
        from msoffcrypto.exceptions import DecryptionError, FileFormatError, InvalidKeyError
    except ImportError as exc:
        raise EncryptedPackageError(_MISSING_DEP_MSG) from exc

    try:
        office_file = msoffcrypto.OfficeFile(stream)
        office_file.load_key(password=password, verify_password=True)
        out = io.BytesIO()
        office_file.decrypt(out)
    except InvalidKeyError as exc:
        raise EncryptedPackageError(
            "password does not match the password used to encrypt this .pptx file"
        ) from exc
    except (DecryptionError, FileFormatError, ValueError) as exc:
        raise EncryptedPackageError(f"unable to decrypt .pptx file: {exc}") from exc

    return out.getvalue()


def encrypt_bytes(plain_bytes: bytes, password: str) -> bytes:
    """Return encrypted OOXML bytes for the given plain `plain_bytes`.

    Uses ECMA-376 Agile Encryption (the format PowerPoint writes when a user sets a
    password in the desktop app).

    Raises :class:`pptx.exc.EncryptedPackageError` if ``msoffcrypto-tool`` is not
    installed or encryption fails.
    """
    try:
        from msoffcrypto.exceptions import EncryptionError
        from msoffcrypto.format.ooxml import OOXMLFile
    except ImportError as exc:
        raise EncryptedPackageError(_MISSING_DEP_MSG) from exc

    try:
        office_file = OOXMLFile(io.BytesIO(plain_bytes))
        out = io.BytesIO()
        office_file.encrypt(password, out)
    except EncryptionError as exc:
        raise EncryptedPackageError(f"unable to encrypt .pptx file: {exc}") from exc

    return out.getvalue()
