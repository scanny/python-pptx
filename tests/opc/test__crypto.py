# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.opc._crypto` module."""

from __future__ import annotations

import importlib.util
import io

import pytest

from pptx.exc import EncryptedPackageError
from pptx.opc import _crypto
from pptx.opc._crypto import (
    _OLE_SIGNATURE,
    decrypt_stream,
    encrypt_bytes,
    is_encrypted_stream,
)

# -- Issue #327: gracefully skip tests that depend on the optional
#    python-ooxml-crypto package when it is not installed. --
requires_ooxml_crypto = pytest.mark.skipif(
    importlib.util.find_spec("ooxml_crypto") is None,
    reason="python-ooxml-crypto is not installed (optional test dependency for issue #327)",
)


class Describe_is_encrypted_stream:
    """Unit-test suite for `pptx.opc._crypto.is_encrypted_stream`."""

    def it_returns_True_when_stream_starts_with_the_OLE_magic(self):
        stream = io.BytesIO(_OLE_SIGNATURE + b"rest-of-cfbf-container")
        assert is_encrypted_stream(stream) is True

    def it_returns_False_for_a_plain_zip_stream(self):
        stream = io.BytesIO(b"PK\x03\x04rest-of-zip")
        assert is_encrypted_stream(stream) is False

    def it_returns_False_for_an_empty_stream(self):
        assert is_encrypted_stream(io.BytesIO(b"")) is False

    def it_restores_the_stream_position(self):
        stream = io.BytesIO(_OLE_SIGNATURE + b"rest")
        stream.seek(3)
        is_encrypted_stream(stream)
        assert stream.tell() == 3


class Describe_decrypt_stream:
    """Unit-test suite for `pptx.opc._crypto.decrypt_stream`."""

    def it_raises_when_ooxml_crypto_is_not_installed(self, monkeypatch: pytest.MonkeyPatch):
        # -- block the ooxml_crypto import so the ImportError branch executes --
        import builtins

        real_import = builtins.__import__

        def fake_import(name: str, *args: object, **kwargs: object):
            if name.startswith("ooxml_crypto"):
                raise ImportError(name)
            return real_import(name, *args, **kwargs)  # pyright: ignore[reportArgumentType]

        monkeypatch.setattr(builtins, "__import__", fake_import)

        with pytest.raises(EncryptedPackageError, match="python-ooxml-crypto"):
            decrypt_stream(io.BytesIO(b""), "pw")

    @requires_ooxml_crypto
    def it_raises_on_wrong_password(self, encrypted_minimal_pptx: bytes):
        with pytest.raises(EncryptedPackageError, match="password does not match"):
            decrypt_stream(io.BytesIO(encrypted_minimal_pptx), "wrong")

    def it_raises_on_malformed_encrypted_input(self):
        # -- bytes that pass the OLE sniff but fail further down the ooxml_crypto path --
        garbage = _OLE_SIGNATURE + b"\x00" * 4096
        with pytest.raises(EncryptedPackageError):
            decrypt_stream(io.BytesIO(garbage), "pw")

    @requires_ooxml_crypto
    def it_returns_plain_bytes_on_success(
        self, encrypted_minimal_pptx: bytes, minimal_pptx_bytes: bytes
    ):
        plain = decrypt_stream(io.BytesIO(encrypted_minimal_pptx), "unittest")
        assert plain == minimal_pptx_bytes

    # -- fixtures ------------------------------------------------------

    @pytest.fixture
    def minimal_pptx_bytes(self) -> bytes:
        import os

        here = os.path.dirname(os.path.dirname(__file__))
        with open(os.path.join(here, "test_files", "minimal.pptx"), "rb") as f:
            return f.read()

    @pytest.fixture
    def encrypted_minimal_pptx(self, minimal_pptx_bytes: bytes) -> bytes:
        return encrypt_bytes(minimal_pptx_bytes, "unittest")


class Describe_encrypt_bytes:
    """Unit-test suite for `pptx.opc._crypto.encrypt_bytes`."""

    def it_raises_when_ooxml_crypto_is_not_installed(self, monkeypatch: pytest.MonkeyPatch):
        import builtins

        real_import = builtins.__import__

        def fake_import(name: str, *args: object, **kwargs: object):
            if name.startswith("ooxml_crypto"):
                raise ImportError(name)
            return real_import(name, *args, **kwargs)  # pyright: ignore[reportArgumentType]

        monkeypatch.setattr(builtins, "__import__", fake_import)

        with pytest.raises(EncryptedPackageError, match="python-ooxml-crypto"):
            encrypt_bytes(b"", "pw")

    @requires_ooxml_crypto
    def it_produces_a_CFBF_container(self):
        import os

        here = os.path.dirname(os.path.dirname(__file__))
        with open(os.path.join(here, "test_files", "minimal.pptx"), "rb") as f:
            plain = f.read()

        encrypted = encrypt_bytes(plain, "pw")

        assert encrypted.startswith(_OLE_SIGNATURE)

    @requires_ooxml_crypto
    def it_produces_bytes_that_round_trip_through_decrypt(self):
        import os

        here = os.path.dirname(os.path.dirname(__file__))
        with open(os.path.join(here, "test_files", "minimal.pptx"), "rb") as f:
            plain = f.read()

        encrypted = encrypt_bytes(plain, "pw")
        recovered = decrypt_stream(io.BytesIO(encrypted), "pw")

        assert recovered == plain


class Describe_missing_dep_message:
    """The message exposed when python-ooxml-crypto is absent."""

    def it_mentions_the_package_and_pip_install(self):
        assert "python-ooxml-crypto" in _crypto._MISSING_DEP_MSG
        assert "pip install" in _crypto._MISSING_DEP_MSG
