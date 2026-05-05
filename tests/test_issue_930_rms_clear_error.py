# pyright: reportPrivateUsage=false

"""Regression test for issue #930 — Azure RMS / AIP-protected ``.pptx`` error clarity.

Issue #930 (https://github.com/scanny/python-pptx/issues/930) reported that
opening an Azure Information Protection (AIP) / Rights Management Services
(RMS) / Information Rights Management (IRM)-protected ``.pptx`` failed with an
opaque :class:`~pptx.exc.PackageNotFoundError`. An RMS-protected file is a
CFBF (OLE2 compound) container — so it does *not* start with the ``PK\\x03\\x04``
zip signature — but decrypting it requires the user's Azure AD identity and
the Microsoft Information Protection SDK (C#/.NET-only); it cannot be opened
by python-pptx with a password.

This module pins the error-clarity contract:

* A CFBF container that carries RMS/IRM markers (``DRMContent`` /
  ``DRMEncryptedTransform``) is detected by :func:`pptx.opc._crypto.is_rms_protected_stream`.
* The ``_PhysPkgReader`` factory raises
  :class:`pptx.exc.RmsProtectedPackageError` with a message pointing at the
  user-guide workaround page.
* :class:`~pptx.exc.RmsProtectedPackageError` is a subclass of
  :class:`~pptx.exc.EncryptedPackageError` so callers that already
  ``except EncryptedPackageError`` continue to catch both kinds.
* Plain CFBF garbage (with no RMS markers) still goes through the existing
  Agile-Encryption path and raises :class:`~pptx.exc.EncryptedPackageError`.
* A non-CFBF, non-zip file still raises :class:`~pptx.exc.PackageNotFoundError`
  so no non-RMS callers regress.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.exc import (
    EncryptedPackageError,
    PackageNotFoundError,
    RmsProtectedPackageError,
)
from pptx.opc._crypto import (
    _OLE_SIGNATURE,
    is_encrypted_stream,
    is_rms_protected_stream,
)
from pptx.opc.serialized import _PhysPkgReader


def _rms_bytes() -> bytes:
    """Return a synthetic RMS-protected CFBF container.

    The bytes are not a complete valid OLE2 compound document — they are the
    minimum needed to exercise the python-pptx error path:

    * the 8-byte OLE2 signature, so :func:`is_encrypted_stream` matches;
    * the UTF-16LE-encoded directory-entry name ``DRMEncryptedTransform``,
      so :func:`is_rms_protected_stream` matches;
    * some filler so python-ooxml-crypto (if installed) would still fail quickly
      — but python-pptx raises :class:`RmsProtectedPackageError` *before*
      forwarding to ooxml_crypto.
    """
    transform = "DRMEncryptedTransform".encode("utf-16-le")
    content = "DRMContent".encode("utf-16-le")
    filler = b"\x00" * 512
    return _OLE_SIGNATURE + filler + transform + filler + content + filler


def _plain_cfbf_bytes() -> bytes:
    """Return a synthetic password-protected-looking CFBF container (no RMS markers)."""
    # -- UTF-16LE-encoded ``EncryptionInfo`` so it's clearly a plain Agile package
    return _OLE_SIGNATURE + b"\x00" * 256 + "EncryptionInfo".encode("utf-16-le") + b"\x00" * 256


class Describe_is_rms_protected_stream:
    """Unit-test suite for :func:`pptx.opc._crypto.is_rms_protected_stream`."""

    def it_detects_DRMEncryptedTransform_marker(self):
        stream = io.BytesIO(_rms_bytes())
        assert is_rms_protected_stream(stream) is True

    def it_detects_DRMContent_only_marker(self):
        content = "DRMContent".encode("utf-16-le")
        stream = io.BytesIO(_OLE_SIGNATURE + b"\x00" * 200 + content + b"\x00" * 200)
        assert is_rms_protected_stream(stream) is True

    def it_rejects_a_plain_CFBF_stream_without_DRM_markers(self):
        stream = io.BytesIO(_plain_cfbf_bytes())
        assert is_rms_protected_stream(stream) is False

    def it_rejects_a_plain_zip_stream(self):
        stream = io.BytesIO(b"PK\x03\x04" + b"\x00" * 100)
        assert is_rms_protected_stream(stream) is False

    def it_rejects_an_empty_stream(self):
        assert is_rms_protected_stream(io.BytesIO(b"")) is False

    def it_restores_the_stream_position(self):
        stream = io.BytesIO(_rms_bytes())
        stream.seek(42)
        is_rms_protected_stream(stream)
        assert stream.tell() == 42

    def it_does_not_match_a_utf8_ascii_DRMContent_substring(self):
        # -- ASCII "DRMContent" (not UTF-16LE) must not false-match; the CFBF
        # --- directory entries are UTF-16LE, so an ASCII hit would be a bug.
        stream = io.BytesIO(_OLE_SIGNATURE + b"DRMContent" + b"\x00" * 100)
        assert is_rms_protected_stream(stream) is False


class Describe_PhysPkgReader_factory_on_rms:
    """Unit-test suite for :meth:`_PhysPkgReader.factory` with an RMS-wrapped stream."""

    def it_raises_RmsProtectedPackageError_on_an_rms_stream_without_password(self):
        stream = io.BytesIO(_rms_bytes())

        with pytest.raises(RmsProtectedPackageError) as exc_info:
            _PhysPkgReader.factory(stream)

        # -- the message must be actionable — point at the user-guide recipe --
        msg = str(exc_info.value)
        assert "RMS" in msg or "AIP" in msg or "IRM" in msg
        assert "user guide" in msg.lower() or "workaround" in msg.lower()

    def and_it_raises_RmsProtectedPackageError_even_when_password_is_supplied(self):
        # -- Even with a password, an RMS-wrapped file cannot be decrypted
        # --- by python-pptx. The error must fire before ooxml_crypto is called. --
        stream = io.BytesIO(_rms_bytes())

        with pytest.raises(RmsProtectedPackageError):
            _PhysPkgReader.factory(stream, password="hunter2")

    def and_it_raises_RmsProtectedPackageError_on_an_rms_file_path(self, tmp_path: object):
        # -- An RMS file read from a filesystem path goes through the "open the file
        # --- and sniff its first bytes" branch; that branch must also raise. --
        p = str(tmp_path) + "/rms.pptx"
        with open(p, "wb") as f:
            f.write(_rms_bytes())

        with pytest.raises(RmsProtectedPackageError):
            _PhysPkgReader.factory(p)

    def but_a_plain_CFBF_stream_still_raises_EncryptedPackageError(self):
        # --- When the CFBF has no DRM markers and no password was supplied, we
        # --- still reach the existing password-protected-file branch. --
        stream = io.BytesIO(_plain_cfbf_bytes())

        with pytest.raises(EncryptedPackageError) as exc_info:
            _PhysPkgReader.factory(stream)

        # -- but it's NOT the RMS subclass -- the vanilla parent class --
        assert not isinstance(exc_info.value, RmsProtectedPackageError)
        assert "password-protected" in str(exc_info.value)

    def but_a_non_pptx_non_cfbf_file_still_raises_PackageNotFoundError(self, tmp_path: object):
        # -- Ordinary "not a package" errors are unchanged.
        p = str(tmp_path) + "/garbage.pptx"
        with open(p, "wb") as f:
            f.write(b"not a pptx, not cfbf, not a zip")

        with pytest.raises(PackageNotFoundError):
            _PhysPkgReader.factory(p)


class Describe_Presentation_open_on_rms:
    """Integration with the top-level :func:`pptx.Presentation` factory."""

    def it_raises_RmsProtectedPackageError_when_opening_an_rms_file(self, tmp_path: object):
        p = str(tmp_path) + "/rms-protected.pptx"
        with open(p, "wb") as f:
            f.write(_rms_bytes())

        with pytest.raises(RmsProtectedPackageError):
            Presentation(p)

    def and_it_raises_on_an_rms_file_like_stream(self):
        stream = io.BytesIO(_rms_bytes())

        with pytest.raises(RmsProtectedPackageError):
            Presentation(stream)

    def it_reports_RmsProtectedPackageError_as_subclass_of_EncryptedPackageError(self):
        # -- legacy callers expecting EncryptedPackageError keep working. --
        assert issubclass(RmsProtectedPackageError, EncryptedPackageError)

        stream = io.BytesIO(_rms_bytes())
        with pytest.raises(EncryptedPackageError):
            Presentation(stream)


class Describe_CFBF_header_sniff:
    """Ensure ``is_encrypted_stream`` still cleanly detects CFBF for our RMS bytes."""

    def it_matches_the_OLE2_header_of_the_synthetic_rms_bytes(self):
        stream = io.BytesIO(_rms_bytes())
        assert is_encrypted_stream(stream) is True
