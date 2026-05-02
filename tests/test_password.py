"""Integration tests for password-protected .pptx read/write.

These exercise the real msoffcrypto-tool integration end-to-end without mocks, so they
complement the more granular tests in :mod:`tests.opc.test__crypto` and
:mod:`tests.opc.test_serialized`.
"""

from __future__ import annotations

import io
import os

import pytest

from pptx import Presentation
from pptx.exc import EncryptedPackageError

from .unitutil.file import absjoin, test_file_dir

_MINIMAL_PPTX = absjoin(test_file_dir, "minimal.pptx")


class DescribePasswordRoundTrip:
    """Save-with-password then open-with-password returns the same content."""

    def it_round_trips_a_password_protected_pptx_through_a_stream(self):
        prs = Presentation(_MINIMAL_PPTX)

        encrypted = io.BytesIO()
        prs.save(encrypted, password="s3cret")

        encrypted.seek(0)
        prs2 = Presentation(encrypted, password="s3cret")

        # -- the encrypted bytes start with the OLE2 compound-document magic --
        encrypted.seek(0)
        assert encrypted.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        # -- and the re-opened presentation behaves like the original --
        assert len(prs2.slides) == len(prs.slides)

    def it_round_trips_a_password_protected_pptx_through_a_file_path(self, tmp_path: object):
        prs = Presentation(_MINIMAL_PPTX)
        out_path = os.path.join(str(tmp_path), "encrypted.pptx")

        prs.save(out_path, password="s3cret")
        prs2 = Presentation(out_path, password="s3cret")

        assert os.path.getsize(out_path) > 0
        assert len(prs2.slides) == len(prs.slides)


class DescribeOpenPasswordProtectedPptx:
    """Error handling when opening an encrypted file."""

    def it_raises_when_no_password_is_supplied(self, encrypted_pptx_stream: io.BytesIO):
        with pytest.raises(EncryptedPackageError, match="password-protected"):
            Presentation(encrypted_pptx_stream)

    def it_raises_on_a_wrong_password(self, encrypted_pptx_stream: io.BytesIO):
        with pytest.raises(EncryptedPackageError, match="does not match"):
            Presentation(encrypted_pptx_stream, password="WRONG")

    # -- fixtures ------------------------------------------------------

    @pytest.fixture
    def encrypted_pptx_stream(self) -> io.BytesIO:
        prs = Presentation(_MINIMAL_PPTX)
        out = io.BytesIO()
        prs.save(out, password="right")
        out.seek(0)
        return out
