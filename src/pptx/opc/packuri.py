"""Re-export of :mod:`ooxml_opc.packuri`.

The implementation that used to live here now lives in the shared
:mod:`ooxml_opc` package, consumed by python-docx, python-pptx, and
python-xlsx. Keeps the ``pptx.opc.packuri.PackURI`` import path working
for every existing caller.
"""

from __future__ import annotations

from ooxml_opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI, PackURI

__all__ = ["CONTENT_TYPES_URI", "PACKAGE_URI", "PackURI"]
