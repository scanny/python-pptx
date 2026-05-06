"""Re-export of :mod:`ooxml_opc.spec`.

The ``default_content_types`` table now lives in the shared
:mod:`ooxml_opc` package.
"""

from __future__ import annotations

from ooxml_opc.spec import default_content_types, image_content_types

__all__ = ["default_content_types", "image_content_types"]
