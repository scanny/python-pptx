"""Re-export of :mod:`ooxml_opc.constants`.

The ``CONTENT_TYPE`` / ``RELATIONSHIP_TYPE`` / ``NAMESPACE`` /
``RELATIONSHIP_TARGET_MODE`` registries now live in the shared
:mod:`ooxml_opc` package. Keeps the ``pptx.opc.constants.*`` import
paths working for every existing caller.
"""

from __future__ import annotations

from ooxml_opc.constants import (
    CONTENT_TYPE,
    NAMESPACE,
    RELATIONSHIP_TARGET_MODE,
    RELATIONSHIP_TYPE,
)

__all__ = [
    "CONTENT_TYPE",
    "NAMESPACE",
    "RELATIONSHIP_TARGET_MODE",
    "RELATIONSHIP_TYPE",
]
