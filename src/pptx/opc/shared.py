"""Re-export of :mod:`ooxml_opc.shared`.

:class:`CaseInsensitiveDict` now lives in the shared :mod:`ooxml_opc`
package. Keeps the ``pptx.opc.shared.CaseInsensitiveDict`` import path
working for every existing caller.
"""

from __future__ import annotations

from ooxml_opc.shared import CaseInsensitiveDict

__all__ = ["CaseInsensitiveDict"]
