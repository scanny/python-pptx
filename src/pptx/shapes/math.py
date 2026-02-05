"""Math shape objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.shapes.base import BaseShape
from pptx.math.math import Math

if TYPE_CHECKING:
    from pptx.oxml.shapes import ShapeElement
    from pptx.types import ProvidesPart


class MathShape(BaseShape):
    """Shape representing a mathematical equation on a slide."""

    def __init__(self, sp: ShapeElement, parent: ProvidesPart):
        super(MathShape, self).__init__(sp, parent)
        self._math = None

    @property
    def math(self) -> Math:
        """Math object providing access to OMML content."""
        if self._math is None:
            from pptx.math.math import Math
            # Get or create the OMML element from the shape
            # Pass the parent so Math can access the slide
            self._math = Math(self._element, self._parent)
        return self._math
