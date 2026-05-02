"""FontPart — binary TrueType font-data part embedded in a presentation.

Embedded fonts are stored as one `.fntdata` part per style per typeface at partnames
like `/ppt/fonts/font1.fntdata`. The `p:presentation` part references them via
`p:embeddedFontLst/p:embeddedFont/p:regular|p:bold|p:italic|p:boldItalic` elements
that each carry an `r:id` relationship reference to a `FontPart`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part

if TYPE_CHECKING:
    from pptx.package import Package


class FontPart(Part):
    """A binary embedded-font part.

    Carries the bytes of a TrueType / OpenType (.ttf / .otf) font file in a partname of the
    form `/ppt/fonts/font{n}.fntdata`, content-type `application/x-fontdata`.
    """

    partname_template = "/ppt/fonts/font%d.fntdata"
    content_type = CT.X_FONTDATA

    @classmethod
    def new(cls, blob: bytes, package: Package) -> FontPart:
        """Return a new |FontPart| containing `blob` added to `package`.

        The new part receives the next available partname of form
        `/ppt/fonts/font{n}.fntdata`.
        """
        return cls(
            package.next_partname(cls.partname_template),
            cls.content_type,
            package,
            blob,
        )
