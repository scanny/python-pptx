"""Part implementing the VBA-style custom-tag feature.

A single |TagsPart| holds one ``p:tagLst`` XML document and is related to
a slide (or slide-layout / slide-master / notes-slide) via the
``TAGS`` relationship type. Each ``p:tag`` child carries a ``name``
/ ``val`` string pair, mirroring the VBA ``Slide.Tags`` API (issue #578).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.oxml.tags import CT_TagList

if TYPE_CHECKING:
    from pptx.opc.packuri import PackURI
    from pptx.package import Package


class TagsPart(XmlPart):
    """Package part holding a single ``p:tagLst`` document.

    Partnames follow the convention ``/ppt/tags/tag[N].xml`` — the same
    convention PowerPoint uses when it writes a tags part.
    """

    _element: CT_TagList

    @classmethod
    def new(cls, package: Package, partname: PackURI) -> TagsPart:
        """Return a newly created, empty |TagsPart| for `package` at `partname`."""
        return cls(partname, CT.PML_TAGS, package, CT_TagList.new())

    @property
    def tag_list(self) -> CT_TagList:
        """The underlying ``p:tagLst`` oxml element."""
        return self._element
