"""Parts implementing the legacy PowerPoint comment feature.

Two distinct parts are modeled:

- |CommentAuthorsPart| — single package-level part (``/ppt/commentAuthors.xml``)
  holding the list of comment-author entries referenced by all slide comments.
- |CommentsPart| — per-slide part (``/ppt/comments/comment[N].xml``) holding the
  list of comments anchored on a single slide.

The "modern comments" format introduced by Office 2018 uses separate MS-extension
namespaces and parts; it is intentionally not modeled here.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.comments import CT_CommentAuthorList, CT_CommentList

if TYPE_CHECKING:
    from pptx.package import Package


class CommentAuthorsPart(XmlPart):
    """Package part for the `p:cmAuthorLst` (``/ppt/commentAuthors.xml``).

    A single instance is shared by every slide that has comments; the part is
    referenced by a package-level relationship of reltype ``COMMENT_AUTHORS``.
    """

    _element: CT_CommentAuthorList

    @classmethod
    def new(cls, package: Package) -> CommentAuthorsPart:
        """Return a newly created, empty |CommentAuthorsPart| for `package`."""
        return cls(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            package,
            CT_CommentAuthorList.new(),
        )

    @property
    def author_list(self) -> CT_CommentAuthorList:
        """The underlying `p:cmAuthorLst` oxml element."""
        return self._element


class CommentsPart(XmlPart):
    """Package part for a slide-level `p:cmLst` (``/ppt/comments/commentN.xml``)."""

    _element: CT_CommentList

    @classmethod
    def new(cls, package: Package, partname: PackURI) -> CommentsPart:
        """Return a newly created, empty |CommentsPart| for `package` at `partname`."""
        return cls(partname, CT.PML_COMMENTS, package, CT_CommentList.new())

    @property
    def comment_list(self) -> CT_CommentList:
        """The underlying `p:cmLst` oxml element."""
        return self._element
