"""lxml custom element classes for the `p:tagLst` tags-part XML.

Implements the ECMA-376 `p:tagLst` root element and its `p:tag` children.
A slide (or slide-layout / slide-master) references a tags part via
``p:cSld/p:custDataLst/p:tags`` whose ``r:id`` points at the target
|TagsPart|. The part body itself is a ``p:tagLst`` with zero-or-more
``p:tag`` entries, each of which has a required ``name`` and ``val`` pair
— the VBA-style custom-tag mechanism PowerPoint exposes via
``SlideRange.Tags`` in its COM object model (see issue #578).
"""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_StringTag(BaseOxmlElement):
    """`p:tag` custom element class.

    A single name/value entry in a ``p:tagLst``. Both attributes are
    required per ``CT_StringTag`` in ``pml.xsd``.
    """

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TagList(BaseOxmlElement):
    """`p:tagLst` custom element class — root element of a tags part."""

    tag_lst: list[CT_StringTag]

    _add_tag: Callable[..., CT_StringTag]

    # -- Assigning the descriptor as ``tag`` mirrors the child element's local
    # -- name, matching the convention used by other ``ZeroOrMore`` members in
    # -- this codebase (see e.g. ``CT_CommentList.cm``). The generated
    # -- ``tag_lst`` read-property masks the ``ZeroOrMore`` object at
    # -- runtime. A pyright ignore is needed because the descriptor
    # -- notionally shadows the ``tag`` property inherited from
    # -- ``lxml._Element`` (which reports this element's own qualified name);
    # -- no consumer of ``CT_TagList`` relies on that shadow — the element
    # -- tag is still exposed through the inherited attribute at runtime.
    tag = ZeroOrMore(  # pyright: ignore[reportAssignmentType,reportIncompatibleMethodOverride]
        "p:tag"
    )

    _tagLst_tmpl = "<p:tagLst %s/>\n" % nsdecls("p")

    @classmethod
    def new(cls) -> CT_TagList:
        """Return a newly created, empty `p:tagLst` element."""
        return cast("CT_TagList", parse_xml(cls._tagLst_tmpl))

    def get_by_name(self, name: str) -> CT_StringTag | None:
        """Return the first `p:tag` child with ``@name == name``, or |None|."""
        for tag in self.tag_lst:
            if tag.name == name:
                return tag
        return None

    def set_tag(self, name: str, value: str) -> CT_StringTag:
        """Set the ``val`` of the `p:tag` named `name`, creating it when absent.

        Returns the affected `p:tag` element.
        """
        existing = self.get_by_name(name)
        if existing is not None:
            existing.val = value
            return existing
        return self._add_tag(name=name, val=value)

    def remove_tag(self, name: str) -> bool:
        """Remove the first `p:tag` having ``@name == name``.

        Returns ``True`` when a matching tag was found and removed, ``False``
        otherwise.
        """
        existing = self.get_by_name(name)
        if existing is None:
            return False
        self.remove(existing)
        return True


class CT_TagsData(BaseOxmlElement):
    """`p:tags` custom element class — child of `p:custDataLst`.

    Carries a required ``r:id`` referencing the related |TagsPart|. This is
    the *reference* element that lives inside a slide / layout / master XML;
    it is not the tags part itself (that is ``p:tagLst``).
    """

    rId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "r:id", XsdString
    )


class CT_CustomerDataList(BaseOxmlElement):
    """`p:custDataLst` custom element class.

    Holds an optional ``p:tags`` reference (and any ``p:custData`` children
    we do not otherwise model). Only the ``p:tags`` child is modeled here —
    that is sufficient for round-tripping custom-tag metadata on slides.
    Per the ECMA-376 `CT_CustomerDataList` schema, ``p:tags`` always
    follows any ``p:custData`` children, so it has no successors.
    """

    get_or_add_tags: Callable[[], CT_TagsData]
    _remove_tags: Callable[[], None]

    tags: CT_TagsData | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:tags", successors=()
    )
