"""lxml custom element classes for legacy PowerPoint-comment XML elements.

Implements the ECMA-376 Part 1 "legacy" comments schema. The "modern comments"
format (2018, MS extension namespace) is intentionally not implemented here; support
for that variant is deferred to a follow-up.

The legacy schema is rooted at two package parts:

- ``p:cmLst`` — one per slide, lists each ``p:cm`` comment on that slide.
- ``p:cmAuthorLst`` — package-level list of ``p:cmAuthor`` entries, one per author.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import Callable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    OxmlElement,
    RequiredAttribute,
    ZeroOrMore,
)

_OFFSET_RE = re.compile(r"([+-])(\d\d):(\d\d)")


def _parse_w3cdtf(value: str) -> dt.datetime | None:
    """Return a naive :class:`datetime.datetime` parsed from W3CDTF string `value`.

    Returns |None| if `value` is not parseable. Timezone offsets are applied to the
    naive timestamp (no tzinfo is attached) — matching the convention used by
    `coreprops`.
    """
    parseable_part = value[:19]
    offset_str = value[19:]
    timestamp: dt.datetime | None = None
    for tmpl in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%Y-%m", "%Y"):
        try:
            timestamp = dt.datetime.strptime(parseable_part, tmpl)
            break
        except ValueError:
            continue
    if timestamp is None:
        return None
    if len(offset_str) == 6:
        match = _OFFSET_RE.match(offset_str)
        if match is None:
            return timestamp
        sign, hours_str, minutes_str = match.groups()
        sign_factor = -1 if sign == "+" else 1
        delta = dt.timedelta(
            hours=int(hours_str) * sign_factor, minutes=int(minutes_str) * sign_factor
        )
        return timestamp + delta
    return timestamp


def _format_w3cdtf(value: dt.datetime) -> str:
    """Return `value` formatted as a W3CDTF string suitable for the `dt` attribute."""
    if value.tzinfo is not None:
        value = value.astimezone(dt.timezone.utc).replace(tzinfo=None)
    return value.strftime("%Y-%m-%dT%H:%M:%S")


class CT_CommentAuthor(BaseOxmlElement):
    """`p:cmAuthor` custom element class.

    Represents a single entry in the package-level comment-author registry.
    """

    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    initials: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "initials", XsdString
    )
    lastIdx: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "lastIdx", XsdUnsignedInt
    )
    clrIdx: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "clrIdx", XsdUnsignedInt
    )


class CT_CommentAuthorList(BaseOxmlElement):
    """`p:cmAuthorLst` custom element class.

    Root element of the package-level comment-authors part.
    """

    cmAuthor_lst: list[CT_CommentAuthor]

    _add_cmAuthor: Callable[..., CT_CommentAuthor]

    cmAuthor = ZeroOrMore("p:cmAuthor")

    _cmAuthorLst_tmpl = "<p:cmAuthorLst %s/>\n" % nsdecls("p")

    @classmethod
    def new(cls) -> CT_CommentAuthorList:
        """Return a newly created empty `p:cmAuthorLst` element."""
        return cast("CT_CommentAuthorList", parse_xml(cls._cmAuthorLst_tmpl))

    def add_author(self, name: str, initials: str) -> CT_CommentAuthor:
        """Append a new `p:cmAuthor` child for `name`/`initials` and return it.

        The new author is assigned the next available ``id``, ``lastIdx`` 0, and
        ``clrIdx`` equal to its id (matching PowerPoint's observed behavior).
        """
        next_id = self._next_author_id()
        return self._add_cmAuthor(
            id=next_id, name=name, initials=initials, lastIdx=0, clrIdx=next_id
        )

    def get_by_id(self, author_id: int) -> CT_CommentAuthor | None:
        """Return the `p:cmAuthor` child with id `author_id`, or |None| if not present."""
        for author in self.cmAuthor_lst:
            if author.id == author_id:
                return author
        return None

    def _next_author_id(self) -> int:
        existing = [a.id for a in self.cmAuthor_lst]
        return (max(existing) + 1) if existing else 0


class CT_Comment(BaseOxmlElement):
    """`p:cm` custom element class representing a single comment on a slide.

    Child elements are ``p:pos`` and ``p:text`` in that order per the ECMA-376 schema.
    They are accessed via direct `.find()` rather than descriptors because both are
    required and must be populated at construction time.
    """

    authorId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "authorId", XsdUnsignedInt
    )
    idx: int = RequiredAttribute("idx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    dt_attr: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dt", XsdString
    )

    @property
    def text_value(self) -> str:
        """The body text of this comment, '' if the `p:text` element is empty or missing."""
        text_el = self.find(qn("p:text"))
        if text_el is None or text_el.text is None:
            return ""
        return text_el.text

    @text_value.setter
    def text_value(self, value: str) -> None:
        text_el = self.find(qn("p:text"))
        if text_el is None:
            text_el = OxmlElement("p:text")
            self.append(text_el)
        text_el.text = value

    @property
    def datetime_value(self) -> dt.datetime | None:
        """Value of the ``dt`` attribute parsed as a naive datetime, or |None|."""
        if self.dt_attr is None:
            return None
        return _parse_w3cdtf(self.dt_attr)

    @datetime_value.setter
    def datetime_value(self, value: dt.datetime | None) -> None:
        if value is None:
            self.dt_attr = None
            return
        self.dt_attr = _format_w3cdtf(value)

    @property
    def position(self) -> tuple[int, int]:
        """(x, y) EMU position of the comment's anchor on the slide.

        Returns (0, 0) if the `p:pos` element is missing or incompletely populated.
        """
        pos_el = self.find(qn("p:pos"))
        if pos_el is None:
            return (0, 0)
        return (int(pos_el.get("x", "0")), int(pos_el.get("y", "0")))

    @position.setter
    def position(self, value: tuple[int, int]) -> None:
        x, y = value
        pos_el = self.find(qn("p:pos"))
        if pos_el is None:
            pos_el = OxmlElement("p:pos")
            # -- `p:pos` must precede `p:text` in document order --
            text_el = self.find(qn("p:text"))
            if text_el is not None:
                text_el.addprevious(pos_el)
            else:
                self.append(pos_el)
        pos_el.set("x", str(int(x)))
        pos_el.set("y", str(int(y)))


class CT_CommentList(BaseOxmlElement):
    """`p:cmLst` custom element class; root element of a slide-comments part."""

    cm_lst: list[CT_Comment]

    _add_cm: Callable[..., CT_Comment]

    cm = ZeroOrMore("p:cm")

    _cmLst_tmpl = "<p:cmLst %s/>\n" % nsdecls("p")

    @classmethod
    def new(cls) -> CT_CommentList:
        """Return a newly created, empty `p:cmLst` element."""
        return cast("CT_CommentList", parse_xml(cls._cmLst_tmpl))

    def add_comment(
        self,
        author_id: int,
        text: str,
        position: tuple[int, int] = (0, 0),
        datetime_value: dt.datetime | None = None,
    ) -> CT_Comment:
        """Append a new `p:cm` child populated from the supplied values, and return it.

        A fresh `idx` scoped to this author (1-based, ascending) is assigned.
        """
        next_idx = self._next_idx_for_author(author_id)
        comment = self._add_cm(authorId=author_id, idx=next_idx)
        # -- `p:pos` must precede `p:text` per schema --
        pos_el = OxmlElement("p:pos")
        pos_el.set("x", str(int(position[0])))
        pos_el.set("y", str(int(position[1])))
        comment.append(pos_el)
        text_el = OxmlElement("p:text")
        text_el.text = text
        comment.append(text_el)
        if datetime_value is not None:
            comment.datetime_value = datetime_value
        return comment

    def _next_idx_for_author(self, author_id: int) -> int:
        used = [c.idx for c in self.cm_lst if c.authorId == author_id]
        return (max(used) + 1) if used else 1
