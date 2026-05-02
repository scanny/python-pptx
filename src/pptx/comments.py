"""High-level API for the legacy PowerPoint review-comment feature.

This module exposes user-facing proxies — |Comment|, |Comments|, |CommentAuthor|, and
|CommentAuthors| — layered on top of the :mod:`pptx.oxml.comments` element classes
and the :mod:`pptx.parts.comments` parts.

Only the ECMA-376 Part 1 "legacy" comment schema is supported. The 2018 "modern
comments" format (MS extension namespace) is not represented by these classes.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from pptx.oxml.comments import (
        CT_Comment,
        CT_CommentAuthor,
        CT_CommentAuthorList,
        CT_CommentList,
    )
    from pptx.parts.comments import CommentAuthorsPart, CommentsPart
    from pptx.slide import Slide


class CommentAuthor:
    """Proxy for a single `p:cmAuthor` element in the comment-authors part.

    Exposes read-only properties for the stable author metadata and the integer id
    used to link comments back to this author.
    """

    def __init__(self, cmAuthor: CT_CommentAuthor):
        self._element = cmAuthor

    @property
    def id(self) -> int:
        """The stable integer id that slide-level comments reference."""
        return self._element.id

    @property
    def name(self) -> str:
        """The author's display name."""
        return self._element.name

    @property
    def initials(self) -> str:
        """The author's initials, shown in the PowerPoint comments pane."""
        return self._element.initials


class CommentAuthors:
    """Package-level collection of comment authors.

    Supports ``len()``, iteration, and indexed / id-based lookup. A new author is
    added with :meth:`add_author`. There is at most one author-list part per package;
    if one does not yet exist when it is first accessed through the high-level API,
    one is created lazily.
    """

    def __init__(self, part: CommentAuthorsPart):
        self._part = part
        self._element: CT_CommentAuthorList = part.author_list

    def __iter__(self) -> Iterator[CommentAuthor]:
        for cmAuthor in self._element.cmAuthor_lst:
            yield CommentAuthor(cmAuthor)

    def __len__(self) -> int:
        return len(self._element.cmAuthor_lst)

    def __getitem__(self, idx: int) -> CommentAuthor:
        return CommentAuthor(self._element.cmAuthor_lst[idx])

    def add_author(self, name: str, initials: str) -> CommentAuthor:
        """Append a new author with `name`/`initials` and return its proxy."""
        return CommentAuthor(self._element.add_author(name, initials))

    def get_by_id(self, author_id: int) -> CommentAuthor | None:
        """Return the author having `author_id`, or |None| if not present."""
        cmAuthor = self._element.get_by_id(author_id)
        if cmAuthor is None:
            return None
        return CommentAuthor(cmAuthor)

    def get_or_add(self, name: str, initials: str) -> CommentAuthor:
        """Return an existing author matching `name` + `initials`, else add one."""
        for author in self:
            if author.name == name and author.initials == initials:
                return author
        return self.add_author(name, initials)


class Comment:
    """Proxy for a single `p:cm` element — one comment on one slide."""

    def __init__(self, cm: CT_Comment, authors: CommentAuthors):
        self._element = cm
        self._authors = authors

    @property
    def text(self) -> str:
        """The body text of this comment."""
        return self._element.text_value

    @text.setter
    def text(self, value: str) -> None:
        self._element.text_value = value

    @property
    def author_id(self) -> int:
        """Integer id linking this comment to its `CommentAuthor`."""
        return self._element.authorId

    @property
    def author(self) -> CommentAuthor | None:
        """The |CommentAuthor| referenced by this comment, or |None| if not present."""
        return self._authors.get_by_id(self._element.authorId)

    @property
    def idx(self) -> int:
        """Per-author sequence index of this comment, starting at 1."""
        return self._element.idx

    @property
    def datetime(self) -> dt.datetime | None:
        """Creation/modification timestamp of the comment, or |None| if unset."""
        return self._element.datetime_value

    @datetime.setter
    def datetime(self, value: dt.datetime | None) -> None:
        self._element.datetime_value = value

    @property
    def position(self) -> tuple[int, int]:
        """Anchor position of the comment on the slide, as (x, y) in EMU."""
        return self._element.position

    @position.setter
    def position(self, value: tuple[int, int]) -> None:
        self._element.position = value


class Comments:
    """Collection of comments on a single slide.

    Supports ``len()``, iteration, and indexed access. A new comment is added via
    :meth:`add_comment`, which requires a `CommentAuthor` so the comment's
    `authorId` reference is well-formed.
    """

    def __init__(self, slide: Slide, part: CommentsPart, authors: CommentAuthors):
        self._slide = slide
        self._part = part
        self._element: CT_CommentList = part.comment_list
        self._authors = authors

    def __iter__(self) -> Iterator[Comment]:
        for cm in self._element.cm_lst:
            yield Comment(cm, self._authors)

    def __len__(self) -> int:
        return len(self._element.cm_lst)

    def __getitem__(self, idx: int) -> Comment:
        return Comment(self._element.cm_lst[idx], self._authors)

    def add_comment(
        self,
        author: CommentAuthor,
        text: str,
        position: tuple[int, int] = (0, 0),
        datetime_value: dt.datetime | None = None,
    ) -> Comment:
        """Append a new comment authored by `author` and return its proxy.

        The new comment is tagged with `author`'s id; `text`, `position` (EMU), and
        the optional `datetime_value` are written directly.
        """
        cm = self._element.add_comment(
            author_id=author.id,
            text=text,
            position=position,
            datetime_value=datetime_value,
        )
        return Comment(cm, self._authors)
