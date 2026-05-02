# pyright: reportPrivateUsage=false

"""Regression test for issue #756 — legacy PowerPoint comments.

Issue #756 (https://github.com/scanny/python-pptx/issues/756) asked for
read/write support of the PowerPoint "review comments" feature — the
yellow sticky-note annotations authors attach to individual slides via
the *Insert > Comment* command. The feature is implemented by the
legacy ECMA-376 Part 1 comment schema shipped under #487:

* ``Slide.comments`` — :class:`pptx.comments.Comments` collection of
  :class:`pptx.comments.Comment` objects anchored on this slide.
* ``Slide.has_comments`` — side-effect-free probe for whether a slide
  already has a comments part.
* Package-level author registry (``p:cmAuthorLst``) accessed through
  :attr:`pptx.comments.Comments._authors`, with
  ``CommentAuthors.add_author(name, initials)`` and ``get_or_add(...)``
  helpers so the comment's ``authorId`` back-reference is well-formed.
* Legacy part topology per ECMA-376 §13.3.3 — the comment-authors part
  hangs off the presentation part (package-scoped); each slide with
  comments carries its own comments part.

This verify suite pins the eight capability scenarios from the #756
reporter's perspective so the issue can be closed:

1. Adding a new comment stamps the body text, author, anchor position,
   and timestamp that PowerPoint would need to render the comment.
2. The same comment reads back through the :class:`Comments` iterator
   with every attribute intact.
3. Multiple comments on one slide preserve insertion order.
4. The package-level author registry deduplicates — two calls to
   ``get_or_add(name, initials)`` with matching values reuse the
   existing :class:`CommentAuthor`.
5. Deleting a comment (via the underlying ``p:cmLst`` remove helper)
   drops it from the slide's comments collection.
6. ``Comment.position`` round-trips (x, y) in English Metric Units.
7. A comment survives ``Presentation.save`` + reopen with every
   attribute preserved.
8. The public API surface ``presentation.slides[0].comments`` is
   reachable off a freshly authored :class:`Presentation`.

Any regression that breaks the comment part topology, the author-id
back-reference, the W3CDTF timestamp, or the ``p:pos`` (x, y) EMU
anchor will reproduce the #756 symptom and be caught here.
"""

from __future__ import annotations

import datetime as dt
import io
from typing import TYPE_CHECKING, Iterator

import pytest

from pptx import Presentation
from pptx.comments import Comment, CommentAuthor, Comments

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationClass


@pytest.fixture
def _restore_part_factory() -> Iterator[None]:
    """Guard against ``PartFactory.part_type_for`` pollution from other tests.

    ``Presentation(stream)`` reopen dispatches through
    ``PartFactory.part_type_for``, which earlier test modules overwrite
    with mocks and do not restore. This fixture snapshots the registry
    on entry and restores it on teardown so that the round-trip
    scenarios below can reopen the presentation cleanly regardless of
    collection order.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.comments import CommentAuthorsPart, CommentsPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_COMMENTS: CommentsPart,
        CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


@pytest.fixture
def prs_with_slide() -> PresentationClass:
    """A fresh one-slide :class:`Presentation` for the write-path scenarios."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[5])
    return prs


class DescribeIssue756Comments:
    """#756 legacy PowerPoint-comments verify-and-close.

    Pins the eight reporter-facing scenarios from the #756 ask so the
    issue can be closed as resolved by the #487 legacy-comments work.
    """

    # -- 1. Add a new comment (text, author, position, timestamp) -------

    def it_adds_a_new_comment_with_text_author_position_and_timestamp(
        self, prs_with_slide: PresentationClass
    ) -> None:
        """Adding a comment stamps every attribute the #756 reporter needs.

        The :class:`CommentAuthor` reference threaded through
        ``add_comment`` is the canonical way to link a new comment back
        to the package-level author registry so PowerPoint can resolve
        the initials-colored dot shown in its comments pane. The
        anchor ``position`` is recorded in English Metric Units, and the
        ``datetime_value`` is written as W3CDTF per ECMA-376 §19.4.1
        (``dt`` attribute on ``p:cm``).
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        alice = comments._authors.add_author("Alice", "A.")
        when = dt.datetime(2025, 3, 14, 15, 9, 26)

        comment = comments.add_comment(
            alice,
            "Can we tighten this wording?",
            position=(914400, 457200),
            datetime_value=when,
        )

        # -- all four reporter-facing attributes populated as given --
        assert isinstance(comment, Comment)
        assert comment.text == "Can we tighten this wording?"
        assert comment.author is not None
        assert comment.author.name == "Alice"
        assert comment.author.initials == "A."
        assert comment.author_id == alice.id
        assert comment.position == (914400, 457200)
        assert comment.datetime == when

    # -- 2. Read comments back ------------------------------------------

    def it_reads_comments_back_through_the_Comments_iterator(
        self, prs_with_slide: PresentationClass
    ) -> None:
        """The :class:`Comments` collection iterates over every added comment.

        Iteration yields :class:`Comment` proxies with every attribute
        readable — this is the read path the #756 reporter wanted for
        "list all review comments on this slide".
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        alice = comments._authors.add_author("Alice", "A.")
        when = dt.datetime(2025, 6, 1, 12, 0, 0)
        comments.add_comment(alice, "first pass", datetime_value=when)

        read_back = list(comments)

        assert len(read_back) == 1
        assert read_back[0].text == "first pass"
        assert read_back[0].author is not None
        assert read_back[0].author.name == "Alice"
        assert read_back[0].datetime == when

    # -- 3. Multiple comments per slide — ordering ----------------------

    def it_preserves_insertion_order_for_multiple_comments(
        self, prs_with_slide: PresentationClass
    ) -> None:
        """Comments iterate in document order, matching insertion order.

        PowerPoint renders comments in the order they appear in the
        ``p:cmLst``; :meth:`Comments.add_comment` appends the new
        ``p:cm`` child, so insertion order and iteration order coincide.
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        alice = comments._authors.add_author("Alice", "A.")

        comments.add_comment(alice, "one")
        comments.add_comment(alice, "two")
        comments.add_comment(alice, "three")

        assert [c.text for c in comments] == ["one", "two", "three"]
        # -- and indexed access is consistent with iteration --
        assert comments[0].text == "one"
        assert comments[2].text == "three"
        assert len(comments) == 3

    # -- 4. Multiple authors — registry deduplicates --------------------

    def it_deduplicates_authors_across_get_or_add_calls(
        self, prs_with_slide: PresentationClass
    ) -> None:
        """Matching ``name``+``initials`` reuses an existing author.

        The package-level author registry (``p:cmAuthorLst``) carries a
        stable integer id that slide-level comments reference via
        ``p:cm/@authorId``. :meth:`CommentAuthors.get_or_add` is the
        safe way to introduce an author without materialising two
        registry entries when the same person adds multiple comments.
        Distinct author profiles (different initials, same name) remain
        separate.
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        authors = comments._authors

        alice_first = authors.get_or_add("Alice", "A.")
        bob = authors.get_or_add("Bob", "B.")
        alice_again = authors.get_or_add("Alice", "A.")
        # -- a different profile (same name, different initials) is a
        # -- distinct entry rather than a dedup hit --
        alice_admin = authors.get_or_add("Alice", "AA")

        assert len(authors) == 3
        assert alice_first.id == alice_again.id
        assert alice_first.id != bob.id
        assert alice_first.id != alice_admin.id
        assert {a.name for a in authors} == {"Alice", "Bob"}
        assert {(a.name, a.initials) for a in authors} == {
            ("Alice", "A."),
            ("Bob", "B."),
            ("Alice", "AA"),
        }

    # -- 5. Delete a comment --------------------------------------------

    def it_removes_a_comment_from_the_slide_collection(
        self, prs_with_slide: PresentationClass
    ) -> None:
        """Removing a ``p:cm`` drops it from the :class:`Comments` iterator.

        The high-level API does not (yet) expose an opinionated delete
        helper — the #756 ask is for read/write support — but the
        underlying ``p:cmLst`` is a standard lxml element tree and the
        user can drop a comment via the generated ``remove_cm``
        descriptor helper. After removal, iteration and indexed access
        reflect the shortened collection.
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        alice = comments._authors.add_author("Alice", "A.")

        comments.add_comment(alice, "keep me")
        second = comments.add_comment(alice, "delete me")
        comments.add_comment(alice, "keep me too")

        assert len(comments) == 3

        # -- remove the middle p:cm element through the generated
        # -- ZeroOrMore remover on the underlying cmLst --
        cmLst = comments._element
        cmLst.remove(second._element)

        # -- collection reflects the deletion --
        assert len(comments) == 2
        assert [c.text for c in comments] == ["keep me", "keep me too"]

    # -- 6. Comment position (EMU) --------------------------------------

    def it_round_trips_the_position_of_a_comment(self, prs_with_slide: PresentationClass) -> None:
        """``Comment.position`` round-trips an (x, y) pair in EMU.

        The anchor position of a comment on the slide surface is the
        ``p:pos/@x`` / ``p:pos/@y`` pair (EMU — English Metric Units).
        Reads and writes go through the :attr:`Comment.position`
        property.
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        alice = comments._authors.add_author("Alice", "A.")

        # -- default position when not provided is (0, 0) --
        defaulted = comments.add_comment(alice, "no position")
        assert defaulted.position == (0, 0)

        # -- explicit write through the setter --
        placed = comments.add_comment(alice, "placed", position=(1234567, 8901234))
        assert placed.position == (1234567, 8901234)

        # -- mutating the setter updates the underlying p:pos --
        placed.position = (7654321, 1234567)
        assert placed.position == (7654321, 1234567)

    # -- 7. Round-trip through save / reopen ----------------------------

    def it_round_trips_a_comment_through_save_and_reopen(
        self, prs_with_slide: PresentationClass, _restore_part_factory: None
    ) -> None:
        """Every comment attribute survives ``Presentation.save`` + reopen.

        The authoritative round-trip guarantee: a comment authored
        in-memory, serialized to a ``.pptx`` byte-stream, and reopened
        via :class:`Presentation` must read back identical text,
        author (resolved through the rebuilt package-level registry),
        position, and timestamp. This exercises the part topology from
        ECMA-376 §13.3.3 — the comment-authors part hangs off the
        presentation part, the comments part hangs off the slide.
        """
        slide = prs_with_slide.slides[0]
        comments = slide.comments
        alice = comments._authors.add_author("Alice", "A.")
        bob = comments._authors.add_author("Bob", "B.")
        when_a = dt.datetime(2025, 2, 3, 4, 5, 6)
        when_b = dt.datetime(2025, 4, 5, 6, 7, 8)

        comments.add_comment(alice, "first pass", position=(100, 200), datetime_value=when_a)
        comments.add_comment(bob, "second pass", position=(300, 400), datetime_value=when_b)

        # -- round-trip --
        buf = io.BytesIO()
        prs_with_slide.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        # -- reopened slide still has a comments part --
        assert slide2.has_comments is True

        reloaded = list(slide2.comments)
        assert len(reloaded) == 2

        # -- comment 1 --
        assert reloaded[0].text == "first pass"
        assert reloaded[0].position == (100, 200)
        assert reloaded[0].datetime == when_a
        assert reloaded[0].author is not None
        assert reloaded[0].author.name == "Alice"
        assert reloaded[0].author.initials == "A."

        # -- comment 2 — author registry round-trips too --
        assert reloaded[1].text == "second pass"
        assert reloaded[1].position == (300, 400)
        assert reloaded[1].datetime == when_b
        assert reloaded[1].author is not None
        assert reloaded[1].author.name == "Bob"
        assert reloaded[1].author.initials == "B."

        # -- both authors resolve to distinct registry ids on reopen --
        assert reloaded[0].author_id != reloaded[1].author_id

    # -- 8. Access via presentation.slides[0].comments ------------------

    def it_exposes_comments_off_presentation_slides_indexed_access(self) -> None:
        """``presentation.slides[0].comments`` is the documented public path.

        The #756 reporter wanted a surface along the lines of
        ``presentation.slides[0].comments``; this scenario asserts that
        exact navigation path off a freshly constructed
        :class:`Presentation` returns a :class:`Comments` collection
        suitable for both reading and writing. This is the user-facing
        API the documentation in ``docs/user/comments.rst`` demonstrates.
        """
        presentation = Presentation()
        presentation.slides.add_slide(presentation.slide_layouts[5])

        # -- before any access, has_comments is side-effect-free False --
        assert presentation.slides[0].has_comments is False

        # -- the documented public path --
        comments = presentation.slides[0].comments

        # -- it is a Comments collection, initially empty, with an
        # -- add_comment helper wired to the package-level author registry --
        assert isinstance(comments, Comments)
        assert len(comments) == 0

        author = comments._authors.get_or_add("Reviewer", "R.")
        assert isinstance(author, CommentAuthor)

        comments.add_comment(author, "from the public API")

        # -- reading back through the same navigation path sees it --
        assert presentation.slides[0].has_comments is True
        assert len(presentation.slides[0].comments) == 1
        assert list(presentation.slides[0].comments)[0].text == ("from the public API")
