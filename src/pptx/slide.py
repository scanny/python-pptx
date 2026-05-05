"""Slide-related objects, including masters, layouts, and notes."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING, Iterator, cast

from pptx.dml.color import RGBColor
from pptx.dml.fill import FillFormat
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.transition import (
    PP_TRANSITION_SIDE_DIRECTION,
    PP_TRANSITION_SPEED,
    PP_TRANSITION_TYPE,
)
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import namespaces, qn
from pptx.shapes.shapetree import (
    LayoutPlaceholders,
    LayoutShapes,
    MasterPlaceholders,
    MasterShapes,
    NotesSlidePlaceholders,
    NotesSlideShapes,
    SlidePlaceholders,
    SlideShapes,
)
from pptx.shared import ElementProxy, ParentedElementProxy, PartElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.comments import Comments
    from pptx.opc.package import Part
    from pptx.oxml.presentation import CT_SlideId, CT_SlideIdList, CT_SlideMasterIdList
    from pptx.oxml.slide import (
        CT_Background,
        CT_CommonSlideData,
        CT_NotesMaster,
        CT_NotesSlide,
        CT_Slide,
        CT_SlideLayout,
        CT_SlideLayoutIdList,
        CT_SlideMaster,
    )
    from pptx.oxml.tags import CT_TagList
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart
    from pptx.presentation import Presentation
    from pptx.shapes.base import BaseShape
    from pptx.shapes.placeholder import LayoutPlaceholder, MasterPlaceholder
    from pptx.shapes.shapetree import NotesSlidePlaceholder
    from pptx.text.text import TextFrame


class _BaseSlide(PartElementProxy):
    """Base class for slide objects, including masters, layouts and notes."""

    _element: CT_Slide

    @lazyproperty
    def background(self) -> _Background:
        """|_Background| object providing slide background properties.

        This property returns a |_Background| object whether or not the
        slide, master, or layout has an explicitly defined background.

        The same |_Background| object is returned on every call for the same
        slide object.
        """
        return _Background(self._element.cSld, self)

    @property
    def name(self) -> str:
        """String representing the internal name of this slide.

        Returns an empty string (`''`) if no name is assigned. Assigning an empty string or |None|
        to this property causes any name to be removed.
        """
        return self._element.cSld.name

    @name.setter
    def name(self, value: str | None):
        new_value = "" if value is None else value
        self._element.cSld.name = new_value


class _BaseMaster(_BaseSlide):
    """Base class for master objects such as |SlideMaster| and |NotesMaster|.

    Provides access to placeholders and regular shapes.
    """

    _element: CT_SlideMaster | CT_NotesMaster  # pyright: ignore[reportIncompatibleVariableOverride]

    @lazyproperty
    def header_footer(self) -> _HeaderFooter:
        """|_HeaderFooter| object controlling header/footer/slide-number/date visibility.

        A header/footer setting on a master applies as an inherited default to all layouts /
        slides that descend from it. See |_HeaderFooter| for the individual toggles.
        """
        return _HeaderFooter(self._element)

    @lazyproperty
    def placeholders(self) -> MasterPlaceholders:
        """|MasterPlaceholders| collection of placeholder shapes in this master.

        Sequence sorted in `idx` order.
        """
        return MasterPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> MasterShapes:
        """
        Instance of |MasterShapes| containing sequence of shape objects
        appearing on this slide.
        """
        return MasterShapes(self._element.spTree, self)


class NotesMaster(_BaseMaster):
    """Proxy for the notes master XML document.

    Provides access to shapes, the most commonly used of which are placeholders.
    """


class NotesSlide(_BaseSlide):
    """Notes slide object.

    Provides access to slide notes placeholder and other shapes on the notes handout
    page.
    """

    element: CT_NotesSlide  # pyright: ignore[reportIncompatibleMethodOverride]

    def clone_master_placeholders(self, notes_master: NotesMaster) -> None:
        """Selectively add placeholder shape elements from `notes_master`.

        Selected placeholder shape elements from `notes_master` are added to the shapes
        collection of this notes slide. Z-order of placeholders is preserved. Certain
        placeholders (header, date, footer) are not cloned.
        """

        def iter_cloneable_placeholders() -> Iterator[MasterPlaceholder]:
            """Generate a reference to each cloneable placeholder in `notes_master`.

            These are the placeholders that should be cloned to a notes slide when the a new notes
            slide is created.
            """
            cloneable = (
                PP_PLACEHOLDER.SLIDE_IMAGE,
                PP_PLACEHOLDER.BODY,
                PP_PLACEHOLDER.SLIDE_NUMBER,
            )
            for placeholder in notes_master.placeholders:
                if placeholder.element.ph_type in cloneable:
                    yield placeholder

        shapes = self.shapes
        for placeholder in iter_cloneable_placeholders():
            shapes.clone_placeholder(cast("LayoutPlaceholder", placeholder))

    @property
    def notes_placeholder(self) -> NotesSlidePlaceholder | None:
        """the notes placeholder on this notes slide, the shape that contains the actual notes text.

        Return |None| if no notes placeholder is present; while this is probably uncommon, it can
        happen if the notes master does not have a body placeholder, or if the notes placeholder
        has been deleted from the notes slide.
        """
        for placeholder in self.placeholders:
            if placeholder.placeholder_format.type == PP_PLACEHOLDER.BODY:
                return placeholder
        return None

    @property
    def notes_text_frame(self) -> TextFrame | None:
        """The text frame of the notes placeholder on this notes slide.

        |None| if there is no notes placeholder. This is a shortcut to accommodate the common case
        of simply adding "notes" text to the notes "page".
        """
        notes_placeholder = self.notes_placeholder
        if notes_placeholder is None:
            return None
        return notes_placeholder.text_frame

    @lazyproperty
    def placeholders(self) -> NotesSlidePlaceholders:
        """Instance of |NotesSlidePlaceholders| for this notes-slide.

        Contains the sequence of placeholder shapes in this notes slide.
        """
        return NotesSlidePlaceholders(self.element.spTree, self)

    @lazyproperty
    def shapes(self) -> NotesSlideShapes:
        """Sequence of shape objects appearing on this notes slide."""
        return NotesSlideShapes(self._element.spTree, self)


class Slide(_BaseSlide):
    """Slide object. Provides access to shapes and slide-level properties."""

    part: SlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    @lazyproperty
    def comments(self) -> Comments:
        """|Comments| collection of legacy review comments anchored on this slide.

        If the slide does not yet have a comments part, a new one is created (and,
        if necessary, the package-level comment-authors part as well). The same
        |Comments| instance is returned on subsequent calls.

        .. versionadded:: 2026.05.0
        """
        # -- local import avoids circular import between slide and comments --
        from pptx.comments import CommentAuthors, Comments

        package = self.part.package
        authors_part = package.get_or_add_comment_authors_part()
        authors = CommentAuthors(authors_part)
        comments_part = self.part.get_or_add_comments_part()
        return Comments(self, comments_part, authors)

    @property
    def has_comments(self) -> bool:
        """`True` if this slide has a legacy comments part, `False` otherwise.

        A comments part is created by :attr:`comments` when first accessed; use this
        property to test whether a comments part already exists without the side
        effect of creating one.

        .. versionadded:: 2026.05.0
        """
        return self.part.comments_part is not None

    @property
    def has_tags(self) -> bool:
        """`True` when this slide has a VBA-style custom-tags part attached.

        A tags part is created lazily on the first *mutation* through
        :attr:`tags` (e.g. ``slide.tags["foo"] = "bar"``). Read-only access
        through :attr:`tags` is side-effect free; use this property to
        distinguish a slide that currently has no tags from one that has
        an empty tags part.

        .. versionadded:: 2026.05.0
        """
        return self.part.tags_part is not None

    @lazyproperty
    def tags(self) -> SlideTags:
        """Dict-like |SlideTags| proxy for this slide's custom name/value tags.

        Tags are the VBA-style string name/value pairs PowerPoint surfaces
        through ``Slide.Tags`` (see issue #578). They are stored in a
        companion |TagsPart| that is created lazily the first time a tag is
        written. Read access through this property never materializes a
        tags part; an empty proxy is returned when one does not exist yet.

        Typical usage::

            slide.tags["priority"] = "high"   # writes tag
            slide.tags["priority"]            # returns "high"
            "priority" in slide.tags          # True
            del slide.tags["priority"]        # removes tag
            list(slide.tags)                  # ["priority", ...]

        The returned object supports ``__getitem__`` / ``__setitem__`` /
        ``__delitem__`` / ``__contains__`` / ``__iter__`` / ``__len__``
        plus the read-only ``get`` convenience method, matching the
        most-frequently-used subset of the built-in ``dict`` API.

        .. versionadded:: 2026.05.0
        """
        return SlideTags(self.part)

    @property
    def effective_background(self) -> _EffectiveBackground | None:
        """Resolved :class:`_EffectiveBackground` from the inheritance chain, or |None|.

        Walks ``slide → layout → master`` in that order and returns a
        read-only |_EffectiveBackground| view of the first ancestor that
        carries an explicit ``p:bg`` child. Returns |None| when no ancestor
        declares a background (a rare case — PowerPoint-authored decks
        almost always carry a master-level ``p:bg``).

        Complements the side-effect-prone :attr:`Slide.background` accessor
        (whose :attr:`~_Background.fill` materializes ``p:bgPr/a:noFill``
        on first read, clobbering inheritance). :attr:`effective_background`
        lets callers **read** the rendered background without disturbing
        the XML — in particular::

            >>> slide.effective_background.fill.fore_color.rgb  # safe

        resolves the inherited layout / master color when the slide has
        no ``p:bg`` of its own. Fixes the round-trip read path reported in
        issue #809.

        .. versionadded:: 2026.05.0
        """
        # -- 1. slide itself --
        if self._element.cSld.bg is not None:
            return _EffectiveBackground("slide", self)
        # -- 2. layout --
        try:
            layout = self.slide_layout
        except (AttributeError, KeyError):  # pragma: no cover - defensive
            return None
        if layout._element.cSld.bg is not None:  # pyright: ignore[reportPrivateUsage]
            return _EffectiveBackground("layout", layout)
        # -- 3. master --
        try:
            master = layout.slide_master
        except (AttributeError, KeyError):  # pragma: no cover - defensive
            return None
        if master._element.cSld.bg is not None:  # pyright: ignore[reportPrivateUsage]
            return _EffectiveBackground("master", master)
        return None

    @property
    def follow_master_background(self) -> _FollowMasterBackground:
        """|True| if this slide inherits the slide master background.

        Returns a |bool|-like object that is also *callable*: read it for a
        boolean, **or** call it to reset a custom background back to the
        inherited master background.

        Reading returns a truthy value when the slide has no ``p:bg``
        child (so the background is inherited from its layout / master)
        and a falsy value when the slide carries its own explicit
        background::

            >>> if slide.follow_master_background:
            ...     # slide inherits from the master
            ...     ...

        Calling it (equivalent to PowerPoint's *Reset Background*
        button) removes any ``p:bg`` child from the slide's ``p:cSld``,
        so the slide once again inherits from its layout / master::

            >>> slide.follow_master_background()  # revert to master inheritance

        A no-op on a slide that already follows the master background.
        Returns the slide for call chaining.

        .. versionchanged:: 2026.05.0
           The returned object is now also callable, providing a
           ``slide.follow_master_background()`` reset method in addition
           to the pre-existing read-as-bool behavior.
        """
        return _FollowMasterBackground(self)

    def copy_background_from(self, source_slide: _BaseSlide) -> None:
        """Deep-copy the background markup of `source_slide` onto this slide.

        Any explicit background on this slide is replaced. If `source_slide`
        has no explicit ``p:bg`` (i.e. it inherits from its master / layout),
        any explicit background on this slide is removed and background
        inheritance is restored.

        `source_slide` may be any slide-like object — a |Slide|, |SlideLayout|,
        or |SlideMaster| — since each carries a ``p:cSld/p:bg`` in the same
        shape. The copy is purely XML-level: a ``p:bgRef`` style reference
        keyed on the master theme is copied as-is without rewriting the
        theme reference, so copying between presentations may yield
        unresolved style references.

        .. versionadded:: 2026.05.0
        """
        cSld = self._element.cSld
        cSld._remove_bg()  # pyright: ignore[reportPrivateUsage]
        source_bg = source_slide._element.cSld.bg
        if source_bg is not None:
            cSld._insert_bg(copy.deepcopy(source_bg))  # pyright: ignore[reportPrivateUsage]

    @property
    def is_hidden(self) -> bool:
        """`True` when this slide is marked hidden in the presentation.

        Reflects the ``p:sld/@show`` attribute. PowerPoint writes
        ``show="0"`` on a slide marked as hidden via the *Hide Slide*
        command; a visible slide omits the attribute (the schema default
        for ``@show`` is ``true``). Accordingly:

        * Reading returns ``True`` when ``@show="0"`` is present and
          ``False`` otherwise (absent attribute or any truthy value).
        * Assigning ``True`` writes ``show="0"`` on the ``p:sld`` element.
        * Assigning ``False`` removes the ``@show`` attribute (if present)
          so the XML round-trips to the default-visible state.

        Hidden slides are skipped by PowerPoint during a normal slide-show
        run but remain in the package and in :attr:`.Presentation.slides`.
        """
        return not self._element.show

    @is_hidden.setter
    def is_hidden(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise TypeError("is_hidden must be a bool, got %s" % type(value).__name__)
        # -- CT_Slide.show is an OptionalAttribute(default=True); assigning
        # -- True (the default) removes the attribute, assigning False
        # -- writes show="0".
        self._element.show = not value

    @property
    def show_master_shapes(self) -> bool:
        """`True` when shapes on the master are shown on this slide.

        Reflects the ``p:sld/@showMasterSp`` attribute. PowerPoint's
        *Hide Background Graphics* checkbox (under *Design > Format
        Background*) writes ``showMasterSp="0"`` on a slide that opts out
        of rendering the master-level decorative shapes (e.g. a company
        logo, a footer bar, or any other shape baked into the master's
        shape tree). Master *placeholders* continue to inherit normally;
        only non-placeholder master shapes are affected by this toggle.

        * Reading returns ``True`` when the attribute is absent or any
          truthy value and ``False`` only when ``@showMasterSp="0"``.
        * Assigning ``True`` (the schema default) removes the attribute
          if present so the XML round-trips to the default-visible state.
        * Assigning ``False`` writes ``showMasterSp="0"`` on the ``p:sld``
          element.
        """
        return self._element.showMasterSp

    @show_master_shapes.setter
    def show_master_shapes(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise TypeError("show_master_shapes must be a bool, got %s" % type(value).__name__)
        # -- CT_Slide.showMasterSp is an OptionalAttribute(default=True);
        # -- assigning True (the default) removes the attribute, assigning
        # -- False writes showMasterSp="0".
        self._element.showMasterSp = value

    @property
    def has_notes_slide(self) -> bool:
        """`True` if this slide has a notes slide, `False` otherwise.

        A notes slide is created by :attr:`.notes_slide` when one doesn't exist; use this property
        to test for a notes slide without the possible side effect of creating one.
        """
        return self.part.has_notes_slide

    @property
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance for this slide.

        If the slide does not have a notes slide, one is created. The same single instance is
        returned on each call.
        """
        return self.part.notes_slide

    @lazyproperty
    def placeholders(self) -> SlidePlaceholders:
        """Sequence of placeholder shapes in this slide."""
        return SlidePlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> SlideShapes:
        """Sequence of shape objects appearing on this slide."""
        return SlideShapes(self._element.spTree, self)

    def clear_shapes(self, preserve_placeholders: bool = True) -> None:
        """Remove every top-level shape from this slide.

        Convenience wrapper around :meth:`SlideShapes.clear`. By default
        placeholders are preserved so the slide continues to inherit from
        its layout; pass ``preserve_placeholders=False`` to remove every
        shape including placeholders::

            slide.clear_shapes()                         # keep placeholders
            slide.clear_shapes(preserve_placeholders=False)  # wipe everything

        Returns ``None`` to match :meth:`list.clear`.

        .. versionadded:: 2026.05.0
        """
        self.shapes.clear(preserve_placeholders=preserve_placeholders)

    def clone_shapes_from(self, other_slide: Slide, include_placeholders: bool = True) -> None:
        """Clone every top-level shape from `other_slide` onto this slide.

        Walks `other_slide.shapes` in document (z-order) and appends a deep
        copy of each shape to this slide's shape tree via
        :meth:`.BaseShape.clone_onto`. Every referenced package part (image
        blobs, chart parts, OLE payloads, SmartArt four-part subgraphs,
        3D-model media, hyperlinks) is re-embedded on this slide's package
        so the result is fully self-contained, and each clone's
        ``cNvPr/@id`` and ``cNvPr/@name`` are made unique against this
        slide's shape tree. Source-slide z-order is preserved.

        Placeholders are handled separately because CLO-2
        (:meth:`.BaseShape.clone_onto`) raises ``NotImplementedError`` on
        placeholders — a placeholder's ``idx`` is inherited from its layout
        and duplicating the element would break the "one shape per idx"
        invariant. When ``include_placeholders`` is ``True`` (the default),
        each placeholder's text body is deep-copied onto this slide's
        matching-``idx`` placeholder (when one exists) one paragraph at a
        time via :meth:`._Paragraph.clone_from`, so run-level formatting
        transfers. When this slide has no placeholder with the same ``idx``
        the source placeholder is skipped (no new placeholder is created).
        Non-text placeholder content (e.g. a picture placeholder's image)
        is *not* cloned in this MVP; only the text body transfers. When
        ``include_placeholders`` is ``False`` every placeholder on the
        source slide is skipped entirely.

        `other_slide` is not mutated. Returns ``None`` — this is a
        mutation method on ``self``.

        .. versionadded:: 2026.05.0
        """
        for shape in other_slide.shapes:
            if shape.is_placeholder:
                if not include_placeholders:
                    continue
                # -- Placeholders holding chart / table / picture content  --
                # -- carry that content as ``<p:graphicFrame>`` or         --
                # -- ``<p:pic>`` XML that is structurally clone-safe even  --
                # -- though the element is flagged as a placeholder. Strip --
                # -- the ``<p:ph>`` marker on the clone and route it       --
                # -- through ``clone_onto`` so the chart / table / image   --
                # -- travels to the target fully re-embedded.              --
                if shape.has_chart or shape.has_table or shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    _clone_non_text_placeholder_onto(shape, self.shapes)
                    continue
                # -- Text-only placeholders fall back to per-paragraph     --
                # -- text-body clone when this slide has a matching-idx    --
                # -- placeholder.                                          --
                try:
                    src_idx = shape.placeholder_format.idx
                except ValueError:  # pragma: no cover -- defensive
                    continue
                tgt_ph = next(
                    (p for p in self.placeholders if p.placeholder_format.idx == src_idx),
                    None,
                )
                if tgt_ph is None or not shape.has_text_frame or not tgt_ph.has_text_frame:
                    continue
                # -- ``text_frame`` only exists on subclasses that override
                # -- ``has_text_frame`` to True (i.e. ``Shape`` and placeholder
                # -- subclasses); narrow the base-class type here.
                src_tf = cast("TextFrame", getattr(shape, "text_frame"))
                tgt_tf = cast("TextFrame", getattr(tgt_ph, "text_frame"))
                _clone_text_frame(src_tf, tgt_tf)
                continue
            shape.clone_onto(self.shapes)

    @property
    def shape_tree_flat(self) -> Iterator[BaseShape]:
        """Iterator over every shape on this slide, including descendants of groups.

        Mirrors PowerPoint's Selection Pane: each shape PowerPoint would list —
        top-level shapes plus every shape nested inside a :class:`.GroupShape`
        at arbitrary depth — is yielded in document (z-order) sequence. Group
        shapes themselves are yielded *before* their children so the caller
        sees the container alongside its contents.

        Convenience wrapper around :meth:`SlideShapes.descendants`; use the
        latter directly when you need it on a layout, master, or notes slide.

        .. versionadded:: 2026.05.0
        """
        return self.shapes.descendants()

    def find_shapes_by_xpath(self, xpath_expr: str) -> list[BaseShape]:
        """Return shapes matching `xpath_expr` evaluated against this slide's ``p:spTree``.

        The expression is evaluated with the standard Open-XML namespace map
        (``pptx.oxml.ns._nsmap``), so the usual prefixes (``p``, ``a``, ``r``,
        ``mc``, ``p14``, …) are available without further declaration. The
        expression is rooted at the slide's ``p:spTree`` element; start with
        ``.//`` to search the whole tree, or use a relative path (``p:sp``,
        etc.) to match direct children.

        Each match that is itself a shape element (``p:sp``, ``p:pic``,
        ``p:cxnSp``, ``p:graphicFrame``, or ``p:grpSp``) is wrapped in the
        appropriate :class:`~pptx.shapes.base.BaseShape` subclass via
        :func:`.SlideShapeFactory` — the same proxy kind
        :attr:`Slide.shapes` would yield for that element. Matches that are
        not shape elements (for example, the ``a:cNvPr`` or ``a:xfrm`` of a
        shape) are resolved to their nearest shape ancestor so the caller
        gets a shape proxy back regardless of how narrow the XPath is.
        Non-shape matches with no shape ancestor (e.g. the ``p:spTree``
        itself) are skipped.

        Duplicates are suppressed: each shape appears at most once in the
        returned list, in XPath-result order, even when the same shape is
        matched by multiple predicates (for instance a ``./@name`` selector
        that also matches a child element).

        Example - find every shape whose ``@name`` is ``"Title 1"``::

            shapes = slide.find_shapes_by_xpath(
                './/p:sp[p:nvSpPr/p:cNvPr/@name="Title 1"]'
            )

        Example - every non-group shape anywhere in the tree::

            shapes = slide.find_shapes_by_xpath(".//p:sp")

        .. versionadded:: 2026.05.0
        """
        from pptx.shapes.shapetree import SlideShapeFactory

        results = self._element.spTree.xpath(xpath_expr)

        seen: set[int] = set()
        shapes: list[BaseShape] = []
        for match in results:
            # -- `xpath` can return strings / numbers / booleans for value
            # -- expressions (e.g. `.//@name`). Skip anything that isn't an
            # -- element node — a shape cannot be wrapped around an
            # -- attribute / text value.
            if not hasattr(match, "tag"):
                continue
            shape_elm = match if _is_shape_elm(match) else _nearest_shape_ancestor(match)
            if shape_elm is None:
                continue
            key = id(shape_elm)
            if key in seen:
                continue
            seen.add(key)
            shapes.append(SlideShapeFactory(shape_elm, self.shapes))
        return shapes

    @property
    def slide_id(self) -> int:
        """Integer value that uniquely identifies this slide within this presentation.

        The slide id does not change if the position of this slide in the slide sequence is changed
        by adding, rearranging, or deleting slides.
        """
        return self.part.slide_id

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object this slide inherits appearance from."""
        return self.part.slide_layout

    @slide_layout.setter
    def slide_layout(self, slide_layout: SlideLayout) -> None:
        """Re-point this slide at `slide_layout`.

        Replaces the slide's underlying ``RT.SLIDE_LAYOUT`` relationship so
        the slide inherits appearance from `slide_layout`. Assigning a
        layout from a different slide master ("cross-master" swap) is
        supported; the slide's own shape tree is not rewritten so any
        placeholder / theme references the slide inherits through the new
        layout's master take effect via the inheritance chain.

        Raises |ValueError| when `slide_layout` is not part of the same
        presentation as this slide.

        Note that the new layout must belong to the same presentation
        (package); to import a layout from another presentation, use
        :meth:`SlideMaster.add_layout` / :meth:`SlideMaster.add_layout_from`
        first and then assign the returned layout here.

        .. versionadded:: 2026.05.0
        """
        self.part.slide_layout = slide_layout

    @property
    def has_animations(self) -> bool:
        """`True` when this slide carries any ``p:timing`` / animation XML.

        Returns `True` if the slide's XML has a ``p:timing`` child
        element containing at least one time-node (``p:tnLst/*``).
        Media-only timing (the minimal ``p:timing`` tree added by
        :meth:`SlideShapes.add_movie` for video playback controls) is
        considered an "animation" here because python-pptx does not
        distinguish media-playback from authored animation until
        downstream item #954 lands the mc:AlternateContent handling.

        This property is side-effect free — unlike :attr:`transition`
        it does not create a ``p:timing`` element on access.

        Structured read/write access to the animation tree (entrance /
        exit / emphasis / motion-path / MORPH effects) is deferred to
        downstream items #102, #264, #861, and #1106. See
        ``docs/dev/analysis/f8-animations-transitions.rst``.

        .. versionadded:: 2026.05.0
        """
        timing = self._element.timing
        if timing is None:
            return False
        tnLst = timing.tnLst
        if tnLst is None:
            return False
        return len(tnLst) > 0

    @property
    def timing_xml(self) -> str | None:
        """XML string of the slide's ``p:timing`` subtree, or ``None`` when absent.

        Read-only debugging accessor included in the F8 MVP so code
        that needs to introspect or hand-patch the timing XML (e.g.
        for the animation items stacked on F8) can do so without
        reaching into ``slide._element``. The returned value is the
        serialized XML exactly as stored.

        .. versionadded:: 2026.05.0
        """
        timing = self._element.timing
        if timing is None:
            return None
        return timing.xml

    @property
    def animation_sequence(self) -> tuple[AnimationEffectView, ...]:
        """Read-only tuple of |AnimationEffectView| in this slide's main sequence.

        Returns each entrance / emphasis / exit / motion-path effect
        PowerPoint would step through when the slide's main animation
        sequence runs, in document order.

        Returns an empty tuple when the slide has no ``p:timing`` subtree
        or when the subtree is media-playback-only (the stub tree added
        by :meth:`.SlideShapes.add_movie` has no main sequence).

        Each :class:`.AnimationEffectView` exposes a handful of read-only
        accessors — ``shape_id``, ``preset_class``, ``preset_id``,
        ``preset_subtype``, ``delay`` — that are enough to describe
        PowerPoint-authored effects for introspection and
        test-harness purposes. Structured authoring (add / remove /
        reorder / retarget effects) is out of scope for this MVP;
        see downstream items #102 (Shape.animation entrance effects),
        #264 (full animation-tree authoring), and #1106 (broader
        entrance/exit/emphasis API) which will layer mutation methods
        on top of this read surface.

        This property is side-effect free — it does not create a
        ``p:timing`` element on access. See
        ``docs/dev/analysis/f8-animations-transitions.rst`` for the
        full animation-tree design map.

        .. versionadded:: 2026.05.0
        """
        from pptx.oxml.timing import iter_main_sequence_effects

        timing = self._element.timing
        if timing is None:
            return ()
        return tuple(AnimationEffectView(par) for par in iter_main_sequence_effects(timing))

    def iter_shape_animations(self) -> Iterator[ShapeAnimation]:
        """Iterate over the shape-targeted animation effects on this slide.

        Yields a :class:`.ShapeAnimation` proxy for each behaviour node
        in the slide's ``p:timing`` subtree that targets a shape via
        ``p:spTgt/@spid`` — namely ``p:anim``, ``p:animEffect``,
        ``p:animMotion``, ``p:animRot``, ``p:animScale``, ``p:animClr``,
        and ``p:set``.

        This is the **read-only MVP** for issue #264. It lets callers
        introspect authored animations (delay, duration, trigger shape,
        effect type) without reaching into ``slide.timing_xml``. Write /
        authoring support — adding entrance / exit / emphasis / motion
        effects, or modifying an effect's delay in-place — is deferred
        to the larger Wave-7 lift (#264 full, #102, #861, #1106). See
        ``docs/dev/analysis/f8-animations-transitions.rst`` for the
        extension-point map.

        .. versionadded:: 2026.05.0
        """
        timing = self._element.timing
        if timing is None:
            return
        tnLst = timing.tnLst
        if tnLst is None:
            return
        # -- Find every behaviour element under this tnLst that targets a
        # -- shape via `p:spTgt/@spid`. Using descendant XPath keeps us
        # -- robust to arbitrary wrapping (`p:par`, `p:seq`, nested
        # -- `p:childTnLst`, etc.) without hard-coding the main-sequence
        # -- structure.
        for spTgt in _xpath_p(tnLst, ".//p:spTgt[@spid]"):
            # -- Walk up to the nearest behaviour ancestor whose tag is one
            # -- of the known shape-animation effect types. `p:tgtEl` may
            # -- live inside a shared `p:cBhvr` / `p:cMediaNode`; we want
            # -- the behaviour element itself (p:anim, p:animEffect, ...).
            effect = _ancestor_with_local_name(
                spTgt,
                (
                    "anim",
                    "animEffect",
                    "animMotion",
                    "animRot",
                    "animScale",
                    "animClr",
                    "set",
                ),
            )
            if effect is None:
                continue
            yield ShapeAnimation(effect, int(spTgt.get("spid")))

    @lazyproperty
    def transition(self) -> Transition:
        """|Transition| proxy for reading / writing this slide's transition.

        Returns a :class:`.Transition` object whether or not the slide
        has an explicit ``p:transition`` element. The returned proxy
        lazily adds a ``p:transition`` child the first time a setter
        is invoked; merely calling :attr:`transition.type` is
        non-destructive.

        The MVP API surfaces:

        * :attr:`Transition.type` — :class:`PP_TRANSITION_TYPE`
          (fade / wipe / push / cover / ... / morph).
        * :attr:`Transition.duration` — int milliseconds or ``None``.
        * :attr:`Transition.advance_on_click` — bool (``@advClick``).
        * :attr:`Transition.advance_after_time` — int ms or ``None``
          (``@advTm``).

        Extended API (structured animation effects, advance-after-time
        on media, MORPH option selection) is deferred — see
        ``docs/dev/analysis/f8-animations-transitions.rst`` and the
        downstream-items matrix.

        .. versionadded:: 2026.05.0
        """
        return Transition(self._element)


class Slides(ParentedElementProxy):
    """Sequence of slides belonging to an instance of |Presentation|.

    Has list semantics for access to individual slides. Supports indexed access, len(), and
    iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldIdLst: CT_SlideIdList, prs: Presentation):
        super(Slides, self).__init__(sldIdLst, prs)
        self._sldIdLst = sldIdLst

    def __getitem__(self, idx: int) -> Slide:
        """Provide indexed access, (e.g. 'slides[0]')."""
        try:
            sldId = self._sldIdLst.sldId_lst[idx]
        except IndexError:
            raise IndexError("slide index out of range")
        return self.part.related_slide(sldId.rId)

    def __iter__(self) -> Iterator[Slide]:
        """Support iteration, e.g. `for slide in slides:`."""
        for sldId in self._sldIdLst.sldId_lst:
            yield self.part.related_slide(sldId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldIdLst)

    def add_slide(self, slide_layout: SlideLayout, index: int | None = None) -> Slide:
        """Return a newly added slide that inherits layout from `slide_layout`.

        `index` is the zero-based position at which the new slide should be
        inserted in this slide sequence. When omitted (or ``None``), the slide
        is appended to the end (the historical behavior). A negative `index`
        counts from the end in the usual Python way; an `index` beyond the end
        is clamped to the last position (matching :meth:`move_slide` and
        :meth:`duplicate`).

        The ``index`` keyword was added in 2026.05.0 (#194).
        """
        rId, slide = self.part.add_slide(slide_layout)
        slide.shapes.clone_layout_placeholders(slide_layout)
        # -- Always append the new p:sldId first so a fresh slide-id is allocated
        # -- against the full set of existing ids. Then reposition if needed. --
        new_sldId = self._sldIdLst.add_sldId(rId)
        if index is not None:
            self._reposition_sldId(new_sldId, index)
        return slide

    def duplicate(self, slide: Slide, index: int | None = None) -> Slide:
        """Return a newly added slide that is a duplicate of `slide`.

        `slide` must be a member of this collection. The duplicate has a deep
        copy of `slide`'s shape tree and inherits from the same slide layout.
        Image, chart, OLE-object, media, and hyperlink relationships on the
        source slide are cloned onto the duplicate: each target part is
        *reused* (a duplicate slide and its source share the same image,
        chart, etc. parts by package reference), while each relationship-id
        is freshly allocated on the new slide.

        `index` is the zero-based position at which the duplicate should be
        inserted in this slide sequence. When omitted (or ``None``), the
        duplicate is appended to the end. A negative `index` counts from the
        end in the usual Python way; an `index` beyond the end is clamped to
        the last position (matching :meth:`move_slide`).

        The duplicate receives a freshly-allocated `slide_id`; the source
        slide's `slide_id` is unchanged.

        Any notes-slide attached to `slide` is *not* copied onto the
        duplicate, because a notes slide carries a back-reference to its
        owning slide. The returned slide has no notes slide attached;
        accessing :attr:`Slide.notes_slide` will create a fresh one on
        demand.

        Raises |ValueError| if `slide` is not a member of this collection.

        .. versionadded:: 2026.05.0
        """
        # -- Locate `slide` in this collection (raises ValueError on miss). --
        self.index(slide)

        # -- Create the duplicate slide part and register it with this presentation. --
        rId, new_slide = self.part.duplicate_slide(slide)

        # -- Always append the new `p:sldId` first so a fresh slide-id is allocated
        # -- against the full set of existing ids. Then reposition if needed. --
        new_sldId = self._sldIdLst.add_sldId(rId)

        if index is not None:
            self._reposition_sldId(new_sldId, index)

        return new_slide

    def delete(self, slide: Slide) -> None:
        """Remove `slide` from this presentation.

        The corresponding `p:sldId` entry is removed from the `p:sldIdLst` and the
        presentation-part's relationship to the slide-part is dropped. Parts that become
        unreachable from the package root (the slide part itself, and any image, media,
        or chart parts only referenced by that slide) are omitted on the next save, so
        they are effectively garbage-collected.

        Parts shared with other slides (for example a chart or image referenced from
        another slide) are preserved because they remain reachable from the package.

        Subsequent use of `slide` is undefined; most operations will raise an exception.

        Raises |ValueError| if `slide` is not a member of this collection.

        .. versionadded:: 2026.05.0
        """
        # -- locate the p:sldId entry for `slide` --
        target_idx = self.index(slide)
        target_sldId = self._sldIdLst.sldId_lst[target_idx]
        rId = target_sldId.rId

        # -- remove the p:sldId from p:sldIdLst so nothing in presentation.xml references
        # -- the slide. This must happen *before* drop_rel so the reference-count check
        # -- in drop_rel sees zero remaining references to rId.
        self._sldIdLst.remove(target_sldId)

        # -- drop the presentation-part -> slide-part relationship. XmlPart.drop_rel only
        # -- removes the rel when its reference count in the part's XML is under 2, which
        # -- it is now (we just removed the only reference). Dropping the rel makes the
        # -- slide part, and everything transitively reachable only from it (notes slide,
        # -- uniquely-referenced images, media, charts), unreachable from the package root
        # -- and therefore omitted from the saved package.
        self.part.drop_rel(rId)

    def add_slide_from_external(self, source_slide: Slide, slide_layout: SlideLayout) -> Slide:
        """Return a new slide cloned from `source_slide` in another presentation.

        The new slide is appended to this collection and bound to
        `slide_layout`, which must belong to the target presentation (the one
        owning this |Slides| object). `source_slide` may come from any
        |Presentation| instance, including this one.

        Provides full-fidelity cross-presentation slide copy: image, media,
        chart (with distinct embedded workbook), embedded OLE-object, and
        external hyperlink relationships are all re-established on the cloned
        slide against the target presentation's package. The notes-slide
        relationship is dropped (a notes slide carries a back-reference to
        its owning slide). See :meth:`Presentation.merge` for appending every
        slide of another presentation in one call.

        .. versionadded:: 2026.05.0
        """
        rId, slide = self.part.add_slide_from_external(source_slide, slide_layout)
        self._sldIdLst.add_sldId(rId)
        return slide

    def get(self, slide_id: int, default: Slide | None = None) -> Slide | None:
        """Return the slide identified by int `slide_id` in this presentation.

        Returns `default` if not found.
        """
        slide = self.part.get_slide(slide_id)
        if slide is None:
            return default
        return slide

    def get_by_slide_id(self, slide_id: int, default: Slide | None = None) -> Slide | None:
        """Return the |Slide| having presentation-stable id `slide_id`.

        `slide_id` is matched against ``p:sldIdLst/p:sldId/@id``. The id is
        stable across slide reordering, so this lookup is the robust
        alternative to ``presentation.slides[index]`` when a caller needs to
        reference a specific slide that may have been moved within the deck.
        Returns `default` (``None`` by default) when no entry has a matching
        id. Equivalent to :meth:`Slides.get`, with the explicit name chosen
        to mirror :meth:`SlideMaster.get_layout` and
        :meth:`SlideLayouts.get_by_id`.

        .. versionadded:: 2026.05.0
        """
        return self.get(slide_id, default)

    def index(self, slide: Slide) -> int:
        """Map `slide` to its zero-based position in this slide sequence.

        Raises |ValueError| on *slide* not present.
        """
        for idx, this_slide in enumerate(self):
            if this_slide == slide:
                return idx
        raise ValueError("%s is not in slide collection" % slide)

    def move_slide(self, slide: Slide, new_idx: int) -> None:
        """Move `slide` to zero-based position `new_idx` in this slide sequence.

        The `slide_id` of `slide` is unchanged by this operation. A `new_idx` of 0 moves
        `slide` to the beginning; a `new_idx` equal to or greater than `len(slides) - 1`
        moves it to the end. Negative values count from the end in the usual Python way.

        Raises |ValueError| if `slide` is not a member of this collection.

        .. versionadded:: 2026.05.0
        """
        # -- locate the corresponding p:sldId child element --
        current_idx = self.index(slide)
        sldId = self._sldIdLst.sldId_lst[current_idx]

        self._reposition_sldId(sldId, new_idx)

    def _reposition_sldId(self, sldId: CT_SlideId, new_idx: int) -> None:
        """Move `sldId` to zero-based position `new_idx` within `p:sldIdLst`.

        Index normalization matches Python list semantics: negative values count from the
        end, and out-of-range values clamp to the first / last position. This helper is
        shared by :meth:`move_slide` and :meth:`duplicate` so that both operations
        interpret `new_idx` identically. It does not allocate or alter the ``id``
        attribute of `sldId` — sldId-value allocation is exclusively the job of
        :meth:`CT_SlideIdList.add_sldId`.
        """
        # -- Normalize `new_idx` against the current sldIdLst length. Both callers hold a
        # -- live reference to `sldId` that is already a child of self._sldIdLst, so `n`
        # -- already counts it. --
        n = len(self._sldIdLst)
        target_idx = max(0, new_idx + n) if new_idx < 0 else min(new_idx, n - 1)

        # -- no-op when `sldId` is already at the target position --
        if self._sldIdLst.sldId_lst.index(sldId) == target_idx:
            return

        # -- reposition the p:sldId element within p:sldIdLst --
        self._sldIdLst.remove(sldId)
        siblings = self._sldIdLst.sldId_lst
        if target_idx >= len(siblings):
            self._sldIdLst.append(sldId)
        else:
            siblings[target_idx].addprevious(sldId)


class SlideLayout(_BaseSlide):
    """Slide layout object.

    Provides access to placeholders, regular shapes, and slide layout-level properties.
    """

    _element: CT_SlideLayout  # pyright: ignore[reportIncompatibleVariableOverride]
    part: SlideLayoutPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def effective_background(self) -> _EffectiveBackground | None:
        """Resolved :class:`_EffectiveBackground` from the layout inheritance chain, or |None|.

        Walks ``layout → master`` in that order and returns a read-only
        |_EffectiveBackground| view of the first ancestor that carries an
        explicit ``p:bg`` child. Returns |None| when neither the layout
        nor its master declares a background.

        See :attr:`Slide.effective_background` for the slide-level
        analogue and the ``p:bg`` / ``p:bgRef`` semantics. This accessor
        is side-effect free — unlike :attr:`SlideLayout.background`.

        .. versionadded:: 2026.05.0
        """
        if self._element.cSld.bg is not None:
            return _EffectiveBackground("layout", self)
        try:
            master = self.slide_master
        except (AttributeError, KeyError):  # pragma: no cover - defensive
            return None
        if master._element.cSld.bg is not None:  # pyright: ignore[reportPrivateUsage]
            return _EffectiveBackground("master", master)
        return None

    @lazyproperty
    def header_footer(self) -> _HeaderFooter:
        """|_HeaderFooter| object controlling header/footer/slide-number/date visibility.

        When absent, the layout inherits its master's header/footer settings; accessing this
        property does not itself add a `p:hf` element to the layout XML. See |_HeaderFooter|.

        .. versionadded:: 2026.05.0
        """
        return _HeaderFooter(self._element)

    def iter_cloneable_placeholders(self) -> Iterator[LayoutPlaceholder]:
        """Generate layout-placeholders on this slide-layout that should be cloned to a new slide.

        Used when creating a new slide from this slide-layout.
        """
        latent_ph_types = (
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        )
        for ph in self.placeholders:
            if ph.element.ph_type not in latent_ph_types:
                yield ph

    @property
    def name(self) -> str:
        """String representing the internal name of this slide layout.

        PowerPoint writes a descriptive name on every layout (e.g. ``"Title Slide"``,
        ``"Blank"``). Google Slides exports frequently leave ``p:cSld/@name`` empty
        or assign non-standard names, defeating lookups like
        ``slide_layouts.get_by_name("Blank")`` (see issue #864). To keep
        ``SlideLayout.name`` a usable identifier, this property falls back to a
        positional name of the form ``"Layout N"`` (1-based index in the master's
        ``slide_layouts`` collection) when no explicit name is set. Assigning a
        non-empty string persists it to ``p:cSld/@name`` and subsequent reads
        return that explicit value verbatim. Assigning an empty string or |None|
        clears the attribute and restores the positional fallback.
        """
        explicit_name = self._element.cSld.name
        if explicit_name:
            return explicit_name
        return self._positional_name()

    @name.setter
    def name(self, value: str | None):
        new_value = "" if value is None else value
        self._element.cSld.name = new_value

    @lazyproperty
    def placeholders(self) -> LayoutPlaceholders:
        """Sequence of placeholder shapes in this slide layout.

        Placeholders appear in `idx` order.
        """
        return LayoutPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> LayoutShapes:
        """Sequence of shapes appearing on this slide layout."""
        return LayoutShapes(self._element.spTree, self)

    @property
    def slide_layout_id(self) -> int | None:
        """Presentation-stable integer id of this layout, or ``None`` when absent.

        The id is drawn from the ``p:sldLayoutId/@id`` attribute on the matching
        entry in the parent master's ``p:sldLayoutIdLst``. PowerPoint assigns
        this id when the layout is first created and preserves it across
        reordering, so it is the stable handle for referencing a specific
        layout by id rather than by position. Returns ``None`` when the layout
        is detached from a master, when no matching entry can be located, or
        when the entry omits the ``@id`` attribute (optional per
        ECMA-376 Part 1 §19.3.1.41).

        Paired with :meth:`SlideMaster.get_layout` and
        :meth:`SlideLayouts.get_by_id` for id-based lookup.

        .. versionadded:: 2026.05.0
        """
        part = self.part
        if part is None:
            return None
        try:
            master = part.slide_master
        except (AttributeError, KeyError):  # pragma: no cover - defensive
            return None
        sldLayoutIdLst = master._element.sldLayoutIdLst  # pyright: ignore[reportPrivateUsage]
        if sldLayoutIdLst is None:
            return None
        master_part = master.part
        for entry in sldLayoutIdLst.sldLayoutId_lst:
            resolved = master_part.related_slide_layout(entry.rId)
            if resolved._element is self._element:
                return entry.id
        return None

    @property
    def slide_layout_type(self) -> str:
        """Layout-kind token recorded in ``p:sldLayout/@type``.

        Returns a string drawn from ``ST_SlideLayoutType`` — ``"title"``,
        ``"obj"``, ``"twoObj"``, ``"titleOnly"``, ``"blank"``, ``"secHead"``,
        ``"picTx"``, ``"cust"``, and so on (see ECMA-376 Part 1 §19.7.15 for
        the full vocabulary). The attribute is optional on ``p:sldLayout`` and
        defaults to ``"cust"``; this property returns that default when the
        attribute is absent, which is the common case for Google Slides exports
        (see issue #864).

        Useful for locating a layout by kind when the layout names are empty
        or custom (e.g. when round-tripping a deck that originated in Google
        Slides). See also :meth:`SlideLayouts.get_by_type`.
        """
        return self._element.type

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide-layout inherits properties."""
        return self.part.slide_master

    @property
    def used_by_slides(self) -> tuple[Slide, ...]:
        """Tuple of slide objects based on this slide layout."""
        # ---getting Slides collection requires going around the horn a bit---
        slides = self.part.package.presentation_part.presentation.slides
        return tuple(s for s in slides if s.slide_layout == self)

    def _positional_name(self) -> str:
        """Return a positional name like ``"Layout 3"`` for this slide layout.

        Used as a fallback when ``p:cSld/@name`` is empty, which is common for
        Google-Slides-authored layouts (see issue #864). The index is 1-based
        and matches this layout's zero-based position within its parent
        master's ``slide_layouts`` collection. Falls back to ``"Layout"`` (no
        index) when the layout has been detached from its master or the master
        context cannot be located — this keeps a ``.name`` read on a
        standalone |SlideLayout| built from a raw XML element from crashing.
        """
        part = self.part
        if part is None:
            return "Layout"
        try:
            master = part.slide_master
        except (AttributeError, KeyError):  # pragma: no cover - defensive
            return "Layout"
        for idx, layout in enumerate(master.slide_layouts):
            if layout is self or layout._element is self._element:
                return "Layout %d" % (idx + 1)
        return "Layout"


class SlideLayouts(ParentedElementProxy):
    """Sequence of slide layouts belonging to a slide-master.

    Supports indexed access, len(), iteration, index() and remove().
    """

    part: SlideMasterPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldLayoutIdLst: CT_SlideLayoutIdList, parent: SlideMaster):
        super(SlideLayouts, self).__init__(sldLayoutIdLst, parent)
        self._sldLayoutIdLst = sldLayoutIdLst

    def __getitem__(self, idx: int) -> SlideLayout:
        """Provides indexed access, e.g. `slide_layouts[2]`."""
        try:
            sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[idx]
        except IndexError:
            raise IndexError("slide layout index out of range")
        return self.part.related_slide_layout(sldLayoutId.rId)

    def __iter__(self) -> Iterator[SlideLayout]:
        """Generate each |SlideLayout| in the collection, in sequence."""
        for sldLayoutId in self._sldLayoutIdLst.sldLayoutId_lst:
            yield self.part.related_slide_layout(sldLayoutId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldLayoutIdLst)

    def get_by_id(self, layout_id: int, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return |SlideLayout| having presentation-stable id `layout_id`.

        `layout_id` is matched against ``p:sldLayoutId/@id`` on each entry in
        the parent master's ``p:sldLayoutIdLst``. This id is stable across
        layout reordering and is the preferred handle for referencing a
        specific layout by identity rather than position (which changes when
        layouts are added, removed, or moved). Returns `default` (``None`` by
        default) when no entry has a matching id.

        See also :attr:`SlideLayout.slide_layout_id` and
        :meth:`SlideMaster.get_layout`.

        .. versionadded:: 2026.05.0
        """
        for entry in self._sldLayoutIdLst.sldLayoutId_lst:
            if entry.id == layout_id:
                return self.part.related_slide_layout(entry.rId)
        return default

    def get_by_name(self, name: str, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return SlideLayout object having `name`, or `default` if not found."""
        for slide_layout in self:
            if slide_layout.name == name:
                return slide_layout
        return default

    def get_by_type(
        self, layout_type: str, default: SlideLayout | None = None
    ) -> SlideLayout | None:
        """Return the first SlideLayout whose ``@type`` equals ``layout_type``.

        ``layout_type`` is a token from ``ST_SlideLayoutType`` such as
        ``"title"``, ``"obj"``, ``"twoObj"``, ``"titleOnly"``, ``"blank"``,
        ``"secHead"``, ``"picTx"``, or ``"cust"``. See ECMA-376 Part 1 §19.7.15
        for the full vocabulary.

        Intended as the layout-lookup path of choice when ``get_by_name()`` is
        unreliable because the source deck was authored in Google Slides (issue
        #864): Google Slides exports leave ``p:cSld/@name`` empty or assign
        non-standard names, but they do preserve the ``@type`` tokens on each
        layout it understands.

        Returns ``default`` (``None`` by default) when no layout matches.
        """
        for slide_layout in self:
            if slide_layout.slide_layout_type == layout_type:
                return slide_layout
        return default

    def index(self, slide_layout: SlideLayout) -> int:
        """Return zero-based index of `slide_layout` in this collection.

        Raises `ValueError` if `slide_layout` is not present in this collection.
        """
        for idx, this_layout in enumerate(self):
            if slide_layout == this_layout:
                return idx
        raise ValueError("layout not in this SlideLayouts collection")

    def remove(self, slide_layout: SlideLayout) -> None:
        """Remove `slide_layout` from the collection.

        Raises ValueError when `slide_layout` is in use; a slide layout which is the basis for one
        or more slides cannot be removed.
        """
        # ---raise if layout is in use---
        if slide_layout.used_by_slides:
            raise ValueError("cannot remove slide-layout in use by one or more slides")

        # ---target layout is identified by its index in this collection---
        target_idx = self.index(slide_layout)

        # --remove layout from p:sldLayoutIds of its master
        # --this stops layout from showing up, but doesn't remove it from package
        target_sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[target_idx]
        self._sldLayoutIdLst.remove(target_sldLayoutId)

        # --drop relationship from master to layout
        # --this removes layout from package, along with everything (only) it refers to,
        # --including images (not used elsewhere) and hyperlinks
        slide_layout.slide_master.part.drop_rel(target_sldLayoutId.rId)


class SlideMaster(_BaseMaster):
    """Slide master object.

    Provides access to slide layouts. Access to placeholders, regular shapes, and slide master-level
    properties is inherited from |_BaseMaster|.
    """

    _element: CT_SlideMaster  # pyright: ignore[reportIncompatibleVariableOverride]

    part: SlideMasterPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def effective_background(self) -> _EffectiveBackground | None:
        """Resolved :class:`_EffectiveBackground` of this master, or |None| when absent.

        The master is the top of the background inheritance chain, so this
        accessor simply wraps the master's own ``p:bg`` (if any) in a
        read-only |_EffectiveBackground| view. Returns |None| when the
        master carries no explicit background.

        Included for completeness with :attr:`Slide.effective_background`
        and :attr:`SlideLayout.effective_background` so callers that walk
        slide-like objects polymorphically can read the rendered
        background without reaching into the subclass hierarchy.

        .. versionadded:: 2026.05.0
        """
        if self._element.cSld.bg is None:
            return None
        return _EffectiveBackground("master", self)

    def add_layout(self, name: str, based_on: SlideLayout | None = None) -> SlideLayout:
        """Return a new |SlideLayout| created on this master with `name`.

        This is the "add a fresh layout to an existing master" path of
        issue #413 — the same-master cousin of :meth:`add_layout_from`
        (#1028, for *cross*-master layout import). It mirrors the *Insert
        Layout* action that PowerPoint's *Slide Master* view exposes:
        a new slide layout is created on this master and appended to its
        :attr:`slide_layouts` collection.

        `name` is the string written to the new layout's ``p:cSld/@name``
        and is the identifier subsequent :meth:`SlideLayouts.get_by_name`
        lookups resolve against. It must be unique within this master's
        layouts — supplying a name that collides with an existing layout
        raises :class:`ValueError`.

        `based_on` optionally selects an existing layout on *this* master
        to use as the starting point. When provided, the new layout's
        shape tree and header/footer settings are deep-cloned from
        `based_on` and then its ``p:cSld/@name`` is overwritten with
        `name`. When |None| (the default), the new layout is built from a
        minimal blank template: no placeholders, a ``@type="cust"``
        marker, and a ``p:clrMapOvr/a:masterClrMapping`` element so the
        color map inherits from this master. In both forms the new
        layout is related to this master via a fresh ``SLIDE_MASTER``
        relationship and a freshly allocated ``p:sldLayoutId`` entry
        is appended to this master's ``p:sldLayoutIdLst``.

        Raises :class:`ValueError` when `based_on` belongs to a different
        master (use :meth:`add_layout_from` for cross-master import), or
        when `name` is already in use on this master.

        .. versionadded:: 2026.05.0
        """
        if not name:
            raise ValueError("name must be a non-empty string")

        # -- guard against a name collision on this master --
        for existing in self.slide_layouts:
            if existing.name == name:
                raise ValueError(
                    "this master already has a layout named %r; pick a "
                    "different name or rename the existing layout first" % name
                )

        if based_on is not None and based_on.slide_master is not self:
            raise ValueError(
                "based_on layout must belong to this master; use "
                "SlideMaster.add_layout_from() to import a layout from "
                "a different master"
            )

        rId, new_layout = self.part.add_layout(name, based_on)
        sldLayoutIdLst = self._element.get_or_add_sldLayoutIdLst()
        sldLayoutIdLst.add_sldLayoutId(rId)
        return new_layout

    def add_layout_from(self, source_layout: SlideLayout) -> SlideLayout:
        """Return a new |SlideLayout| on this master cloned from `source_layout`.

        `source_layout` may come from any |Presentation| instance,
        including the one this master belongs to. Provides the
        cross-master layout-import use-case of issue #1028: the layout
        is needed on a deck whose master does not contain it (or on a
        different master within the same deck).

        A new slide-layout part is materialised in this presentation's
        package and related to this master via a fresh
        ``SLIDE_LAYOUT`` relationship; a new ``p:sldLayoutId`` entry is
        appended to this master's ``p:sldLayoutIdLst`` with a
        freshly-allocated layout id. Images and external hyperlinks
        referenced by the source layout are cloned into this package
        (images are content-deduplicated against existing parts).

        The cloned layout inherits its *theme* (color, font, and format
        scheme) from *this* master, not from `source_layout`'s original
        master. The layout's shape tree, placeholder geometry, and
        text content are preserved verbatim; colors, fonts, and any
        other theme-resolved formatting are re-rendered by this
        master's theme. This is the same inheritance model PowerPoint
        applies when a user drags a layout between masters.

        Raises |ValueError| when this master already contains a layout
        with the same name, to prevent the user from accidentally
        creating two same-named layouts that can no longer be
        distinguished via :meth:`SlideLayouts.get_by_name`.

        .. versionadded:: 2026.05.0
        """
        # -- Guard against a name collision on this master --------------
        source_name = source_layout.name
        if source_name:
            for existing in self.slide_layouts:
                if existing.name == source_name:
                    raise ValueError(
                        "destination master already has a layout named %r; "
                        "rename the source or destination layout before "
                        "calling add_layout_from()" % source_name
                    )

        rId, new_layout = self.part.add_layout_from(source_layout)
        sldLayoutIdLst = self._element.get_or_add_sldLayoutIdLst()
        sldLayoutIdLst.add_sldLayoutId(rId)
        return new_layout

    @property
    def name(self) -> str:
        """String representing the internal name of this slide master.

        PowerPoint rarely writes a ``p:cSld/@name`` on a slide master, so the underlying
        attribute is typically empty. To provide a usable identifier (see issue #679), this
        property falls back to a positional name of the form ``"Master N"`` (1-based index in
        ``presentation.slide_masters``) when no explicit name is set on the master. Assigning
        a non-empty string persists it to ``p:cSld/@name`` and subsequent reads return that
        explicit value verbatim. Assigning an empty string or |None| clears the attribute and
        restores the positional fallback.
        """
        explicit_name = self._element.cSld.name
        if explicit_name:
            return explicit_name
        return self._positional_name()

    @name.setter
    def name(self, value: str | None):
        new_value = "" if value is None else value
        self._element.cSld.name = new_value

    @lazyproperty
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| object providing access to this slide-master's layouts."""
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)

    def get_layout(self, layout_id: int, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return |SlideLayout| having presentation-stable id `layout_id`.

        `layout_id` is matched against ``p:sldLayoutId/@id`` on each entry in
        this master's ``p:sldLayoutIdLst``. The id is stable across layout
        reordering, so this lookup is the robust alternative to
        ``slide_master.slide_layouts[index]`` when a caller needs to reference
        a specific layout that may have been moved within the master.
        Returns `default` (``None`` by default) when no entry has a matching
        id. Equivalent to ``slide_master.slide_layouts.get_by_id(layout_id, default)``.

        .. versionadded:: 2026.05.0
        """
        return self.slide_layouts.get_by_id(layout_id, default)

    @lazyproperty
    def theme_colors(self) -> dict[str, RGBColor]:
        """Mapping of scheme-color name (e.g. ``"accent1"``) to |RGBColor|.

        Keys are the scheme-color names from the theme's `a:clrScheme`
        (``"dk1"``, ``"lt1"``, ``"dk2"``, ``"lt2"``, ``"accent1"``..``"accent6"``,
        ``"hlink"``, ``"folHlink"``) plus the color-map aliases from this
        master's ``p:clrMap`` (``"bg1"``, ``"bg2"``, ``"tx1"``, ``"tx2"``).

        Useful for calling :meth:`ColorFormat.to_rgb` on a scheme color.

        .. versionadded:: 2026.05.0
        """
        return _resolve_theme_colors(self.part)

    def _positional_name(self) -> str:
        """Return a positional name like ``"Master 3"`` for this master.

        Used as a fallback when ``p:cSld/@name`` is empty, which is the common case for
        PowerPoint-authored masters (see issue #679). The index is 1-based and matches this
        master's zero-based position in ``presentation.slide_masters``. Falls back to
        ``"Master"`` (no index) when the master has been detached from its presentation or
        the presentation context cannot be located — this avoids crashing a ``.name`` read
        on a standalone |SlideMaster| built from an XML element.
        """
        part = self.part
        if part is None:
            return "Master"
        try:
            presentation_part = part.package.presentation_part
        except AttributeError:  # pragma: no cover - defensive against exotic packages
            return "Master"
        try:
            slide_masters = presentation_part.presentation.slide_masters
        except AttributeError:  # pragma: no cover - defensive
            return "Master"
        for idx, master in enumerate(slide_masters):
            if master is self or master._element is self._element:
                return "Master %d" % (idx + 1)
        return "Master"


class SlideMasters(ParentedElementProxy):
    """Sequence of |SlideMaster| objects belonging to a presentation.

    Has list access semantics, supporting indexed access, len(), and iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldMasterIdLst: CT_SlideMasterIdList, parent: Presentation):
        super(SlideMasters, self).__init__(sldMasterIdLst, parent)
        self._sldMasterIdLst = sldMasterIdLst

    def __getitem__(self, idx: int) -> SlideMaster:
        """Provides indexed access, e.g. `slide_masters[2]`."""
        try:
            sldMasterId = self._sldMasterIdLst.sldMasterId_lst[idx]
        except IndexError:
            raise IndexError("slide master index out of range")
        return self.part.related_slide_master(sldMasterId.rId)

    def __iter__(self):
        """Generate each |SlideMaster| instance in the collection, in sequence."""
        for smi in self._sldMasterIdLst.sldMasterId_lst:
            yield self.part.related_slide_master(smi.rId)

    def __len__(self):
        """Support len() built-in function, e.g. `len(slide_masters) == 4`."""
        return len(self._sldMasterIdLst)


class _HeaderFooter(ElementProxy):
    """Provides access to header/footer/slide-number/date placeholder visibility.

    This object is returned by :attr:`~.SlideMaster.header_footer` on |SlideMaster|,
    |SlideLayout| and |NotesMaster|. It reads and writes the `p:hf` child element of its parent.
    Each toggle is an independent boolean that defaults to `True` when no explicit value is
    set (matching the `p:hf` XSD defaults).

    Setting a property to ``False`` hides that placeholder kind; setting back to ``True``
    restores inherited (visible) behavior. When every attribute is ``True``, no underlying
    `p:hf` element is kept in the XML — absence is identical to "all visible".
    """

    def __init__(self, element: CT_SlideMaster | CT_SlideLayout | CT_NotesMaster):
        super(_HeaderFooter, self).__init__(element)
        self._slide_element = element

    @property
    def slide_number_visible(self) -> bool:
        """`True` when the slide-number placeholder is shown on this master or layout.

        Defaults to `True` (the `p:hf/@sldNum` XSD default) when no `p:hf` is present.
        """
        hf = self._slide_element.hf
        return True if hf is None else hf.sldNum

    @slide_number_visible.setter
    def slide_number_visible(self, value: bool):
        self._apply_flag("sldNum", value)

    @property
    def header_visible(self) -> bool:
        """`True` when the header placeholder is shown on this master or layout."""
        hf = self._slide_element.hf
        return True if hf is None else hf.hdr

    @header_visible.setter
    def header_visible(self, value: bool):
        self._apply_flag("hdr", value)

    @property
    def footer_visible(self) -> bool:
        """`True` when the footer placeholder is shown on this master or layout."""
        hf = self._slide_element.hf
        return True if hf is None else hf.ftr

    @footer_visible.setter
    def footer_visible(self, value: bool):
        self._apply_flag("ftr", value)

    @property
    def date_visible(self) -> bool:
        """`True` when the date placeholder is shown on this master or layout."""
        hf = self._slide_element.hf
        return True if hf is None else hf.dt

    @date_visible.setter
    def date_visible(self, value: bool):
        self._apply_flag("dt", value)

    def _apply_flag(self, attr_name: str, value: bool):
        """Set `attr_name` on the `p:hf` element; remove `p:hf` when every flag is `True`.

        Adds `p:hf` when needed. The removal when every flag is back to True keeps the XML
        minimal and round-trips cleanly to an element-absent state.
        """
        if not isinstance(value, bool):
            raise TypeError(f"header_footer flag must be a bool, got {type(value).__name__}")
        hf = self._slide_element.get_or_add_hf()
        setattr(hf, attr_name, value)
        if hf.sldNum and hf.hdr and hf.ftr and hf.dt:
            self._slide_element.remove(hf)


class Transition(ElementProxy):
    """Proxy for the `p:transition` child of a `p:sld`.

    This object is returned by :attr:`Slide.transition` whether or not
    the slide has an explicit transition. Access is therefore always
    safe; the proxy adds a ``p:transition`` child element lazily the
    first time a setter is invoked (or :meth:`_get_or_add_transition`
    is called internally by the proxy).

    Public surface:

    * :attr:`type` — :class:`.PP_TRANSITION_TYPE` enum member.
      Reading returns :attr:`PP_TRANSITION_TYPE.NONE` when there is
      no transition variant present.
    * :attr:`duration` — int milliseconds or `None`. Read/write of
      the ``p14:dur`` extension attribute. Use :attr:`speed` for
      the coarse ``@spd`` selector.
    * :attr:`speed` — :class:`.PP_TRANSITION_SPEED` (``slow`` /
      ``med`` / ``fast``). Read/write of ``p:transition/@spd``.
    * :attr:`advance_on_click` — bool, ``p:transition/@advClick``.
    * :attr:`advance_after_time` — int ms or `None`, ``p:transition/@advTm``.
    * :attr:`wipe_direction` — :class:`.PP_TRANSITION_SIDE_DIRECTION`
      or ``None``. Reads/writes the ``@dir`` attribute on ``p:wipe``.
      ``None`` is returned when the current variant is not ``p:wipe``.
    * :attr:`morph_option` — the matching granularity for a MORPH
      transition (one of ``"byObject"``, ``"byWord"``, ``"byChar"``;
      default ``"byObject"``). Only meaningful when :attr:`type` is
      :attr:`PP_TRANSITION_TYPE.MORPH`.

    MORPH and ``mc:AlternateContent`` wrapping. Assigning
    :attr:`PP_TRANSITION_TYPE.MORPH` wraps the ``p:transition`` in an
    ``mc:AlternateContent`` that carries a ``p:fade`` fallback — this
    is how PowerPoint emits a MORPH transition so that viewers
    without the Office 2010 ``p14`` extension render a graceful
    fade instead of dropping the transition silently. Assigning any
    other transition type (or ``NONE``) unwraps it back to a plain
    ``p:transition`` child. See issue #942 and
    ``docs/dev/analysis/f8-animations-transitions.rst``.

    Out-of-scope (see the F8 analysis doc):

    * Other per-variant attributes such as ``@dir`` on ``p:cover`` /
      ``p:strips`` / ``p:zoom``, ``@orient`` on ``p:split`` /
      ``p:blinds``, ``@thruBlk`` on ``p:fade`` / ``p:cut``, and
      ``@spokes`` on ``p:wheel`` — these follow the same pattern as
      :attr:`wipe_direction` and can be added incrementally.
    * Sound-action ``p:sndAc`` (not tracked as a separate item).

    .. versionadded:: 2026.05.0
    """

    def __init__(self, sld: CT_Slide):
        super(Transition, self).__init__(sld)
        self._sld = sld

    # -- type --------------------------------------------------------

    @property
    def type(self) -> PP_TRANSITION_TYPE:
        """The :class:`.PP_TRANSITION_TYPE` selected on this slide.

        Returns :attr:`PP_TRANSITION_TYPE.NONE` when the slide has no
        ``p:transition`` element (direct or mc-wrapped), or when the
        ``p:transition`` has no variant child (it is valid for
        ``p:transition`` to carry only attributes such as ``@advTm``).
        """
        transition = self._sld.transition_effective
        if transition is None:
            return PP_TRANSITION_TYPE.NONE
        variant_tag = transition.variant_tag
        if variant_tag is None:
            return PP_TRANSITION_TYPE.NONE
        # -- extract the local name (e.g. "fade" from "p:fade" or "morph" --
        # -- from "p14:morph") and look up the enum member.
        local_name = variant_tag.split(":", 1)[1]
        try:
            return PP_TRANSITION_TYPE.from_xml(local_name)
        except ValueError:  # pragma: no cover - defensive, every variant is enumerated
            return PP_TRANSITION_TYPE.NONE

    @type.setter
    def type(self, value: PP_TRANSITION_TYPE | None) -> None:
        if value is None or value == PP_TRANSITION_TYPE.NONE:
            transition = self._sld.transition_effective
            if transition is None:
                return
            # -- MORPH was wrapped in mc:AlternateContent; drop the whole
            # -- wrapper on clear-to-NONE.
            if self._sld.transition_is_alt_content_wrapped:
                self._sld.remove_transition_effective()
                return
            transition._remove_variant()
            # -- if the transition element is now empty of attributes and
            # -- children, drop it for minimal XML round-trip.
            if not transition.attrib and len(transition) == 0:
                self._sld.remove(transition)
            return
        PP_TRANSITION_TYPE.validate(value)
        if value is PP_TRANSITION_TYPE.MORPH:
            # -- ensure the transition is mc:AlternateContent-wrapped and
            # -- then switch its variant to p14:morph --
            transition = self._sld.get_or_add_transition_effective()
            if not self._sld.transition_is_alt_content_wrapped:
                transition = self._sld.wrap_transition_in_alt_content()
            transition.set_variant("p14:morph")
        else:
            # -- unwrap from any existing mc:AlternateContent before
            # -- switching to a plain p: variant (the wrapper is only
            # -- needed for the p14:morph child).
            if self._sld.transition_is_alt_content_wrapped:
                self._sld.unwrap_transition_from_alt_content()
            transition = self._sld.get_or_add_transition()
            transition.set_variant("p:%s" % value.xml_value)

    # -- morph_option (MORPH-only) -----------------------------------

    _MORPH_OPTIONS = ("byObject", "byWord", "byChar")

    @property
    def morph_option(self) -> str | None:
        """Matching-granularity token for a MORPH transition.

        Returns one of ``"byObject"``, ``"byWord"``, or ``"byChar"``.
        Returns ``None`` when :attr:`type` is not
        :attr:`PP_TRANSITION_TYPE.MORPH`.

        The default (schema-level) when a MORPH transition is authored
        without an explicit ``@option`` attribute is ``"byObject"``.
        """
        transition = self._sld.transition_effective
        if transition is None:
            return None
        if transition.variant_tag != "p14:morph":
            return None
        morph = transition.find(qn("p14:morph"))
        if morph is None:
            return None
        # -- CT_TransitionMorph.option defaults to "byObject" when absent --
        return morph.option

    @morph_option.setter
    def morph_option(self, value: str) -> None:
        if value not in self._MORPH_OPTIONS:
            raise ValueError(
                "morph_option must be one of %r, got %r" % (self._MORPH_OPTIONS, value)
            )
        transition = self._sld.transition_effective
        if transition is None or transition.variant_tag != "p14:morph":
            raise ValueError(
                "morph_option only applies to a MORPH transition; "
                "set transition.type = PP_TRANSITION_TYPE.MORPH first"
            )
        morph = transition.find(qn("p14:morph"))
        assert morph is not None  # invariant: variant_tag check above
        morph.option = value

    # -- duration (ms) -----------------------------------------------

    @property
    def duration(self) -> int | None:
        """Transition duration in milliseconds, or ``None`` when unspecified.

        Reads the ``p14:dur`` attribute (Office 2010 extension). When
        absent the property returns ``None`` — PowerPoint then falls
        back to a default derived from ``@spd`` (slow=1600ms,
        med=1000ms, fast=500ms).
        """
        transition = self._sld.transition_effective
        if transition is None:
            return None
        return transition.dur

    @duration.setter
    def duration(self, value: int | None) -> None:
        if value is None:
            transition = self._sld.transition_effective
            if transition is None:
                return
            transition.dur = None
            if (
                not transition.attrib
                and len(transition) == 0
                and not self._sld.transition_is_alt_content_wrapped
            ):
                self._sld.remove(transition)
            return
        transition = self._sld.get_or_add_transition_effective()
        transition.dur = value

    # -- speed (slow / med / fast) -----------------------------------

    @property
    def speed(self) -> PP_TRANSITION_SPEED:
        """Coarse slow / med / fast selector for this transition.

        Reflects ``p:transition/@spd``. Defaults to
        :attr:`PP_TRANSITION_SPEED.FAST` (the schema default) when no
        ``p:transition`` element is present or ``@spd`` is absent.

        This is the ISO-29500-defined selector; for millisecond
        precision use :attr:`duration` (which writes the 2010-extension
        ``p14:dur`` attribute).
        """
        transition = self._sld.transition
        if transition is None:
            return PP_TRANSITION_SPEED.FAST
        return PP_TRANSITION_SPEED.from_xml(transition.spd)

    @speed.setter
    def speed(self, value: PP_TRANSITION_SPEED) -> None:
        PP_TRANSITION_SPEED.validate(value)
        transition = self._sld.get_or_add_transition()
        transition.spd = PP_TRANSITION_SPEED.to_xml(value)

    # -- advance_on_click --------------------------------------------

    @property
    def advance_on_click(self) -> bool:
        """Whether a mouse click advances past the transition.

        Reflects ``p:transition/@advClick``. Defaults to ``True``
        (schema default) when no ``p:transition`` element exists or
        the attribute is absent.
        """
        transition = self._sld.transition_effective
        if transition is None:
            return True
        return transition.advClick

    @advance_on_click.setter
    def advance_on_click(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise TypeError("advance_on_click must be a bool, got %s" % type(value).__name__)
        transition = self._sld.get_or_add_transition_effective()
        transition.advClick = value

    # -- advance_after_time ------------------------------------------

    @property
    def advance_after_time(self) -> int | None:
        """Auto-advance delay in ms, or ``None`` when not auto-advancing.

        Reflects ``p:transition/@advTm``. The attribute is absent when
        the slide only advances on click; setting the property to
        ``None`` removes the attribute, restoring click-only advance.
        """
        transition = self._sld.transition_effective
        if transition is None:
            return None
        return transition.advTm

    @advance_after_time.setter
    def advance_after_time(self, value: int | None) -> None:
        if value is None:
            transition = self._sld.transition_effective
            if transition is None:
                return
            if "advTm" in transition.attrib:
                del transition.attrib["advTm"]
            if (
                not transition.attrib
                and len(transition) == 0
                and not self._sld.transition_is_alt_content_wrapped
            ):
                self._sld.remove(transition)
            return
        if not isinstance(value, int) or value < 0:
            raise ValueError(
                "advance_after_time must be a non-negative int (milliseconds), got %r" % (value,)
            )
        transition = self._sld.get_or_add_transition_effective()
        transition.advTm = value

    # -- advance_after_seconds (seconds convenience on top of advance_after_time) --

    @property
    def advance_after_seconds(self) -> float | None:
        """Auto-advance delay expressed in *seconds*, or ``None`` when unset.

        Thin convenience view over :attr:`advance_after_time`: reads the
        ``p:transition/@advTm`` (milliseconds) and returns the value as
        ``float`` seconds; returns ``None`` when the attribute is absent.

        Setting to a non-negative number writes ``@advTm`` in milliseconds
        (``int(round(value * 1000))``); setting to ``None`` clears
        ``@advTm``. Use :attr:`advance_after_time` directly when you need
        millisecond precision or want to avoid the float round-trip.

        .. versionadded:: 2026.05.0
        """
        ms = self.advance_after_time
        if ms is None:
            return None
        return ms / 1000.0

    @advance_after_seconds.setter
    def advance_after_seconds(self, value: float | int | None) -> None:
        if value is None:
            self.advance_after_time = None
            return
        if not isinstance(value, (int, float)) or value < 0:
            raise ValueError(
                "advance_after_seconds must be a non-negative number (seconds), "
                "got %r" % (value,)
            )
        self.advance_after_time = int(round(float(value) * 1000))

    # -- wipe_direction (per-variant flag on p:wipe) -----------------

    @property
    def wipe_direction(self) -> PP_TRANSITION_SIDE_DIRECTION | None:
        """Direction of a wipe transition.

        Returns the :class:`.PP_TRANSITION_SIDE_DIRECTION` member
        corresponding to ``p:transition/p:wipe/@dir`` (schema default
        ``"l"``) when the current variant is ``p:wipe``. Returns
        ``None`` when the transition is a different variant (or
        absent) — ``@dir`` has no meaning outside the directional
        variants.
        """
        transition = self._sld.transition
        if transition is None:
            return None
        wipe = transition.find(qn("p:wipe"))
        if wipe is None:
            return None
        return PP_TRANSITION_SIDE_DIRECTION.from_xml(wipe.get("dir", "l"))

    @wipe_direction.setter
    def wipe_direction(self, value: PP_TRANSITION_SIDE_DIRECTION) -> None:
        PP_TRANSITION_SIDE_DIRECTION.validate(value)
        transition = self._sld.get_or_add_transition()
        wipe = transition.find(qn("p:wipe"))
        if wipe is None:
            # -- ensure the variant is p:wipe; replace any current variant --
            wipe = transition.set_variant("p:wipe")
            if wipe is None:  # pragma: no cover - set_variant returns element
                return
        wipe.set("dir", PP_TRANSITION_SIDE_DIRECTION.to_xml(value))


class AnimationEffectView(ElementProxy):
    """Read-only proxy for a single entrance / exit / emphasis effect.

    An *effect* is one step in the slide's main animation sequence —
    the unit PowerPoint advances through on each click (or time
    trigger) when running the slide. Each effect is represented in the
    XML by an effect-level ``p:par`` whose ``p:cTn`` child carries a
    ``@presetClass`` attribute. See
    :func:`pptx.oxml.timing.iter_main_sequence_effects` for how the
    main sequence is walked.

    The MVP surfaces five read-only accessors — ``shape_id``,
    ``preset_class``, ``preset_id``, ``preset_subtype``, and ``delay``.
    They describe the effect well enough to round-trip a
    PowerPoint-authored timing and to power introspection-based
    regression tests (see issue #256). Authoring helpers —
    ``AnimationEffectView.set_preset()``, ``AnimationSequence.add(...)``,
    reorder / delete — are deferred to the animation-authoring items
    (#102, #264, #1106) that stack on this foundation.

    Instances are not constructed by client code; obtain them from
    :attr:`.Slide.animation_sequence`.

    .. note::
       Renamed from ``AnimationEffect`` (which collided with the
       authoring class :class:`pptx.animation.AnimationEffect`).
       ``AnimationEffect`` remains available on ``pptx.slide`` as a
       deprecated alias for one release; new code should use
       ``AnimationEffectView``.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, par):
        super(AnimationEffectView, self).__init__(par)
        self._par = par

    @property
    def shape_id(self) -> int | None:
        """ID of the shape this effect targets, or ``None`` when absent.

        Reads ``@spid`` from the first ``p:spTgt`` descendant of the
        effect's ``p:par``. Matches :attr:`.BaseShape.shape_id` on the
        corresponding shape in this slide. Some effects (e.g. timed
        color-scheme changes, or effects targeting a graphic-frame
        sub-element via ``p:subSpTgt``) do not carry a ``p:spTgt`` at
        all; ``None`` is returned for those.

        .. versionadded:: 2026.05.0
        """
        from pptx.oxml.timing import first_spTgt_spid

        return first_spTgt_spid(self._par)

    @property
    def preset_class(self) -> str | None:
        """Preset-class string — ``"entr"`` / ``"exit"`` / ``"emph"`` / ....

        Reads ``p:cTn/@presetClass`` of the effect-level ``p:par``.
        Canonical ST_TLTimeNodePresetClassType values from pml.xsd
        include ``entr`` (entrance), ``exit``, ``emph`` (emphasis),
        ``path`` (motion path), ``verb`` (OLE verb), and
        ``mediacall``. Returns ``None`` on malformed XML where the
        ``p:cTn`` is missing or the attribute is absent (both unusual
        since :func:`iter_main_sequence_effects` filters on its
        presence).

        .. versionadded:: 2026.05.0
        """
        cTn = self._par.find(qn("p:cTn"))
        if cTn is None:
            return None
        return cTn.presetClass

    @property
    def preset_id(self) -> int | None:
        """Preset id (int) — identifies the specific effect, or ``None``.

        Reads ``p:cTn/@presetID`` of the effect-level ``p:par``. This
        integer identifier disambiguates within a preset class (e.g.
        within ``entr`` the value ``1`` is ``appear``, ``2`` is
        ``fly-in``, ``10`` is ``fade``, etc.). Values correspond to
        the ``presetID`` column in Microsoft's ``[MS-PPTX]`` animation
        preset documentation.

        .. versionadded:: 2026.05.0
        """
        cTn = self._par.find(qn("p:cTn"))
        if cTn is None:
            return None
        return cTn.presetID

    @property
    def preset_subtype(self) -> int | None:
        """Preset subtype (int) — direction / variant selector, or ``None``.

        Reads ``p:cTn/@presetSubtype`` of the effect-level ``p:par``.
        The subtype combines with ``preset_id`` to pick, for example,
        the direction of a "fly-in" entrance (``8`` = from-left,
        ``4`` = from-top, etc.). A value of ``0`` typically means the
        default variant for the selected preset. ``None`` when the
        attribute is absent.

        .. versionadded:: 2026.05.0
        """
        cTn = self._par.find(qn("p:cTn"))
        if cTn is None:
            return None
        return cTn.presetSubtype

    @property
    def delay(self) -> int | str | None:
        """Start-delay before the effect fires, or ``None``.

        Walks the effect ``p:par``'s ``p:cTn/p:stCondLst/p:cond`` and
        returns the ``@delay`` value. The value is:

        * ``0`` — fires immediately on the click that advances to
          this step (``with previous`` semantics).
        * An ``int`` > 0 — delay in milliseconds after the trigger.
        * ``"indefinite"`` — fires on an explicit click-trigger
          (the PowerPoint default for ``on click`` effects).
        * ``None`` — no ``p:stCondLst/p:cond/@delay`` present.

        The MVP reads only the first ``p:cond``, which covers the
        overwhelming majority of PowerPoint-authored effects (full
        ``p:cond`` list semantics — multiple conditions, ``@evt``,
        trigger-by-shape — are downstream #264 / #861).

        .. versionadded:: 2026.05.0
        """
        cTn = self._par.find(qn("p:cTn"))
        if cTn is None:
            return None
        stCondLst = cTn.find(qn("p:stCondLst"))
        if stCondLst is None:
            return None
        cond = stCondLst.find(qn("p:cond"))
        if cond is None:
            return None
        delay_str = cond.get("delay")
        if delay_str is None:
            return None
        if delay_str == "indefinite":
            return "indefinite"
        try:
            return int(delay_str)
        except ValueError:  # pragma: no cover - defensive, schema forbids
            return None


# -- Deprecation alias --------------------------------------------------
# ``AnimationEffect`` (the read-only introspection proxy that pairs with
# :attr:`Slide.animation_sequence`, originally shipped by Wave 5 #256)
# collided in name with :class:`pptx.animation.AnimationEffect` (the
# authoring API shipped by Wave 5 #102). The introspection proxy has
# been renamed :class:`AnimationEffectView`; this alias preserves
# backwards compatibility for one release. New code should use
# :class:`AnimationEffectView` or import :class:`AnimationEffect` from
# :mod:`pptx.animation` for authoring.
AnimationEffect = AnimationEffectView


class ShapeAnimation(ElementProxy):
    """Read-only proxy for a single shape-targeted animation effect.

    Returned by :meth:`.Slide.iter_shape_animations`. Wraps one of the
    shape-animation behaviour elements (``p:anim``, ``p:animEffect``,
    ``p:animMotion``, ``p:animRot``, ``p:animScale``, ``p:animClr``,
    ``p:set``) together with the ``shape_id`` resolved from that
    effect's ``p:spTgt/@spid``.

    MVP read surface (issue #264):

    * :attr:`shape_id` — ``int``, the ``@spid`` of the targeted shape.
    * :attr:`effect_type` — ``str``, the local name of the behaviour
      element (``"anim"``, ``"animEffect"``, ``"set"``, …).
    * :attr:`delay_ms` — ``int`` / ``"indefinite"`` / ``None``, the
      value of the effect's first ``p:stCondLst/p:cond/@delay``.
    * :attr:`duration_ms` — ``int`` / ``"indefinite"`` / ``None``, the
      ``@dur`` attribute of the effect's ``p:cTn`` child.
    * :attr:`element` — the underlying ``lxml`` element so callers that
      need to hand-edit the XML today can still do so (the full-authoring
      API is deferred to Wave 7).

    Write support (adjusting delay, adding new entrance effects, etc.)
    is deferred — see ``docs/dev/analysis/f8-animations-transitions.rst``
    and the downstream-items matrix.

    .. versionadded:: 2026.05.0
    """

    _SHAPE_EFFECT_TAGS = {
        "anim",
        "animEffect",
        "animMotion",
        "animRot",
        "animScale",
        "animClr",
        "set",
    }

    def __init__(self, effect_elm: BaseOxmlElement, shape_id: int):
        super(ShapeAnimation, self).__init__(effect_elm)
        self._effect = effect_elm
        self._shape_id = shape_id

    @property
    def shape_id(self) -> int:
        """``int`` ``@spid`` of the shape this effect targets.

        .. versionadded:: 2026.05.0
        """
        return self._shape_id

    @property
    def effect_type(self) -> str:
        """Local name of the behaviour element (e.g. ``"anim"``, ``"set"``).

        .. versionadded:: 2026.05.0
        """
        tag = self._effect.tag
        # -- strip the namespace prefix — tag looks like
        # -- "{http://.../main}anim"
        return tag.rsplit("}", 1)[-1] if "}" in tag else tag

    @property
    def delay_ms(self) -> int | str | None:
        """First start-condition delay of this effect.

        Returns the value of the effect's
        ``p:cTn/p:stCondLst/p:cond/@delay`` attribute — an ``int``
        (milliseconds), the string ``"indefinite"``, or ``None`` when
        no start-condition delay is specified.

        This is the exact read-half of the feature requested in #264
        ("is there a way to modify the delay of shape animations?").
        A write accessor is tracked under #861 and layered onto this
        proxy when it lands.

        .. versionadded:: 2026.05.0
        """
        cTn = self._cTn
        if cTn is None:
            return None
        conds = _xpath_p(cTn, "p:stCondLst/p:cond")
        if not conds:
            return None
        delay = conds[0].get("delay")
        if delay is None:
            return None
        if delay == "indefinite":
            return "indefinite"
        return int(delay)

    @property
    def duration_ms(self) -> int | str | None:
        """``@dur`` of the effect's ``p:cTn`` child, or ``None`` when absent.

        ``int`` milliseconds, or the string ``"indefinite"``.

        .. versionadded:: 2026.05.0
        """
        cTn = self._cTn
        if cTn is None:
            return None
        # CT_TLCommonTimeNodeData.dur returns the already-converted
        # ST_TLTime value when ``p:cTn`` was parsed with the typed class,
        # but behaviour-scoped p:cTn elements may be plain
        # BaseOxmlElement in early implementations — fall back to the
        # raw attribute.
        dur = getattr(cTn, "dur", None)
        if dur is not None:
            return dur
        raw = cTn.get("dur")
        if raw is None:
            return None
        if raw == "indefinite":
            return "indefinite"
        return int(raw)

    @property
    def element(self) -> BaseOxmlElement:
        """The underlying behaviour element (``p:anim``, ``p:set``, …).

        Exposed so callers can hand-edit the XML today; the full typed
        authoring API is deferred.

        .. versionadded:: 2026.05.0
        """
        return self._effect

    # -- private helpers --------------------------------------------

    @property
    def _cTn(self):
        """Return the behaviour's ``p:cTn`` descendant, or ``None``.

        In every shape-animation behaviour type a ``p:cTn`` lives under
        a ``p:cBhvr`` wrapper (or directly under ``p:set`` via
        ``p:cBhvr``). We use a descendant XPath so we don't have to
        hard-code each wrapper shape.
        """
        cTns = _xpath_p(self._effect, ".//p:cTn")
        return cTns[0] if cTns else None


def _xpath_p(elm, expr: str):
    """Run an XPath expression on `elm` with the ``p`` prefix bound.

    Works uniformly whether `elm` is a :class:`BaseOxmlElement` (which
    carries the Open-XML namespace map built-in but overrides ``xpath``
    to reject the ``namespaces=`` kwarg) or a plain ``lxml._Element``
    (for which we must supply the namespaces explicitly).
    """
    from pptx.oxml.xmlchemy import BaseOxmlElement

    if isinstance(elm, BaseOxmlElement):
        return elm.xpath(expr)
    return elm.xpath(expr, namespaces=namespaces("p"))


def _ancestor_with_local_name(elm, local_names):
    """Walk up from `elm` returning the first ancestor whose local tag is in `local_names`.

    Returns ``None`` when no such ancestor exists. Used by
    :meth:`Slide.iter_shape_animations` to locate the behaviour element
    wrapping a ``p:spTgt`` match.
    """
    current = elm.getparent()
    while current is not None:
        tag = current.tag
        local = tag.rsplit("}", 1)[-1] if "}" in tag else tag
        if local in local_names:
            return current
        current = current.getparent()
    return None


_SHAPE_TAGS = frozenset(
    (
        qn("p:sp"),
        qn("p:pic"),
        qn("p:cxnSp"),
        qn("p:graphicFrame"),
        qn("p:grpSp"),
    )
)


def _is_shape_elm(elm) -> bool:
    """Return ``True`` when `elm` is one of the spTree shape elements.

    Recognises the five shape tag names directly rather than by oxml class
    hierarchy: ``CT_ShapeNonVisual`` and ``CT_GroupShapeNonVisual`` also
    inherit from :class:`BaseShapeElement` so an ``isinstance`` check would
    over-match helper elements like ``p:nvSpPr``.
    """
    return getattr(elm, "tag", None) in _SHAPE_TAGS


def _nearest_shape_ancestor(elm):
    """Walk up from `elm` returning the first ancestor that is a shape element.

    A *shape element* is one of ``p:sp`` / ``p:pic`` / ``p:cxnSp`` /
    ``p:graphicFrame`` / ``p:grpSp``. Returns ``None`` when no such
    ancestor exists (for example, when walking up from the ``p:spTree``
    element itself).

    Used by :meth:`Slide.find_shapes_by_xpath` to resolve an XPath match
    that landed on a child element (e.g. ``p:cNvPr`` or ``a:xfrm``) back
    to the owning shape so a shape proxy can be returned.
    """
    current = elm.getparent()
    while current is not None:
        if _is_shape_elm(current):
            return current
        current = current.getparent()
    return None


class SlideTags:
    """Dict-like proxy for a slide's VBA-style custom tag collection.

    Instances are obtained via :attr:`Slide.tags`. The proxy backs onto
    an optional |TagsPart| that is materialized lazily — reading or
    iterating a slide that has no tags returns empty results without
    creating any XML. The tags part (and the ``p:cSld/p:custDataLst/p:tags``
    reference inside the slide) is created the first time a tag is
    written through :meth:`__setitem__`.

    Tag names and values are ``str`` per the ECMA-376 ``CT_StringTag``
    schema. A tag name is unique within a slide; writing an existing
    name overwrites the stored value.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, slide_part: SlidePart):
        self._slide_part = slide_part

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        tag_list = self._tag_list_or_none
        if tag_list is None:
            return False
        return tag_list.get_by_name(name) is not None

    def __delitem__(self, name: str) -> None:
        if not isinstance(name, str):
            raise TypeError("tag name must be a str, got %s" % type(name).__name__)
        tag_list = self._tag_list_or_none
        if tag_list is None or not tag_list.remove_tag(name):
            raise KeyError(name)

    def __getitem__(self, name: str) -> str:
        if not isinstance(name, str):
            raise TypeError("tag name must be a str, got %s" % type(name).__name__)
        tag_list = self._tag_list_or_none
        if tag_list is not None:
            tag = tag_list.get_by_name(name)
            if tag is not None:
                return tag.val
        raise KeyError(name)

    def __iter__(self) -> Iterator[str]:
        tag_list = self._tag_list_or_none
        if tag_list is None:
            return
        for tag in tag_list.tag_lst:
            yield tag.name

    def __len__(self) -> int:
        tag_list = self._tag_list_or_none
        if tag_list is None:
            return 0
        return len(tag_list.tag_lst)

    def __setitem__(self, name: str, value: str) -> None:
        if not isinstance(name, str):
            raise TypeError("tag name must be a str, got %s" % type(name).__name__)
        if not isinstance(value, str):
            raise TypeError("tag value must be a str, got %s" % type(value).__name__)
        tags_part = self._slide_part.get_or_add_tags_part()
        tags_part.tag_list.set_tag(name, value)

    def get(self, name: str, default: str | None = None) -> str | None:
        """Return value of tag `name` if present, otherwise `default`.

        Mirrors :meth:`dict.get`. Never raises ``KeyError`` for a missing
        tag, and never materializes a tags part.

        .. versionadded:: 2026.05.0
        """
        if not isinstance(name, str):
            return default
        tag_list = self._tag_list_or_none
        if tag_list is None:
            return default
        tag = tag_list.get_by_name(name)
        if tag is None:
            return default
        return tag.val

    def keys(self) -> list[str]:
        """Return a list of tag names in document order.

        .. versionadded:: 2026.05.0
        """
        return list(iter(self))

    def values(self) -> list[str]:
        """Return a list of tag values in document order.

        .. versionadded:: 2026.05.0
        """
        tag_list = self._tag_list_or_none
        if tag_list is None:
            return []
        return [tag.val for tag in tag_list.tag_lst]

    def items(self) -> list[tuple[str, str]]:
        """Return a list of ``(name, value)`` pairs in document order.

        .. versionadded:: 2026.05.0
        """
        tag_list = self._tag_list_or_none
        if tag_list is None:
            return []
        return [(tag.name, tag.val) for tag in tag_list.tag_lst]

    @property
    def _tag_list_or_none(self) -> CT_TagList | None:
        """Return the backing ``p:tagLst`` element, or ``None`` when no tags part exists."""
        tags_part = self._slide_part.tags_part
        if tags_part is None:
            return None
        return tags_part.tag_list


class _FollowMasterBackground:
    """Dual-form return value of :attr:`Slide.follow_master_background`.

    Behaves as a boolean when interrogated (``True`` when the owning
    slide has no ``p:bg`` child and therefore inherits its background
    from the master / layout, ``False`` when the slide carries an
    explicit background). Behaves as a zero-argument method when
    *called* (``slide.follow_master_background()``) — calling it drops
    any ``p:bg`` child on the slide's ``p:cSld`` and returns the owning
    |Slide|, matching PowerPoint's *Reset Background* button.

    Instances are not constructed by client code; they are produced by
    :attr:`Slide.follow_master_background` on each access.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, slide: Slide):
        self._slide = slide

    def __bool__(self) -> bool:
        return self._slide._element.bg is None  # pyright: ignore[reportPrivateUsage]

    def __call__(self) -> Slide:
        """Remove any custom background on the slide, restoring master inheritance.

        Idempotent — a no-op on a slide that already follows the
        master background. Returns the owning |Slide| so the call can
        be chained.
        """
        # -- CT_CommonSlideData owns the optional `p:bg` child; _remove_bg
        # -- is the generated ZeroOrOne helper and is safe to call when
        # -- no p:bg is present.
        cSld = self._slide._element.cSld  # pyright: ignore[reportPrivateUsage]
        cSld._remove_bg()  # pyright: ignore[reportPrivateUsage]
        return self._slide

    def __eq__(self, other: object) -> bool:
        return bool(self) == other

    def __ne__(self, other: object) -> bool:
        return not self.__eq__(other)

    def __hash__(self) -> int:
        return hash(bool(self))

    def __repr__(self) -> str:
        return repr(bool(self))


class _Background(ElementProxy):
    """Provides access to slide background properties.

    Note that the presence of this object does not by itself imply an
    explicitly-defined background; a slide with an inherited background still
    has a |_Background| object.
    """

    def __init__(self, cSld: CT_CommonSlideData, parent: _BaseSlide | None = None):
        super(_Background, self).__init__(cSld)
        self._cSld = cSld
        self._parent = parent

    @property
    def bg_element(self) -> CT_Background | None:
        """The underlying ``p:bg`` element, or ``None`` when the background is inherited.

        Use this property for raw-XML operations such as copying the background
        markup from one slide to another. Unlike :attr:`fill`, reading this
        property does not materialize a ``p:bg`` subtree when none is present,
        so inheritance from the master or layout remains intact.

        .. versionadded:: 2026.05.0
        """
        return self._cSld.bg

    @property
    def part(self):
        """The package part containing this background (the slide/master/layout part)."""
        if self._parent is None:
            raise ValueError("background has no parent slide")
        return self._parent.part

    @lazyproperty
    def fill(self) -> FillFormat:
        """|FillFormat| instance for this background.

        This |FillFormat| object is used to interrogate or specify the fill
        of the slide background.

        Note that accessing this property is potentially destructive. A slide
        background can also be specified by a background style reference and
        accessing this property will remove that reference, if present, and
        replace it with NoFill. This is frequently the case for a slide
        master background.

        This is also the case when there is no explicitly defined background
        (background is inherited); merely accessing this property will cause
        the background to be set to NoFill and the inheritance link will be
        interrupted. This is frequently the case for a slide background.

        Of course, if you are accessing this property in order to set the
        fill, then these changes are of no consequence, but the existing
        background cannot be reliably interrogated using this property unless
        you have already established it is an explicit fill.

        If the background is already a fill, then accessing this property
        makes no changes to the current background.
        """
        bgPr = self._cSld.get_or_add_bgPr()
        part = self._parent if self._parent is not None else None
        return FillFormat.from_fill_parent(bgPr, part)


class _EffectiveBackground:
    """Read-only view of the slide background PowerPoint would actually render.

    Unlike :class:`_Background` — which is bound to a single slide / layout /
    master and is side-effect-prone (its :attr:`~_Background.fill` accessor
    materializes a ``p:bgPr/a:noFill`` subtree on first read) — an
    |_EffectiveBackground| resolves the inheritance chain
    *slide → layout → master* and exposes the first ancestor that carries an
    explicit ``p:bg``. Access is purely read-only: no accessor on this class
    mutates the underlying XML.

    Returned from :attr:`Slide.effective_background` (and
    :attr:`SlideLayout.effective_background` /
    :attr:`SlideMaster.effective_background`). Instances are not constructed
    by client code.

    .. versionadded:: 2026.05.0
    """

    def __init__(self, source: str, owner: _BaseSlide):
        self._source = source
        self._owner = owner

    @property
    def source(self) -> str:
        """``"slide"``, ``"layout"``, or ``"master"`` — which ancestor supplied the background.

        Lets callers tell whether the rendered background is a slide-level
        override (``"slide"``), inherited from the layout (``"layout"``), or
        inherited from the master (``"master"``). See
        :attr:`Slide.effective_background`.
        """
        return self._source

    @property
    def owner(self) -> _BaseSlide:
        """The slide-like object (|Slide| / |SlideLayout| / |SlideMaster|) that owns the `p:bg`.

        Useful when the caller wants to operate on the explicit background
        with the full :class:`_Background` surface — for example,
        ``effective.owner.background.fill`` walks back to the mutable
        proxy. Reading this property does not mutate anything.
        """
        return self._owner

    @property
    def bg_element(self) -> CT_Background:
        """The underlying ``p:bg`` element that supplies the rendered background.

        Guaranteed to be non-|None| — an |_EffectiveBackground| is only
        produced when a ``p:bg`` was resolved in the inheritance chain.
        """
        bg = self._owner._element.cSld.bg  # pyright: ignore[reportPrivateUsage]
        assert bg is not None  # guaranteed by the factory
        return bg

    @property
    def fill(self) -> FillFormat | None:
        """|FillFormat| reading the resolved ``p:bg/p:bgPr`` non-destructively.

        Returns |None| when the resolved ``p:bg`` carries a ``p:bgRef``
        (a theme-keyed background-style reference) instead of an explicit
        ``p:bgPr`` — ``p:bgRef`` is not an :class:`FillFormat` surface and
        has no directly-readable ``fore_color``. Callers that need to
        inspect a style reference can read :attr:`bg_element` directly.

        Unlike :attr:`_Background.fill`, reading this property is
        **side-effect free**: the underlying XML is not modified, so the
        inheritance link remains intact.
        """
        bg = self.bg_element
        bgPr = bg.bgPr
        if bgPr is None:
            return None
        return FillFormat.from_fill_parent(bgPr, self._owner)


def _clone_text_frame(src: TextFrame, tgt: TextFrame) -> None:
    """Replace `tgt`'s paragraph content with deep-copies of `src`'s paragraphs.

    Clears every paragraph on `tgt` and appends a fresh paragraph for each
    paragraph on `src`, calling :meth:`._Paragraph.clone_from` to deep-copy
    the source paragraph's ``a:pPr``, runs, and ``a:endParaRPr`` onto the
    new paragraph. Used by :meth:`.Slide.clone_shapes_from` to transfer a
    placeholder's text body without going through the placeholder element
    itself (which can't be cloned via :meth:`.BaseShape.clone_onto`).

    Works across text frames, slides, and presentations — each paragraph is
    deep-copied so the source and target remain fully independent after the
    call returns.
    """
    # -- clear `tgt` down to one empty paragraph so we can overwrite that
    # -- one with the first source paragraph, then append the rest.
    tgt.clear()
    src_paragraphs = src.paragraphs
    if not src_paragraphs:  # pragma: no cover -- a TextFrame always has >=1 `a:p`
        return
    # -- after `clear()`, tgt has exactly one `a:p`; reuse it for paragraph 0 --
    tgt.paragraphs[0].clone_from(src_paragraphs[0])
    for src_p in src_paragraphs[1:]:
        new_p = tgt.add_paragraph()
        new_p.clone_from(src_p)


def _clone_non_text_placeholder_onto(shape: BaseShape, shape_tree: SlideShapes) -> None:
    """Clone a chart / table / picture placeholder onto `shape_tree` as non-placeholder.

    Placeholders holding chart / table / picture content carry that content
    in structurally clone-safe XML (``<p:graphicFrame>`` or ``<p:pic>``).
    ``BaseShape.clone_onto`` refuses to clone placeholders because duplicating
    a placeholder ``idx`` would break the slide invariant — but the underlying
    content is perfectly cloneable. This helper strips the ``<p:ph>`` marker
    (the placeholder flag) from a deep-copy of the source shape, then routes
    the now-non-placeholder shape through ``clone_onto``.

    The resulting shape loses its layout inheritance (its size / position /
    title-from-layout properties become explicit) but the chart / table /
    image content is fully preserved and re-embedded on the target package.
    """
    import copy as _copy

    from pptx.oxml.ns import qn as _qn

    # -- Deep-copy the source shape's element so we can strip <p:ph> without
    # -- mutating the source. ``BaseShape`` stores the element at ._element
    # -- and the ``is_placeholder`` flag hinges on the presence of <p:ph>
    # -- inside <p:nvSpPr>/<p:nvPicPr>/<p:nvGraphicFramePr>.
    src_elm = _copy.deepcopy(shape._element)
    for ph in src_elm.xpath(".//p:ph"):
        ph.getparent().remove(ph)

    # -- Construct a temporary proxy bound to the cleaned element. Use the
    # -- same shape-factory the source slide uses so we pick up Picture /
    # -- GraphicFrame with the right proxy class.
    temp_shape = shape._parent._shape_factory(src_elm)  # pyright: ignore[reportPrivateUsage]
    temp_shape.clone_onto(shape_tree)


def _parse_theme_element(theme_part: Part) -> BaseOxmlElement | None:
    """Return root `a:theme` element of `theme_part` or `None` on failure.

    The theme part may be an |XmlPart| (already parsed) or a plain |Part|
    (raw blob) depending on whether a subtype is registered for it.
    """
    from pptx.oxml import parse_xml

    element = getattr(theme_part, "_element", None)
    if element is not None:
        return cast("BaseOxmlElement", element)
    try:
        return cast("BaseOxmlElement", parse_xml(theme_part.blob))
    except Exception:  # pragma: no cover - defensive
        return None


def _resolve_theme_colors(slide_master_part: Part) -> dict[str, RGBColor]:
    """Return scheme-name-to-|RGBColor| mapping for `slide_master_part`.

    The mapping fuses the slide-master's `p:clrMap` aliases with the RGB
    values found in the related theme part's `a:clrScheme`. Each `a:clrScheme`
    child (e.g. `a:accent1`) contains exactly one color-choice element; we
    resolve it to an RGB value using a best-effort walk of the color element.
    """
    # -- scheme entries carry one of these color elements as their child --
    _SRGB = qn("a:srgbClr")
    _SYSC = qn("a:sysClr")
    _PRST = qn("a:prstClr")

    # -- 1. Follow the theme relationship from this master --
    try:
        theme_part = slide_master_part.part_related_by(RT.THEME)
    except KeyError:
        return {}
    # -- The theme part is registered as a plain Part (content is in its blob);
    # -- parse its XML lazily here so callers don't pay if they never resolve
    # -- a scheme color.
    theme_elm = _parse_theme_element(theme_part)
    if theme_elm is None:
        return {}

    # -- 2. Collect scheme colors from `a:clrScheme` --
    colors: dict[str, RGBColor] = {}
    clrScheme = theme_elm.find(qn("a:themeElements") + "/" + qn("a:clrScheme"))
    if clrScheme is None:
        return colors
    for child in clrScheme:
        # -- strip namespace from entry tag, e.g. "{..}accent1" -> "accent1" --
        tag = child.tag
        name = tag.split("}", 1)[1] if "}" in tag else tag
        inner = next(iter(child), None)
        if inner is None:
            continue
        rgb: RGBColor | None = None
        if inner.tag == _SRGB:
            rgb = RGBColor.from_string(inner.get("val", "000000"))
        elif inner.tag == _SYSC:
            last_clr = inner.get("lastClr")
            if last_clr is not None:
                rgb = RGBColor.from_string(last_clr)
        elif inner.tag == _PRST:
            from pptx.dml._preset_colors import PRESET_COLORS

            prst_name = inner.get("val", "")
            hex_str = PRESET_COLORS.get(prst_name)
            if hex_str is not None:
                rgb = RGBColor.from_string(hex_str)
        if rgb is not None:
            colors[name] = rgb

    # -- 3. Apply the master's `p:clrMap` aliases (bg1/bg2/tx1/tx2) --
    master_elm = getattr(slide_master_part, "_element", None)
    if master_elm is not None:
        clrMap = master_elm.find(qn("p:clrMap"))
        if clrMap is not None:
            for alias in ("bg1", "bg2", "tx1", "tx2"):
                target = clrMap.get(alias)
                if target and target in colors:
                    colors[alias] = colors[target]

    return colors
