"""Slide-related objects, including masters, layouts, and notes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast

from pptx.dml.color import RGBColor
from pptx.dml.fill import FillFormat
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.transition import PP_TRANSITION_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
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
    from pptx.opc.package import Part
    from pptx.comments import Comments
    from pptx.oxml.presentation import CT_SlideId, CT_SlideIdList, CT_SlideMasterIdList
    from pptx.oxml.slide import (
        CT_CommonSlideData,
        CT_NotesMaster,
        CT_NotesSlide,
        CT_Slide,
        CT_SlideLayout,
        CT_SlideLayoutIdList,
        CT_SlideMaster,
    )
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart
    from pptx.presentation import Presentation
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
        return _Background(self._element.cSld)

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
    def shapes(self):
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
        """
        return self.part.comments_part is not None

    @property
    def follow_master_background(self):
        """|True| if this slide inherits the slide master background.

        Assigning |False| causes background inheritance from the master to be
        interrupted; if there is no custom background for this slide,
        a default background is added. If a custom background already exists
        for this slide, assigning |False| has no effect.

        Assigning |True| causes any custom background for this slide to be
        deleted and inheritance from the master restored.
        """
        return self._element.bg is None

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
        """
        timing = self._element.timing
        if timing is None:
            return None
        return timing.xml

    @property
    def animation_sequence(self) -> tuple[AnimationEffect, ...]:
        """Read-only tuple of |AnimationEffect| in this slide's main sequence.

        Returns each entrance / emphasis / exit / motion-path effect
        PowerPoint would step through when the slide's main animation
        sequence runs, in document order.

        Returns an empty tuple when the slide has no ``p:timing`` subtree
        or when the subtree is media-playback-only (the stub tree added
        by :meth:`.SlideShapes.add_movie` has no main sequence).

        Each :class:`.AnimationEffect` exposes a handful of read-only
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
        """
        from pptx.oxml.timing import iter_main_sequence_effects

        timing = self._element.timing
        if timing is None:
            return ()
        return tuple(AnimationEffect(par) for par in iter_main_sequence_effects(timing))

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

    def add_slide(self, slide_layout: SlideLayout) -> Slide:
        """Return a newly added slide that inherits layout from `slide_layout`."""
        rId, slide = self.part.add_slide(slide_layout)
        slide.shapes.clone_layout_placeholders(slide_layout)
        self._sldIdLst.add_sldId(rId)
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

        This method implements the *basic* cross-presentation slide copy path.
        It handles slides whose only relationships are to the slide layout, to
        image parts, to external hyperlinks, and (dropped) to notes slides.
        Any other related part type -- such as charts, embedded OLE objects,
        or media -- raises |NotImplementedError|. Full-fidelity cross-part
        relationship cloning is tracked as Foundation F1.
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

    @lazyproperty
    def header_footer(self) -> _HeaderFooter:
        """|_HeaderFooter| object controlling header/footer/slide-number/date visibility.

        When absent, the layout inherits its master's header/footer settings; accessing this
        property does not itself add a `p:hf` element to the layout XML. See |_HeaderFooter|.
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
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide-layout inherits properties."""
        return self.part.slide_master

    @property
    def used_by_slides(self):
        """Tuple of slide objects based on this slide layout."""
        # ---getting Slides collection requires going around the horn a bit---
        slides = self.part.package.presentation_part.presentation.slides
        return tuple(s for s in slides if s.slide_layout == self)


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

    def get_by_name(self, name: str, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return SlideLayout object having `name`, or `default` if not found."""
        for slide_layout in self:
            if slide_layout.name == name:
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

    @lazyproperty
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| object providing access to this slide-master's layouts."""
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)

    @lazyproperty
    def theme_colors(self) -> dict[str, RGBColor]:
        """Mapping of scheme-color name (e.g. ``"accent1"``) to |RGBColor|.

        Keys are the scheme-color names from the theme's `a:clrScheme`
        (``"dk1"``, ``"lt1"``, ``"dk2"``, ``"lt2"``, ``"accent1"``..``"accent6"``,
        ``"hlink"``, ``"folHlink"``) plus the color-map aliases from this
        master's ``p:clrMap`` (``"bg1"``, ``"bg2"``, ``"tx1"``, ``"tx2"``).

        Useful for calling :meth:`ColorFormat.to_rgb` on a scheme color.
        """
        return _resolve_theme_colors(self.part)


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

    MVP surface:

    * :attr:`type` — :class:`.PP_TRANSITION_TYPE` enum member.
      Reading returns :attr:`PP_TRANSITION_TYPE.NONE` when there is
      no transition variant present.
    * :attr:`duration` — int milliseconds or `None`. Read/write of
      the ``p14:dur`` extension attribute. ``@spd`` (slow/med/fast)
      is left at its default when setting; downstream items can
      expose a separate speed accessor.
    * :attr:`advance_on_click` — bool, ``p:transition/@advClick``.
    * :attr:`advance_after_time` — int ms or `None`, ``p:transition/@advTm``.

    Out-of-scope (see ``docs/dev/analysis/f8-animations-transitions.rst``):

    * Per-variant attributes such as ``@dir`` on ``p:fade`` or
      ``p:cover`` (downstream #1004).
    * ``mc:AlternateContent`` wrapping for MORPH (downstream #942).
    * Sound-action ``p:sndAc`` (not tracked as a separate item).
    """

    def __init__(self, sld: CT_Slide):
        super(Transition, self).__init__(sld)
        self._sld = sld

    # -- type --------------------------------------------------------

    @property
    def type(self) -> PP_TRANSITION_TYPE:
        """The :class:`.PP_TRANSITION_TYPE` selected on this slide.

        Returns :attr:`PP_TRANSITION_TYPE.NONE` when the slide has no
        ``p:transition`` element, or when ``p:transition`` has no
        variant child (it is valid for ``p:transition`` to carry only
        attributes such as ``@advTm``).
        """
        transition = self._sld.transition
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
            transition = self._sld.transition
            if transition is None:
                return
            transition._remove_variant()
            # -- if the transition element is now empty of attributes and
            # -- children, drop it for minimal XML round-trip.
            if not transition.attrib and len(transition) == 0:
                self._sld.remove(transition)
            return
        PP_TRANSITION_TYPE.validate(value)
        transition = self._sld.get_or_add_transition()
        # -- MORPH lives in the p14 namespace; everything else in p: --
        if value is PP_TRANSITION_TYPE.MORPH:
            transition.set_variant("p14:morph")
        else:
            transition.set_variant("p:%s" % value.xml_value)

    # -- duration (ms) -----------------------------------------------

    @property
    def duration(self) -> int | None:
        """Transition duration in milliseconds, or ``None`` when unspecified.

        Reads the ``p14:dur`` attribute (Office 2010 extension). When
        absent the property returns ``None`` — PowerPoint then falls
        back to a default derived from ``@spd`` (slow=1600ms,
        med=1000ms, fast=500ms).
        """
        transition = self._sld.transition
        if transition is None:
            return None
        return transition.dur

    @duration.setter
    def duration(self, value: int | None) -> None:
        if value is None:
            transition = self._sld.transition
            if transition is None:
                return
            transition.dur = None
            if not transition.attrib and len(transition) == 0:
                self._sld.remove(transition)
            return
        transition = self._sld.get_or_add_transition()
        transition.dur = value

    # -- advance_on_click --------------------------------------------

    @property
    def advance_on_click(self) -> bool:
        """Whether a mouse click advances past the transition.

        Reflects ``p:transition/@advClick``. Defaults to ``True``
        (schema default) when no ``p:transition`` element exists or
        the attribute is absent.
        """
        transition = self._sld.transition
        if transition is None:
            return True
        return transition.advClick

    @advance_on_click.setter
    def advance_on_click(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise TypeError(
                "advance_on_click must be a bool, got %s" % type(value).__name__
            )
        transition = self._sld.get_or_add_transition()
        transition.advClick = value

    # -- advance_after_time ------------------------------------------

    @property
    def advance_after_time(self) -> int | None:
        """Auto-advance delay in ms, or ``None`` when not auto-advancing.

        Reflects ``p:transition/@advTm``. The attribute is absent when
        the slide only advances on click; setting the property to
        ``None`` removes the attribute, restoring click-only advance.
        """
        transition = self._sld.transition
        if transition is None:
            return None
        return transition.advTm

    @advance_after_time.setter
    def advance_after_time(self, value: int | None) -> None:
        if value is None:
            transition = self._sld.transition
            if transition is None:
                return
            if "advTm" in transition.attrib:
                del transition.attrib["advTm"]
            if not transition.attrib and len(transition) == 0:
                self._sld.remove(transition)
            return
        if not isinstance(value, int) or value < 0:
            raise ValueError(
                "advance_after_time must be a non-negative int (milliseconds), got %r"
                % (value,)
            )
        transition = self._sld.get_or_add_transition()
        transition.advTm = value


class AnimationEffect(ElementProxy):
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
    ``AnimationEffect.set_preset()``, ``AnimationSequence.add(...)``,
    reorder / delete — are deferred to the animation-authoring items
    (#102, #264, #1106) that stack on this foundation.

    Instances are not constructed by client code; obtain them from
    :attr:`.Slide.animation_sequence`.
    """

    def __init__(self, par):
        super(AnimationEffect, self).__init__(par)
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


class _Background(ElementProxy):
    """Provides access to slide background properties.

    Note that the presence of this object does not by itself imply an
    explicitly-defined background; a slide with an inherited background still
    has a |_Background| object.
    """

    def __init__(self, cSld: CT_CommonSlideData):
        super(_Background, self).__init__(cSld)
        self._cSld = cSld

    @lazyproperty
    def fill(self):
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
        return FillFormat.from_fill_parent(bgPr)


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
