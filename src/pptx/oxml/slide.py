"""Slide-related custom element classes, including those for masters."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml import parse_from_template, parse_xml
from pptx.oxml.dml.fill import CT_GradientFillProperties
from pptx.oxml.ns import nsdecls, nsuri, qn
from pptx.oxml.simpletypes import (
    XsdBoolean,
    XsdString,
    XsdStringEnumeration,
    XsdUnsignedInt,
)
from pptx.oxml.timing import (
    CT_SlideTiming,
    CT_TimeNodeList,
    CT_TLCommonTimeNodeData,
    CT_TLShapeTargetElement,
    CT_TLTimeCondition,
    CT_TLTimeConditionList,
    CT_TLTimeNodeParallel,
    CT_TLTimeNodeSequence,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
    ZeroOrOneChoice,
)

__all__ = [
    "CT_Background",
    "CT_BackgroundProperties",
    "CT_CommonSlideData",
    "CT_HeaderFooter",
    "CT_NotesMaster",
    "CT_NotesSlide",
    "CT_Slide",
    "CT_SlideLayout",
    "CT_SlideLayoutIdList",
    "CT_SlideLayoutIdListEntry",
    "CT_SlideMaster",
    "CT_SlideTiming",
    "CT_SlideTransition",
    "CT_TimeNodeList",
    "CT_TLCommonTimeNodeData",
    "CT_TLMediaNodeVideo",
    "CT_TLShapeTargetElement",
    "CT_TLTimeCondition",
    "CT_TLTimeConditionList",
    "CT_TLTimeNodeParallel",
    "CT_TLTimeNodeSequence",
    "CT_SideDirectionTransition",
    "CT_TransitionMorph",
    "CT_TransitionVariant",
]

if TYPE_CHECKING:
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.oxml.tags import CT_CustomerDataList


class _BaseSlideElement(BaseOxmlElement):
    """Base class for the six slide types, providing common methods."""

    cSld: CT_CommonSlideData

    @property
    def spTree(self) -> CT_GroupShape:
        """Return required `p:cSld/p:spTree` grandchild."""
        return self.cSld.spTree


class CT_Background(BaseOxmlElement):
    """`p:bg` element."""

    _insert_bgPr: Callable[[CT_BackgroundProperties], None]

    # ---these two are actually a choice, not a sequence, but simpler for
    # ---present purposes this way.
    _tag_seq = ("p:bgPr", "p:bgRef")
    bgPr: CT_BackgroundProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bgPr", successors=()
    )
    bgRef = ZeroOrOne("p:bgRef", successors=())
    del _tag_seq

    def add_noFill_bgPr(self):
        """Return a new `p:bgPr` element with noFill properties."""
        xml = "<p:bgPr %s>\n" "  <a:noFill/>\n" "  <a:effectLst/>\n" "</p:bgPr>" % nsdecls("a", "p")
        bgPr = cast(CT_BackgroundProperties, parse_xml(xml))
        self._insert_bgPr(bgPr)
        return bgPr


class CT_BackgroundProperties(BaseOxmlElement):
    """`p:bgPr` element."""

    _tag_seq = (
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:blipFill",
        "a:pattFill",
        "a:grpFill",
        "a:effectLst",
        "a:effectDag",
        "a:extLst",
    )
    eg_fillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        ),
        successors=_tag_seq[6:],
    )
    del _tag_seq

    def _new_gradFill(self):
        """Override default to add default gradient subtree."""
        return CT_GradientFillProperties.new_gradFill()


class CT_HeaderFooter(BaseOxmlElement):
    """`p:hf` element, specifying slide-number / header / footer / date placeholder visibility.

    Appears as a direct child of `p:sldLayout`, `p:sldMaster`, `p:notesMaster`, and
    `p:handoutMaster`. Each attribute is an optional boolean with a default of `True`; when the
    attribute is absent the corresponding placeholder is considered visible. Assigning `False`
    hides the placeholder on this layout or master.
    """

    sldNum: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sldNum", XsdBoolean, default=True
    )
    hdr: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "hdr", XsdBoolean, default=True
    )
    ftr: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "ftr", XsdBoolean, default=True
    )
    dt: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dt", XsdBoolean, default=True
    )


class CT_CommonSlideData(BaseOxmlElement):
    """`p:cSld` element."""

    _insert_bg: Callable[[CT_Background], CT_Background]
    _remove_bg: Callable[[], None]
    get_or_add_bg: Callable[[], CT_Background]
    get_or_add_custDataLst: Callable[[], "CT_CustomerDataList"]
    _remove_custDataLst: Callable[[], None]

    _tag_seq = ("p:bg", "p:spTree", "p:custDataLst", "p:controls", "p:extLst")
    bg: CT_Background | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bg", successors=_tag_seq[1:]
    )
    spTree: CT_GroupShape = OneAndOnlyOne("p:spTree")  # pyright: ignore[reportAssignmentType]
    custDataLst: "CT_CustomerDataList | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:custDataLst", successors=_tag_seq[3:]
    )
    del _tag_seq
    name: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString, default=""
    )

    def get_or_add_bgPr(self) -> CT_BackgroundProperties:
        """Return `p:bg/p:bgPr` grandchild.

        If no such grandchild is present, any existing `p:bg` child is first removed and a new
        default `p:bg` with noFill settings is added.
        """
        bg = self.bg
        if bg is None or bg.bgPr is None:
            bg = self._change_to_noFill_bg()
        return cast(CT_BackgroundProperties, bg.bgPr)

    def _change_to_noFill_bg(self) -> CT_Background:
        """Establish a `p:bg` child with no-fill settings.

        Any existing `p:bg` child is first removed.
        """
        self._remove_bg()
        bg = self.get_or_add_bg()
        bg.add_noFill_bgPr()
        return bg


class CT_NotesMaster(_BaseSlideElement):
    """`p:notesMaster` element, root of a notes master part."""

    get_or_add_hf: Callable[[], CT_HeaderFooter]

    _tag_seq = ("p:cSld", "p:clrMap", "p:hf", "p:notesStyle", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[3:]
    )
    del _tag_seq

    @classmethod
    def new_default(cls) -> CT_NotesMaster:
        """Return a new `p:notesMaster` element based on the built-in default template."""
        return cast(CT_NotesMaster, parse_from_template("notesMaster"))


class CT_NotesSlide(_BaseSlideElement):
    """`p:notes` element, root of a notes slide part."""

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    @classmethod
    def new(cls) -> CT_NotesSlide:
        """Return a new ``<p:notes>`` element based on the default template.

        Note that the template does not include placeholders, which must be subsequently cloned
        from the notes master.
        """
        return cast(CT_NotesSlide, parse_from_template("notes"))


class CT_Slide(_BaseSlideElement):
    """`p:sld` element, root element of a slide part (XML document)."""

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    clrMapOvr = ZeroOrOne("p:clrMapOvr", successors=_tag_seq[2:])
    transition: CT_SlideTransition | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:transition", successors=_tag_seq[3:]
    )
    timing: CT_SlideTiming | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:timing", successors=_tag_seq[4:]
    )
    show: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "show", XsdBoolean, default=True
    )
    showMasterSp: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showMasterSp", XsdBoolean, default=True
    )
    del _tag_seq

    @classmethod
    def new(cls) -> CT_Slide:
        """Return new `p:sld` element configured as base slide shape."""
        return cast(CT_Slide, parse_xml(cls._sld_xml()))

    @property
    def bg(self):
        """Return `p:bg` grandchild or None if not present."""
        return self.cSld.bg

    # -- transition resolution (mc:AlternateContent-aware) ----------------
    #
    # PowerPoint writes a MORPH transition wrapped in an
    # `mc:AlternateContent` container — the `mc:Choice` carries the
    # `p:transition/p14:morph` subtree, the `mc:Fallback` carries a
    # `p:transition/p:fade` for pre-2010 viewers. The helpers below let
    # the downstream `Transition` proxy treat a wrapped and a plain
    # `p:transition` uniformly. They live on ``CT_Slide`` (rather than
    # on ``CT_SlideTransition``) because the wrap/unwrap operations
    # must mutate the slide's children.

    @property
    def transition_effective(self) -> CT_SlideTransition | None:
        """The effective `p:transition` element, wherever it lives on this slide.

        Returns the direct `p:sld/p:transition` child when present.
        When no direct child exists but an `mc:AlternateContent`
        wrapper is in `p:transition`'s slot (i.e. a child of this
        `p:sld`) whose first `mc:Choice` holds a `p:transition`,
        returns that wrapped `p:transition`. Returns ``None`` when
        neither form is present. See issue #942 and
        ``docs/dev/analysis/f8-animations-transitions.rst``.
        """
        direct = self.transition
        if direct is not None:
            return direct
        for ac in self._transition_alt_content_lst:
            choices = list(ac.iterchildren(qn("mc:Choice")))
            if not choices:
                continue
            wrapped = choices[0].find(qn("p:transition"))
            if wrapped is not None:
                return cast(CT_SlideTransition, wrapped)
        return None

    def get_or_add_transition_effective(self) -> CT_SlideTransition:
        """Return the effective `p:transition`, adding a direct child if none exists.

        When a wrapped transition already exists (inside
        `mc:AlternateContent`) it is returned as-is. Otherwise a new
        direct `p:transition` child is appended.
        """
        effective = self.transition_effective
        if effective is not None:
            return effective
        return self.get_or_add_transition()

    @property
    def transition_is_alt_content_wrapped(self) -> bool:
        """True when the effective `p:transition` sits inside an `mc:AlternateContent`.

        False when there is no transition, or when the transition is a
        plain child of `p:sld`.
        """
        return self.transition is None and self.transition_effective is not None

    @property
    def _transition_alt_content_lst(self) -> list[BaseOxmlElement]:
        """`mc:AlternateContent` children of `p:sld` that sit in the transition slot.

        An `mc:AlternateContent` wraps a `p:transition` when its first
        `mc:Choice` contains a `p:transition`. PowerPoint puts the
        wrapper in between `p:clrMapOvr` and `p:timing` — i.e. exactly
        where a bare `p:transition` would otherwise sit.
        """
        return cast(
            "list[BaseOxmlElement]",
            [
                ac
                for ac in self.iterchildren(qn("mc:AlternateContent"))
                if self._ac_wraps_transition(ac)
            ],
        )

    @staticmethod
    def _ac_wraps_transition(ac: BaseOxmlElement) -> bool:
        """True when the first `mc:Choice` of `ac` contains a `p:transition`."""
        choices = list(ac.iterchildren(qn("mc:Choice")))
        if not choices:
            return False
        return choices[0].find(qn("p:transition")) is not None

    def wrap_transition_in_alt_content(self) -> CT_SlideTransition:
        """Wrap the direct `p:transition` child in an `mc:AlternateContent`.

        Moves the existing direct `p:transition` into an
        ``mc:Choice Requires="p14"`` and adds a sibling
        ``mc:Fallback`` containing a minimal ``p:transition/p:fade``
        so viewers that don't understand the 2010 extension render
        a plain fade instead of silently dropping the transition.

        If the transition is already wrapped, this is a no-op and the
        existing wrapped `p:transition` is returned. If there is no
        transition at all, a new direct one is created first.

        Returns the effective `p:transition` after wrapping.
        """
        # -- already wrapped? nothing to do --
        if self.transition_is_alt_content_wrapped:
            return cast(CT_SlideTransition, self.transition_effective)

        # -- ensure there is a direct transition to wrap --
        inner = self.get_or_add_transition()

        # -- copy plain-namespace transition attributes (spd/advClick/advTm)
        # -- onto the fallback so an older viewer still gets the same
        # -- advance / speed behaviour. @p14:dur is intentionally omitted
        # -- from the fallback because the `p14` namespace is exactly what
        # -- the fallback is *not* expected to understand.
        fallback_attr_str = "".join(
            ' %s="%s"' % (k, v) for k, v in inner.attrib.items() if not k.startswith("{")
        )

        ac_xml = (
            '<mc:AlternateContent xmlns:mc="%s">\n'
            '  <mc:Choice xmlns:p14="%s" Requires="p14">\n'
            "    <_placeholder_/>\n"
            "  </mc:Choice>\n"
            "  <mc:Fallback>\n"
            '    <p:transition xmlns:p="%s"%s>\n'
            "      <p:fade/>\n"
            "    </p:transition>\n"
            "  </mc:Fallback>\n"
            "</mc:AlternateContent>"
        ) % (nsuri("mc"), nsuri("p14"), nsuri("p"), fallback_attr_str)

        ac = parse_xml(ac_xml)
        # -- replace the placeholder inside mc:Choice with the real transition --
        choice = ac.find(qn("mc:Choice"))
        placeholder = choice.find("_placeholder_")
        # -- detach inner from its current (direct-child) slot first --
        self.remove(inner)
        choice.replace(placeholder, inner)

        # -- insert the wrapper where the direct transition used to be --
        # -- successors are p:timing, p:extLst. --
        inserted = False
        for tag in ("p:timing", "p:extLst"):
            sibling = self.find(qn(tag))
            if sibling is not None:
                sibling.addprevious(ac)
                inserted = True
                break
        if not inserted:
            self.append(ac)

        return cast(CT_SlideTransition, inner)

    def unwrap_transition_from_alt_content(self) -> CT_SlideTransition | None:
        """Move a wrapped `p:transition` back to being a direct `p:sld` child.

        If there's no wrapping `mc:AlternateContent`, this is a no-op.
        The `mc:Fallback` content is discarded. Returns the (now-direct)
        `p:transition` element, or ``None`` when no transition existed.
        """
        ac_lst = self._transition_alt_content_lst
        if not ac_lst:
            return self.transition
        ac = ac_lst[0]
        choices = list(ac.iterchildren(qn("mc:Choice")))
        inner = choices[0].find(qn("p:transition")) if choices else None
        if inner is None:
            # -- nothing to recover; just drop the wrapper --
            self.remove(ac)
            return self.transition
        # -- detach inner transition from mc:Choice, remove the whole
        # -- mc:AlternateContent, then re-insert the transition in its
        # -- schema-correct position.
        choices[0].remove(inner)
        self.remove(ac)
        # -- also remove any stale direct p:transition (shouldn't happen
        # -- in normal operation but guard against double-insertion). --
        existing = self.transition
        if existing is not None:
            self.remove(existing)
        self._insert_transition(inner)
        return cast(CT_SlideTransition, inner)

    def remove_transition_effective(self) -> None:
        """Remove any `p:transition` from this slide, whether direct or mc-wrapped.

        A no-op if no transition is present.
        """
        for ac in list(self._transition_alt_content_lst):
            self.remove(ac)
        direct = self.transition
        if direct is not None:
            self.remove(direct)

    def get_or_add_childTnLst(self):
        """Return parent element for a new `p:video` child element.

        The `p:video` element causes play controls to appear under a video
        shape (pic shape containing video). There can be more than one video
        shape on a slide, which causes the precondition to vary.

        The method handles three cases:

        1. No `p:timing` anywhere on the slide -- a fresh one is created.
        2. An existing `p:timing` (plain or wrapped in
           ``mc:AlternateContent``/``mc:Choice``, see issue #954) already
           contains the expected ``p:tnLst/p:par/p:cTn/p:childTnLst``
           descendant -- that descendant is returned so the new video
           merges alongside any pre-existing animation/media entries.
        3. An existing `p:timing` is present but does not match the
           expected structure -- it is replaced with a freshly built
           ``p:timing`` subtree (preserving the ``mc:AlternateContent``
           wrapper when one is in use).
        """
        childTnLst = self._childTnLst
        if childTnLst is None:
            childTnLst = self._add_childTnLst()
        return childTnLst

    def _add_childTnLst(self):
        """Add `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.

        Any existing `p:timing` child element is replaced with a freshly
        built minimal timing subtree. When the existing `p:timing` is
        wrapped inside an `mc:AlternateContent/mc:Choice` (as produced
        when PowerPoint authors a slide whose timing references the
        2010+ extension namespace) the replacement happens *inside* the
        wrapper so that only one `p:timing` element ends up on the
        slide -- see issue #954.
        """
        existing_timing, existing_parent = self._existing_timing_and_parent
        new_timing = parse_xml(self._childTnLst_timing_xml())
        if existing_timing is not None and existing_parent is not None:
            existing_parent.replace(existing_timing, new_timing)
        else:
            # --- no pre-existing timing anywhere; insert at the schema-correct
            # --- position among p:sld's direct children ---
            self._insert_timing(new_timing)
        return new_timing.xpath("./p:tnLst/p:par/p:cTn/p:childTnLst")[0]

    @property
    def _childTnLst(self):
        """Return `p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.

        Searches both the plain ``./p:timing`` location and the
        ``mc:AlternateContent/mc:Choice`` wrapped location introduced
        by PowerPoint 2010+ timing extensions (issue #954). Returns
        ``None`` when no `p:timing` with the expected subtree is
        present.
        """
        timing, _ = self._existing_timing_and_parent
        if timing is None:
            return None
        childTnLsts = timing.xpath("./p:tnLst/p:par/p:cTn/p:childTnLst")
        if not childTnLsts:
            return None
        return childTnLsts[0]

    @property
    def _existing_timing_and_parent(self):
        """The first `p:timing` element on the slide and its direct parent.

        Returns ``(timing, parent)`` where both items are ``None`` when
        no `p:timing` is present. ``parent`` is ``self`` (the `p:sld`
        element) when `p:timing` is an unwrapped direct child, or an
        ``mc:Choice`` element when the timing has been wrapped inside
        an ``mc:AlternateContent`` block.

        Only the *first* ``mc:Choice`` of any ``mc:AlternateContent``
        is searched — the preferred-rendering slot. ``mc:Fallback``
        timing is intentionally ignored; the fallback subtree is
        preserved as-is on save.
        """
        # --- plain direct-child `p:timing` ---
        timing = self.timing
        if timing is not None:
            return timing, self
        # --- fall back to the wrapped form used by PowerPoint 2010+ ---
        for ac in self.iterchildren(qn("mc:AlternateContent")):
            choices = list(ac.iterchildren(qn("mc:Choice")))
            if not choices:
                continue
            choice = choices[0]
            wrapped = choice.find(qn("p:timing"))
            if wrapped is not None:
                return wrapped, choice
        return None, None

    @staticmethod
    def _childTnLst_timing_xml():
        return (
            "<p:timing %s>\n"
            "  <p:tnLst>\n"
            "    <p:par>\n"
            '      <p:cTn id="1" dur="indefinite" restart="never" nodeType="'
            'tmRoot">\n'
            "        <p:childTnLst/>\n"
            "      </p:cTn>\n"
            "    </p:par>\n"
            "  </p:tnLst>\n"
            "</p:timing>" % nsdecls("p")
        )

    @staticmethod
    def _sld_xml():
        return (
            "<p:sld %s>\n"
            "  <p:cSld>\n"
            "    <p:spTree>\n"
            "      <p:nvGrpSpPr>\n"
            '        <p:cNvPr id="1" name=""/>\n'
            "        <p:cNvGrpSpPr/>\n"
            "        <p:nvPr/>\n"
            "      </p:nvGrpSpPr>\n"
            "      <p:grpSpPr/>\n"
            "    </p:spTree>\n"
            "  </p:cSld>\n"
            "  <p:clrMapOvr>\n"
            "    <a:masterClrMapping/>\n"
            "  </p:clrMapOvr>\n"
            "</p:sld>" % nsdecls("a", "p", "r")
        )


class CT_SlideLayout(_BaseSlideElement):
    """`p:sldLayout` element, root of a slide layout part."""

    get_or_add_hf: Callable[[], CT_HeaderFooter]

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:hf", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[5:]
    )
    del _tag_seq
    # -- `@type` identifies the layout kind ("title", "obj", "twoObj", "cust", ...).
    # -- Per ECMA-376 Part 1 §19.3.1.39 / `CT_SlideLayout`, the attribute is optional
    # -- and defaults to "cust". PowerPoint always writes it; Google Slides exports
    # -- routinely omit it (issue #864). Exposed as a raw string rather than an
    # -- enum because downstream consumers typically pattern-match it and the
    # -- full ST_SlideLayoutType vocabulary (36 values) has no precedent elsewhere
    # -- in the library's enum layer.
    type: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", XsdString, default="cust"
    )

    @classmethod
    def new_blank(cls, name: str = "") -> CT_SlideLayout:
        """Return a new `p:sldLayout` element configured as a minimal blank layout.

        The returned layout carries ``@type="cust"`` (the schema default for custom
        layouts), ``@preserve="1"``, a ``p:cSld/@name`` matching `name`, and an
        empty ``p:spTree`` (no placeholders). A ``p:clrMapOvr/a:masterClrMapping``
        element is included so the layout's color-map inherits from its master.
        """
        name_attr = ' name="%s"' % name if name else ""
        xml = (
            '<p:sldLayout %s type="cust" preserve="1">\n'
            "  <p:cSld%s>\n"
            "    <p:spTree>\n"
            "      <p:nvGrpSpPr>\n"
            '        <p:cNvPr id="1" name=""/>\n'
            "        <p:cNvGrpSpPr/>\n"
            "        <p:nvPr/>\n"
            "      </p:nvGrpSpPr>\n"
            "      <p:grpSpPr/>\n"
            "    </p:spTree>\n"
            "  </p:cSld>\n"
            "  <p:clrMapOvr>\n"
            "    <a:masterClrMapping/>\n"
            "  </p:clrMapOvr>\n"
            "</p:sldLayout>" % (nsdecls("a", "p", "r"), name_attr)
        )
        return cast(CT_SlideLayout, parse_xml(xml))


class CT_SlideLayoutIdList(BaseOxmlElement):
    """`p:sldLayoutIdLst` element, child of `p:sldMaster`.

    Contains references to the slide layouts that inherit from the slide master.
    """

    sldLayoutId_lst: list[CT_SlideLayoutIdListEntry]

    _add_sldLayoutId: Callable[..., CT_SlideLayoutIdListEntry]
    sldLayoutId = ZeroOrMore("p:sldLayoutId")

    def add_sldLayoutId(self, rId: str) -> CT_SlideLayoutIdListEntry:
        """Create and return a new `p:sldLayoutId` child referencing `rId`.

        The new entry receives a freshly-allocated `@id` attribute in the
        ST_SlideLayoutId range (2147483648 .. 4294967295). The next integer
        above the highest value currently in use is picked; when that would
        overflow the range, the lowest unused value from the bottom is chosen.
        """
        return self._add_sldLayoutId(id=self._next_id, rId=rId)

    @property
    def _next_id(self) -> int:
        """Next available `p:sldLayoutId/@id` value as an `int`."""
        MIN_LAYOUT_ID = 2147483648
        MAX_LAYOUT_ID = 4294967295

        used_ids = [int(s) for s in cast("list[str]", self.xpath("./p:sldLayoutId/@id"))]
        simple_next = max([MIN_LAYOUT_ID - 1] + used_ids) + 1
        if simple_next <= MAX_LAYOUT_ID:
            return simple_next

        # -- fall back to search for next unused from bottom --
        valid_used_ids = sorted(id for id in used_ids if MIN_LAYOUT_ID <= id <= MAX_LAYOUT_ID)
        return next(
            candidate_id
            for candidate_id, used_id in enumerate(valid_used_ids, start=MIN_LAYOUT_ID)
            if candidate_id != used_id
        )


class CT_SlideLayoutIdListEntry(BaseOxmlElement):
    """`p:sldLayoutId` element, child of `p:sldLayoutIdLst`.

    Contains a reference to a slide layout.
    """

    # -- `@id` is the presentation-stable layout identifier (ECMA-376 Part 1
    # -- §19.3.1.41 / ST_SlideLayoutId). Optional per the schema; PowerPoint
    # -- always writes it, but Google-Slides exports occasionally omit it.
    # -- Values start at 2147483648, which fits inside XsdUnsignedInt's 32-bit
    # -- unsigned range (0-4294967295).
    id: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "id", XsdUnsignedInt
    )
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideMaster(_BaseSlideElement):
    """`p:sldMaster` element, root of a slide master part."""

    get_or_add_hf: Callable[[], CT_HeaderFooter]
    get_or_add_sldLayoutIdLst: Callable[[], CT_SlideLayoutIdList]

    _tag_seq = (
        "p:cSld",
        "p:clrMap",
        "p:sldLayoutIdLst",
        "p:transition",
        "p:timing",
        "p:hf",
        "p:txStyles",
        "p:extLst",
    )
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    sldLayoutIdLst: CT_SlideLayoutIdList = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldLayoutIdLst", successors=_tag_seq[3:]
    )
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[6:]
    )
    del _tag_seq


class CT_TLMediaNodeVideo(BaseOxmlElement):
    """`p:video` element, specifying video media details."""

    _tag_seq = ("p:cMediaNode",)
    cMediaNode = OneAndOnlyOne("p:cMediaNode")
    del _tag_seq


# ---------------------------------------------------------------------------
# Slide-transition element classes (Foundation F8 MVP)
# ---------------------------------------------------------------------------
#
# These classes surface the `p:transition` element and its variant children
# (`p:fade`, `p:push`, ..., `p14:morph`) so a slide can round-trip any
# PowerPoint-authored transition and python-pptx code can read/write the
# common attributes (`@spd`, `@advClick`, `@advTm`). See
# `docs/dev/analysis/f8-animations-transitions.rst` for scope details.


# -- full set of transition-variant choice tags (from pml.xsd CT_SlideTransition).
# -- ``p14:morph`` is handled specially because it lives in the 2010 extension
# -- namespace, not in the core ``p:`` namespace.
_TRANSITION_VARIANT_TAGS_P = (
    "p:blinds",
    "p:checker",
    "p:circle",
    "p:dissolve",
    "p:comb",
    "p:cover",
    "p:cut",
    "p:diamond",
    "p:fade",
    "p:newsflash",
    "p:plus",
    "p:pull",
    "p:push",
    "p:random",
    "p:randomBar",
    "p:split",
    "p:strips",
    "p:wedge",
    "p:wheel",
    "p:wipe",
    "p:zoom",
)
_TRANSITION_VARIANT_TAG_P14_MORPH = "p14:morph"


class CT_SlideTransition(BaseOxmlElement):
    """`p:transition` element — specifies a slide transition.

    Schema excerpt (from ``pml.xsd`` ``CT_SlideTransition``)::

        <xsd:sequence>
          <xsd:choice minOccurs="0" maxOccurs="1">
            <xsd:element name="blinds"    .../>
            ...
            <xsd:element name="zoom"      .../>
          </xsd:choice>
          <xsd:element name="sndAc"  minOccurs="0" .../>
          <xsd:element name="extLst" minOccurs="0" .../>
        </xsd:sequence>
        <xsd:attribute name="spd"      type="ST_TransitionSpeed" default="fast"/>
        <xsd:attribute name="advClick" type="xsd:boolean"        default="true"/>
        <xsd:attribute name="advTm"    type="xsd:unsignedInt"/>

    MVP notes:

    * The variant child-element choice is not expressed as a
      :class:`ZeroOrOneChoice` because the ``p14:morph`` sibling (Office
      2010 extension) sits in a different namespace and
      :class:`Choice` only groups tags in one namespace in the existing
      xmlchemy. Instead, the :meth:`variant_tag` / :meth:`set_variant`
      helpers handle variant selection explicitly. Downstream #942
      (full MORPH support, including ``mc:AlternateContent`` wrapping)
      will extend this.
    * `@spd` is the coarse ``slow`` / ``med`` / ``fast`` selector from
      ST_TransitionSpeed. For millisecond-precision duration PowerPoint
      writes the ``p14:dur`` attribute — *not* a standard ISO-29500
      field. :class:`.Transition.duration` reads/writes the ``p14:dur``
      attribute when present (see MVP scope in the F8 analysis doc).
    """

    spd: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "spd", XsdString, default="fast"
    )
    advClick: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "advClick", XsdBoolean, default=True
    )
    advTm = OptionalAttribute("advTm", XsdUnsignedInt)

    @property
    def variant_tag(self) -> str | None:
        """Namespaced-prefixed tag name of the transition variant child, or None.

        Returns strings like ``"p:fade"``, ``"p:push"``, or
        ``"p14:morph"``. Returns ``None`` when ``p:transition`` has no
        variant child (meaning "rely on application defaults" per the
        schema's ``minOccurs=0`` on the inner choice).
        """
        for child in self:
            tag = child.tag
            for p_tag in _TRANSITION_VARIANT_TAGS_P:
                if tag == qn(p_tag):
                    return p_tag
            if tag == qn(_TRANSITION_VARIANT_TAG_P14_MORPH):
                return _TRANSITION_VARIANT_TAG_P14_MORPH
        return None

    def _remove_variant(self) -> None:
        """Remove any existing transition-variant child element."""
        variant_tags = _TRANSITION_VARIANT_TAGS_P + (_TRANSITION_VARIANT_TAG_P14_MORPH,)
        for child in list(self):
            for tag in variant_tags:
                if child.tag == qn(tag):
                    self.remove(child)
                    break

    def set_variant(self, nsptag: str | None) -> BaseOxmlElement | None:
        """Replace the transition-variant child with a new one named `nsptag`.

        `nsptag` is one of the sentinel tags listed in
        :data:`_TRANSITION_VARIANT_TAGS_P` (e.g. ``"p:fade"``) or
        ``"p14:morph"``. Passing ``None`` removes the variant child.
        Returns the newly created child element, or ``None`` when a
        ``None`` tag was passed.

        The inserted element is minimal — just the bare tag. Downstream
        items can layer on direction / orientation / loop-sound
        attributes without touching this MVP helper.
        """
        self._remove_variant()
        if nsptag is None:
            return None
        nspfx = nsptag.split(":", 1)[0]
        if nspfx == "p14":
            xml_str = '<p14:morph xmlns:p14="%s"/>' % nsuri("p14")
        else:
            local = nsptag.split(":", 1)[1]
            xml_str = '<p:%s xmlns:p="%s"/>' % (local, nsuri("p"))
        new_child = parse_xml(xml_str)
        # -- the variant child must precede `p:sndAc` and `p:extLst`; since the
        # -- existing `p:sndAc` / `p:extLst` are not expected in common fixtures
        # -- we insert at index 0, which the schema permits because the choice
        # -- comes first in the sequence.
        self.insert(0, new_child)
        return new_child

    @property
    def dur(self) -> int | None:
        """Value of the ``p14:dur`` attribute in ms, or ``None`` when absent.

        This is the PowerPoint-2010 extension attribute Microsoft added
        for per-transition precise duration. It complements the ISO
        ``@spd`` attribute (slow / med / fast) with a ms value. Missing
        from ISO-29500 but present in every modern file.
        """
        val = self.get(qn("p14:dur"))
        if val is None:
            return None
        return int(val)

    @dur.setter
    def dur(self, value: int | None) -> None:
        attr = qn("p14:dur")
        if value is None:
            if attr in self.attrib:
                del self.attrib[attr]
            return
        if not isinstance(value, int) or value < 0:
            raise ValueError(
                "transition duration must be a non-negative int (milliseconds), got %r" % value
            )
        self.set(attr, str(value))


class CT_TransitionVariant(BaseOxmlElement):
    """Base class for every transition-variant choice element.

    Most variants (``p:circle``, ``p:dissolve``, ``p:diamond``,
    ``p:newsflash``, ``p:plus``, ``p:random``, ``p:wedge``) are the
    empty ``CT_Empty`` type and carry no attributes. The directional
    / orientation / loop variants that *do* have attributes
    (``CT_EightDirectionTransition``, ``CT_OptionalBlackTransition``,
    etc.) would ordinarily get subclasses, but for MVP round-trip
    purposes a single class registered for every variant tag
    preserves the XML byte-for-byte via lxml's default attribute
    handling. Downstream items (#942 MORPH, #256 timings) will add
    typed descriptors per variant as needed.
    """


class ST_TransitionMorphOption(XsdStringEnumeration):
    """Valid values for ``p14:morph/@option``.

    PowerPoint emits one of three tokens (``byObject``, ``byWord``,
    ``byChar``) when a MORPH transition carries an explicit matching
    granularity. The default (when the attribute is absent) is
    ``byObject``.
    """

    BY_OBJECT = "byObject"
    BY_WORD = "byWord"
    BY_CHAR = "byChar"

    _members = (BY_OBJECT, BY_WORD, BY_CHAR)


class CT_TransitionMorph(CT_TransitionVariant):
    """`p14:morph` — the Office 2010 MORPH transition element.

    Carries the ``@option`` attribute, one of ``byObject`` /
    ``byWord`` / ``byChar`` (default ``byObject``). Assignments are
    validated against :class:`ST_TransitionMorphOption`.

    The higher-level proxy is :attr:`pptx.slide.Transition.morph_option`.
    See issue #942 and ``docs/dev/analysis/f8-animations-transitions.rst``
    for the full plan, including the ``mc:AlternateContent`` wrapper
    that wraps a written MORPH transition.
    """

    option: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "option", ST_TransitionMorphOption, default="byObject"
    )


class CT_SideDirectionTransition(CT_TransitionVariant):
    """Typed variant for side-direction transitions (``p:wipe``, ``p:push``).

    Reproduces ``CT_SideDirectionTransition`` from ``pml.xsd``. Carries
    a single ``@dir`` attribute drawn from ``ST_TransitionSideDirectionType``
    (``l`` / ``u`` / ``r`` / ``d``). The schema default is ``"l"``.

    Added for #1004 to surface the ``@dir`` attribute on wipe/push
    transitions. Additional direction-variant sub-types (eight-direction,
    corner-direction, orientation, in-out) can follow the same pattern
    incrementally.
    """

    dir: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", XsdString, default="l"
    )
