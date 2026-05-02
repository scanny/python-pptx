"""Custom element classes for the view-properties part (``ppt/viewProps.xml``).

The view-properties part records editor-view settings that PowerPoint
restores when the presentation is re-opened — which editor view was last
active (``@lastView``), whether the comments pane is visible
(``@showComments``), and per-view zoom / scroll-position state. See
``docs/dev/analysis/view-props.rst`` for the design notes.
"""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_ViewType,
    XsdBoolean,
    XsdLong,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)

# -- default ``p:viewPr`` element used when the package does not yet have one.
# -- matches the minimum PowerPoint tolerates: just a ``<p:viewPr/>`` root.
_VIEW_PROPS_DEFAULT_XML = "<p:viewPr %s/>" % nsdecls("p", "a")

# -- XML fragment used to bootstrap a ``p:cViewPr`` subtree (with a ``p:scale``
# -- of 1.0 and an origin at 0,0). Wrapped in the appropriate parent element by
# -- the caller.
_CVIEWPR_DEFAULT_XML = (
    '<p:cViewPr %s><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale>'
    '<p:origin x="0" y="0"/></p:cViewPr>' % nsdecls("p", "a")
)


class CT_ViewProperties(BaseOxmlElement):
    """`p:viewPr` element, root of the view-properties part."""

    get_or_add_normalViewPr: Callable[[], "CT_NormalViewProperties"]
    get_or_add_slideViewPr: Callable[[], "CT_SlideViewProperties"]
    get_or_add_notesViewPr: Callable[[], "CT_NotesViewProperties"]
    get_or_add_outlineViewPr: Callable[[], "CT_OutlineViewProperties"]
    get_or_add_sorterViewPr: Callable[[], "CT_SlideSorterViewProperties"]

    normalViewPr: "CT_NormalViewProperties | None" = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:normalViewPr",
            successors=(
                "p:slideViewPr",
                "p:outlineViewPr",
                "p:notesTextViewPr",
                "p:sorterViewPr",
                "p:notesViewPr",
                "p:gridSpacing",
                "p:extLst",
            ),
        )
    )
    slideViewPr: "CT_SlideViewProperties | None" = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:slideViewPr",
            successors=(
                "p:outlineViewPr",
                "p:notesTextViewPr",
                "p:sorterViewPr",
                "p:notesViewPr",
                "p:gridSpacing",
                "p:extLst",
            ),
        )
    )
    outlineViewPr: "CT_OutlineViewProperties | None" = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:outlineViewPr",
            successors=(
                "p:notesTextViewPr",
                "p:sorterViewPr",
                "p:notesViewPr",
                "p:gridSpacing",
                "p:extLst",
            ),
        )
    )
    sorterViewPr: "CT_SlideSorterViewProperties | None" = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:sorterViewPr",
            successors=("p:notesViewPr", "p:gridSpacing", "p:extLst"),
        )
    )
    notesViewPr: "CT_NotesViewProperties | None" = (
        ZeroOrOne(  # pyright: ignore[reportAssignmentType]
            "p:notesViewPr", successors=("p:gridSpacing", "p:extLst")
        )
    )

    lastView: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "lastView", ST_ViewType, default="sldView"
    )
    showComments: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showComments", XsdBoolean, default=True
    )

    @staticmethod
    def new_default() -> "CT_ViewProperties":
        """Return a freshly-parsed, empty ``<p:viewPr/>`` element.

        Suitable as the starting point for a default :class:`.ViewPropsPart`.
        """
        return cast("CT_ViewProperties", parse_xml(_VIEW_PROPS_DEFAULT_XML))


class CT_CommonViewProperties(BaseOxmlElement):
    """`p:cViewPr` element, common scale + origin state shared by every view.

    Carries the zoom ratio (``p:scale``) and the scroll-position origin
    (``p:origin``) of the containing view. The element itself is modeled
    loosely — the library reads / writes the zoom ratio on its required
    ``p:scale`` child via :class:`CT_Scale2D`, and passes the rest through
    verbatim.
    """

    get_or_add_scale: Callable[[], "CT_Scale2D"]

    scale: "CT_Scale2D | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:scale", successors=("p:origin",)
    )

    @staticmethod
    def new_default() -> "CT_CommonViewProperties":
        """Return a freshly-parsed, zoom=1.0 ``<p:cViewPr>`` subtree."""
        return cast("CT_CommonViewProperties", parse_xml(_CVIEWPR_DEFAULT_XML))


class CT_SlideViewProperties(BaseOxmlElement):
    """`p:slideViewPr` element — slide-editing view state."""

    get_or_add_cSldViewPr: Callable[[], "CT_CommonSlideViewProperties"]

    cSldViewPr: "CT_CommonSlideViewProperties | None" = ZeroOrOne(
        "p:cSldViewPr"
    )  # pyright: ignore[reportAssignmentType]


class CT_NotesViewProperties(BaseOxmlElement):
    """`p:notesViewPr` element — notes-page view state."""

    get_or_add_cSldViewPr: Callable[[], "CT_CommonSlideViewProperties"]

    cSldViewPr: "CT_CommonSlideViewProperties | None" = ZeroOrOne(
        "p:cSldViewPr"
    )  # pyright: ignore[reportAssignmentType]


class CT_CommonSlideViewProperties(BaseOxmlElement):
    """`p:cSldViewPr` element — common slide-view wrapper around ``p:cViewPr``."""

    get_or_add_cViewPr: Callable[[], CT_CommonViewProperties]

    cViewPr: CT_CommonViewProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:cViewPr", successors=("p:guideLst",)
    )


class CT_OutlineViewProperties(BaseOxmlElement):
    """`p:outlineViewPr` element — outline-view state."""

    get_or_add_cViewPr: Callable[[], CT_CommonViewProperties]

    cViewPr: CT_CommonViewProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:cViewPr", successors=("p:sldLst", "p:extLst")
    )


class CT_SlideSorterViewProperties(BaseOxmlElement):
    """`p:sorterViewPr` element — slide-sorter-view state.

    Carries the ``@showFormatting`` attribute (whether the sorter renders
    text in its slide thumbnails) in addition to the common ``p:cViewPr``
    zoom / origin state.
    """

    get_or_add_cViewPr: Callable[[], CT_CommonViewProperties]

    cViewPr: CT_CommonViewProperties | None = ZeroOrOne(
        "p:cViewPr", successors=("p:extLst",)
    )  # pyright: ignore[reportAssignmentType]
    showFormatting: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "showFormatting", XsdBoolean, default=True
    )


class CT_NormalViewProperties(BaseOxmlElement):
    """`p:normalViewPr` element — normal-view pane state.

    Per the schema this element does *not* carry a zoom or origin — those
    live on the sibling ``p:slideViewPr`` / ``p:notesViewPr`` / ``…``
    elements under ``p:viewPr``. The normal-view element records the
    left / top pane-split positions and a handful of UI toggles.

    Additional attributes (``showOutlineIcons``, ``snapVertSplitter``,
    etc.) are not modeled yet; they are round-tripped verbatim by lxml.
    """


class CT_Scale2D(BaseOxmlElement):
    """`p:scale` element — zoom ratio for the containing ``p:cViewPr``.

    Has required child elements ``a:sx`` (horizontal ratio) and ``a:sy``
    (vertical ratio), each of type :class:`CT_Ratio`.
    """

    get_or_add_sx: Callable[[], "CT_Ratio"]
    get_or_add_sy: Callable[[], "CT_Ratio"]

    sx: "CT_Ratio | None" = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:sx", successors=("a:sy",)
    )
    sy: "CT_Ratio | None" = ZeroOrOne("a:sy")  # pyright: ignore[reportAssignmentType]

    @property
    def zoom(self) -> float:
        """The zoom of the containing view as a `float` ratio.

        Returns ``sx.n / sx.d`` (e.g. 1.24 for ``<a:sx n="124" d="100"/>``).
        Returns 1.0 when ``a:sx`` is absent or has ``d = 0``.
        """
        sx = self.sx
        if sx is None or sx.d == 0:
            return 1.0
        return sx.n / sx.d

    def set_zoom(self, value: float) -> None:
        """Write `value` (e.g. 1.0 for 100%) into both ``a:sx`` and ``a:sy``.

        The ratio is encoded with ``d = 100000`` to keep full precision for
        common percentage values without accumulating float drift on repeat
        save/open cycles.
        """
        d = 100000
        n = int(round(value * d))
        sx = self.get_or_add_sx()
        sy = self.get_or_add_sy()
        sx.n = n
        sx.d = d
        sy.n = n
        sy.d = d


class CT_Ratio(BaseOxmlElement):
    """`a:sx` / `a:sy` element — numerator/denominator ratio.

    Maps ECMA-376 ``CT_Ratio`` (``xsd:long n`` / ``xsd:long d``).
    """

    n: int = RequiredAttribute("n", XsdLong)  # pyright: ignore[reportAssignmentType]
    d: int = RequiredAttribute("d", XsdLong)  # pyright: ignore[reportAssignmentType]
