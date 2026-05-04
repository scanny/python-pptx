# pyright: reportPrivateUsage=false

"""Regression test for issue #537 — "font.color and font.fill seem to do the same thing".

Issue #537 (https://github.com/scanny/python-pptx/issues/537) reports confusion
between :attr:`.Font.color` and :attr:`.Font.fill.fore_color`: both exist on
the same |Font| object and, for a run with a solid fill, both appear to drive
the same colour. The reporter asked what the difference is.

The distinction is intentional and narrow:

  * :attr:`.Font.color` is a *shortcut* for :attr:`.Font.fill.fore_color` when
    the run's fill is (or is being promoted to) a solid ``<a:solidFill>``.
    It is the quickest way to write a solid-colour run — the common case.
  * :attr:`.Font.fill` returns the full |FillFormat| proxy, which also
    supports :meth:`.FillFormat.gradient`, :meth:`.FillFormat.patterned`,
    the pattern :attr:`.FillFormat.back_color`, and future fill kinds.
    It is the full-surface API the library exposes on top of
    ``a:rPr/EG_FillProperties``.

So the two are *not* interchangeable. They overlap only on the solid-fill
happy path, where they point at the same ``a:solidFill/a:srgbClr`` element —
i.e. they are two read/write views onto the same XML. Everywhere else the
shortcut narrows the surface to just the run's foreground colour.

This suite pins that contract from both directions:

  * when the run's fill is solid, the two paths read and write the *same*
    XML — ``font.color.rgb = ...`` and ``font.fill.fore_color.rgb = ...``
    are equivalent, and both surface the same `<a:srgbClr>` element;
  * ``font.fill`` exposes capabilities (``gradient``, ``patterned``,
    ``back_color``, ``gradient_stops``, …) that ``font.color`` does not —
    the shortcut is deliberately *narrow*;
  * ``font.fill.type`` reports ``MSO_FILL.SOLID`` / ``MSO_FILL.GRADIENT`` /
    ``MSO_FILL.PATTERNED`` according to which fill the caller selected, so
    callers can branch on "is this run actually a solid colour?" before
    reaching for ``font.color``.

#537 is closed with documentation (``docs/user/text.rst`` → "Font color vs.
font fill"), not with a code change. A future regression that silently made
``font.color`` *not* a shortcut for ``font.fill.fore_color`` — or that widened
``font.color`` into a full fill surface — would reproduce the original
confusion and be caught here.
"""

from __future__ import annotations

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL, MSO_PATTERN_TYPE, MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.util import Inches


@pytest.fixture
def _run_font():
    """Yield a fresh :class:`~pptx.text.text.Font` on a textbox run.

    Textbox runs carry an untouched ``a:rPr`` with no fill element, which
    is the state where the :attr:`.Font.color` shortcut's deferred-promotion
    path (``_FontColorFormat``) kicks in — i.e. the setup that most
    reliably exercises the shortcut vs. full-surface contract this suite
    pins.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    paragraph = tb.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "hi"
    return run.font


class DescribeIssue537FontFillVsColorVerify:
    """#537 — font.color is a shortcut for font.fill.fore_color on solid fills."""

    # -- equivalence on the solid-fill happy path --------------------------

    def it_writes_the_same_xml_whether_color_or_fill_is_used(self):
        """Both authoring paths produce byte-identical ``a:solidFill`` XML.

        The #537 reporter's concern is "these look like two ways to do the
        same thing" — that's true on the solid-fill path and this test pins
        it explicitly. Two sibling runs, one coloured via ``font.color.rgb``
        and the other via ``font.fill.solid()`` + ``fill.fore_color.rgb``,
        must end up with the same ``a:rPr/a:solidFill/a:srgbClr`` structure.
        A regression that made the two paths diverge in what XML they write
        would reintroduce the confusion #537 names.
        """
        from lxml import etree  # pyright: ignore[reportMissingTypeStubs]

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        paragraph = tb.text_frame.paragraphs[0]

        run_a = paragraph.add_run()
        run_a.text = "A"
        run_a.font.fill.solid()
        run_a.font.fill.fore_color.rgb = RGBColor(0xAB, 0xCD, 0xEF)

        run_b = paragraph.add_run()
        run_b.text = "B"
        run_b.font.color.rgb = RGBColor(0xAB, 0xCD, 0xEF)

        # -- strip xmlns attrs that lxml repeats on subtree serialization
        # -- so the comparison is structural, not namespace-declaration noise.
        xml_a = etree.tostring(run_a.font._rPr.find(qn("a:solidFill"))).decode()
        xml_b = etree.tostring(run_b.font._rPr.find(qn("a:solidFill"))).decode()
        assert xml_a == xml_b

    def it_reads_the_same_rgb_value_through_both_paths(self, _run_font):
        """A value written via one path reads back through the other.

        If ``font.fill.fore_color`` and ``font.color`` are the same proxy
        on a solid fill, a write via either must be visible through both.
        """
        font = _run_font
        font.fill.solid()
        font.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)

        assert font.color.rgb == RGBColor(0x12, 0x34, 0x56)
        assert font.fill.fore_color.rgb == RGBColor(0x12, 0x34, 0x56)
        assert font.color.type == font.fill.fore_color.type

    def it_surfaces_the_same_srgbClr_element_through_both_paths(self, _run_font):
        """The two proxies point at the *same* ``a:srgbClr`` XML element.

        Not just equal values — the identity of the underlying element is
        the same. A copy-on-read regression that cloned the solidFill
        subtree would defeat the shortcut's whole purpose.
        """
        font = _run_font
        font.fill.solid()
        font.fill.fore_color.rgb = RGBColor(0xDE, 0xAD, 0xBE)

        # -- locate the single a:srgbClr the run carries --
        srgbClr_nodes = font._rPr.findall(".//" + qn("a:srgbClr"))
        assert len(srgbClr_nodes) == 1
        srgbClr = srgbClr_nodes[0]

        # -- both read proxies bottom out at that same element. `_xClr` is
        # -- the private handle on the color object inside each ColorFormat.
        assert font.color._color._xClr is srgbClr
        assert font.fill.fore_color._color._xClr is srgbClr

    def it_shares_the_colorformat_proxy_when_fill_is_already_solid(self, _run_font):
        """Shortcut contract: on a solid fill, ``font.color is font.fill.fore_color``.

        ``Font.color`` delegates to ``fill.fore_color`` when the fill is
        already solid. ``fore_color`` on ``_SolidFill`` is a
        ``@lazyproperty``, so repeated access returns the same
        |ColorFormat| instance — and ``Font.color`` is cached on first
        access too. The net effect is that the shortcut returns the very
        same object, not a fresh proxy.
        """
        font = _run_font
        font.fill.solid()
        assert font.color is font.fill.fore_color

    # -- fill-only capabilities are NOT accessible via font.color ---------

    def it_exposes_gradient_on_font_fill_but_not_on_font_color(self, _run_font):
        """``font.fill.gradient()`` exists; ``font.color.gradient()`` doesn't.

        The #537 reporter asks "what's the difference?" The answer is that
        ``font.fill`` is the full |FillFormat|. ``font.color`` is a
        narrower |ColorFormat| proxy — it has no ``gradient`` method.
        """
        font = _run_font

        # -- the fill surface supports gradient --
        assert hasattr(font.fill, "gradient")
        font.fill.gradient()
        assert font.fill.type == MSO_FILL.GRADIENT

        # -- the color shortcut does not --
        assert not hasattr(font.color, "gradient")

    def it_exposes_patterned_on_font_fill_but_not_on_font_color(self, _run_font):
        """``font.fill.patterned()`` exists; ``font.color.patterned()`` doesn't.

        Same contract as gradient: the full |FillFormat| can switch the run
        to a pattern fill; the shortcut |ColorFormat| cannot.
        """
        font = _run_font

        assert hasattr(font.fill, "patterned")
        font.fill.patterned()
        assert font.fill.type == MSO_FILL.PATTERNED

        assert not hasattr(font.color, "patterned")

    def it_exposes_back_color_on_font_fill_but_not_on_font_color(self, _run_font):
        """Pattern fills need both a fore and a back colour.

        Only :class:`.FillFormat` exposes :attr:`~.FillFormat.back_color`.
        The :attr:`.Font.color` shortcut is a foreground-colour proxy and
        has no back-colour surface, which is exactly the "narrower" part
        of the shortcut-vs-full-surface distinction the docs clarify.
        """
        font = _run_font
        font.fill.patterned()
        font.fill.pattern = MSO_PATTERN_TYPE.HORIZONTAL_BRICK

        assert hasattr(font.fill, "back_color")
        # -- back_color writes land on a:bgClr --
        font.fill.back_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        bgClr = font._rPr.find(qn("a:pattFill") + "/" + qn("a:bgClr"))
        assert bgClr is not None

        assert not hasattr(font.color, "back_color")

    def it_exposes_gradient_stops_on_font_fill_but_not_on_font_color(self, _run_font):
        """A gradient is defined by its stops; only ``font.fill`` exposes them.

        Demonstrates that branching to the gradient surface *requires*
        going through ``font.fill`` — the shortcut has nowhere to hang
        stops off.
        """
        font = _run_font
        font.fill.gradient()

        assert hasattr(font.fill, "gradient_stops")
        # -- default gradient has exactly two stops (see FillFormat.gradient docs) --
        assert len(font.fill.gradient_stops) == 2

        assert not hasattr(font.color, "gradient_stops")

    # -- fill.type reports the actual fill flavour -------------------------

    def it_reports_MSO_FILL_SOLID_after_font_fill_solid(self, _run_font):
        """``font.fill.type`` is the canonical "is this a solid fill?" probe.

        Callers who need to distinguish the solid-fill happy path
        (``font.color`` usable) from other fills (gradient / pattern /
        picture / background) should branch on ``font.fill.type``.
        """
        font = _run_font
        font.fill.solid()
        assert font.fill.type == MSO_FILL.SOLID

    def it_reports_MSO_FILL_GRADIENT_after_font_fill_gradient(self, _run_font):
        font = _run_font
        font.fill.gradient()
        assert font.fill.type == MSO_FILL.GRADIENT

    def it_reports_MSO_FILL_PATTERNED_after_font_fill_patterned(self, _run_font):
        font = _run_font
        font.fill.patterned()
        assert font.fill.type == MSO_FILL.PATTERNED

    def it_reports_MSO_FILL_BACKGROUND_after_font_fill_background(self, _run_font):
        """The "no fill" (transparent) state also surfaces through fill.type.

        Completes the branch table a caller needs to distinguish the
        solid-fill happy path from every other state.
        """
        font = _run_font
        font.fill.background()
        assert font.fill.type == MSO_FILL.BACKGROUND

    # -- color-type reads on non-solid fills -------------------------------

    def it_reports_color_type_None_when_fill_is_gradient(self, _run_font):
        """On a non-solid fill the shortcut reports "no colour set".

        The ``Font.color`` shortcut only binds to ``a:solidFill`` — when
        the fill is a gradient or pattern, ``color.type`` is |None| (the
        "deferred" state from issue #1111's ``_FontColorFormat``). Callers
        who rely on that to detect "run is not a flat solid colour" would
        be broken by a regression that started reporting some other type.
        """
        font = _run_font
        font.fill.gradient()
        assert font.color.type is None

    def it_reports_color_type_None_when_fill_is_patterned(self, _run_font):
        """Same no-colour-shortcut story for pattern fills.

        Patterned fills *do* have a fore colour; it just isn't the
        ``font.color`` shortcut — you have to go via
        ``font.fill.fore_color``.
        """
        font = _run_font
        font.fill.patterned()
        font.fill.fore_color.rgb = RGBColor(0xAA, 0xBB, 0xCC)

        # -- the shortcut refuses to bind --
        assert font.color.type is None
        # -- but the full-fill surface still has the pattern's fore colour --
        assert font.fill.fore_color.rgb == RGBColor(0xAA, 0xBB, 0xCC)

    # -- theme-color path mirrors the same shortcut contract ---------------

    def it_routes_theme_color_writes_through_the_same_solidFill(self, _run_font):
        """``font.color.theme_color`` is the same shortcut with a scheme colour.

        Tested in addition to the RGB variant because the
        ``_FontColorFormat`` deferred-promotion path has a separate setter
        for ``theme_color``; a regression there would reproduce the
        original #1111 "inherited colour silently promoted" symptom the
        shortcut is designed to avoid.
        """
        font = _run_font
        font.fill.solid()
        font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

        schemeClr = font._rPr.find(qn("a:solidFill") + "/" + qn("a:schemeClr"))
        assert schemeClr is not None
        assert schemeClr.get("val") == "accent1"

        # -- and fill.fore_color sees the same theme colour --
        assert font.fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1
