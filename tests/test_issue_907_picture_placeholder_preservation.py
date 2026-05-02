# pyright: reportPrivateUsage=false

"""Regression test for issue #907 — picture-placeholder formatting lost on insert_picture.

Issue #907 (https://github.com/scanny/python-pptx/issues/907) reports that
when a slide layout defines a picture placeholder whose ``p:spPr`` carries
styling decorations (``a:prstGeom`` + ``a:ln`` outline + ``a:effectLst``
soft-edges, etc.), calling :meth:`PicturePlaceholder.insert_picture` on the
slide-level placeholder replaces the ``p:sp`` with a ``p:pic`` but emits a
completely empty ``p:spPr``. PowerPoint renders the inserted picture with no
outline and no soft edges even though the layout clearly intended them.

The fix copies inheritable styling children (``a:prstGeom`` adjustments,
``a:ln``, ``a:effectLst``) from the slide placeholder's own ``p:spPr``
(when present) or from the layout placeholder's ``p:spPr`` (when the slide
placeholder's own spPr does not define them) into the new ``p:pic/p:spPr``.
Fill-related children (``a:blipFill``, ``a:solidFill``, ...) and ``a:xfrm``
are deliberately NOT copied — the picture supplies its own fill, and the
transform is either inherited from the layout (default ``crop=True``) or
explicitly emitted by ``_fit_pic_to_placeholder`` (``crop=False``).
"""

from __future__ import annotations

import io

from lxml import etree
from PIL import Image

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.shapes.placeholder import PlaceholderPicture


def _tiny_png_bytes() -> io.BytesIO:
    """Return a BytesIO holding a 10x10 red PNG, suitable for insert_picture."""
    buf = io.BytesIO()
    Image.new("RGB", (10, 10), (255, 0, 0)).save(buf, format="PNG")
    buf.seek(0)
    return buf


def _decorate_layout_picture_ph_spPr(layout) -> None:
    """Inject a red ``a:ln`` and a softEdge ``a:effectLst`` onto the layout PICTURE ph."""
    for ph in layout.placeholders:
        if ph.placeholder_format.type != PP_PLACEHOLDER.PICTURE:
            continue
        spPr = ph._element.spPr
        ln_xml = (
            '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' w="38100">'
            '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
            "</a:ln>"
        )
        effectLst_xml = (
            '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:softEdge rad="63500"/>'
            "</a:effectLst>"
        )
        spPr.append(etree.fromstring(ln_xml))
        spPr.append(etree.fromstring(effectLst_xml))
        return
    raise AssertionError("no PICTURE placeholder on layout")


def _find_picture_ph(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            return ph
    raise AssertionError("no PICTURE placeholder on slide")


class DescribeIssue907PicturePlaceholderPreservation(object):
    """Verify #907: insert_picture preserves inherited spPr decorations."""

    def it_copies_layout_ln_and_effectLst_into_the_new_pic_spPr(self):
        """When the layout's picture-ph has ``a:ln`` + ``a:effectLst`` the
        inserted picture's ``p:pic/p:spPr`` carries equivalent children.
        """
        prs = Presentation()
        # -- layout 8 is the built-in "picture with caption" layout --
        layout = prs.slide_layouts[8]
        _decorate_layout_picture_ph_spPr(layout)

        slide = prs.slides.add_slide(layout)
        picture_ph = _find_picture_ph(slide)

        placeholder_pic = picture_ph.insert_picture(_tiny_png_bytes())

        assert isinstance(placeholder_pic, PlaceholderPicture)
        pic_spPr = placeholder_pic._element.spPr

        ln = pic_spPr.find(qn("a:ln"))
        assert ln is not None, "a:ln outline was dropped during insert_picture"
        assert ln.get("w") == "38100"
        srgb = ln.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
        assert srgb is not None
        assert srgb.get("val") == "FF0000"

        effectLst = pic_spPr.find(qn("a:effectLst"))
        assert effectLst is not None, "a:effectLst (soft edge) was dropped"
        softEdge = effectLst.find(qn("a:softEdge"))
        assert softEdge is not None
        assert softEdge.get("rad") == "63500"

    def it_preserves_those_decorations_through_a_save_and_reopen_cycle(self):
        """The spPr decorations must survive a round-trip through a ZIP save/reopen."""
        prs = Presentation()
        layout = prs.slide_layouts[8]
        _decorate_layout_picture_ph_spPr(layout)

        slide = prs.slides.add_slide(layout)
        picture_ph = _find_picture_ph(slide)
        picture_ph.insert_picture(_tiny_png_bytes())

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        reopened = Presentation(buf)
        # -- the placeholder is now a PlaceholderPicture on the reopened slide --
        ph_pic = None
        for shape in reopened.slides[0].placeholders:
            if shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                ph_pic = shape
                break
        assert ph_pic is not None
        assert isinstance(ph_pic, PlaceholderPicture)

        spPr = ph_pic._element.spPr
        ln = spPr.find(qn("a:ln"))
        assert ln is not None
        assert ln.get("w") == "38100"
        effectLst = spPr.find(qn("a:effectLst"))
        assert effectLst is not None
        softEdge = effectLst.find(qn("a:softEdge"))
        assert softEdge is not None
        assert softEdge.get("rad") == "63500"

        # -- and the picture part is referenced (i.e. the image survived too) --
        blipFill = ph_pic._element.find(qn("p:blipFill"))
        assert blipFill is not None
        blip = blipFill.find(qn("a:blip"))
        assert blip is not None
        assert blip.get(qn("r:embed")) is not None

    def it_does_not_copy_a_blipFill_from_the_layout(self):
        """Regression guard: an ``a:blipFill`` (or other fill) on the layout
        placeholder must NOT be copied into the inserted picture — the
        picture provides its own fill and we don't want the layout's sample
        image to override the newly inserted one.
        """
        prs = Presentation()
        layout = prs.slide_layouts[8]
        # -- stash a bogus blipFill on the layout picture placeholder spPr --
        for ph in layout.placeholders:
            if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                spPr = ph._element.spPr
                spPr.append(
                    etree.fromstring(
                        '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/'
                        'drawingml/2006/main"><a:srgbClr val="00FF00"/></a:solidFill>'
                    )
                )
                break

        slide = prs.slides.add_slide(layout)
        picture_ph = _find_picture_ph(slide)
        placeholder_pic = picture_ph.insert_picture(_tiny_png_bytes())

        pic_spPr = placeholder_pic._element.spPr
        # -- the solidFill must NOT have been copied into the pic's spPr --
        assert pic_spPr.find(qn("a:solidFill")) is None

    def it_prefers_slide_spPr_decorations_over_layout_spPr_decorations(self):
        """If the slide-level placeholder overrides the layout's decorations,
        the slide's own values win (direct > inherited).
        """
        prs = Presentation()
        layout = prs.slide_layouts[8]
        _decorate_layout_picture_ph_spPr(layout)  # layout ln w=38100

        slide = prs.slides.add_slide(layout)
        picture_ph = _find_picture_ph(slide)
        # -- override the outline thickness on the slide-level placeholder --
        slide_spPr = picture_ph._element.spPr
        slide_spPr.append(
            etree.fromstring(
                '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                ' w="76200">'
                '<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>'
                "</a:ln>"
            )
        )

        placeholder_pic = picture_ph.insert_picture(_tiny_png_bytes())
        ln = placeholder_pic._element.spPr.find(qn("a:ln"))
        assert ln is not None
        # -- slide value beats layout value --
        assert ln.get("w") == "76200"
        srgb = ln.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
        assert srgb is not None
        assert srgb.get("val") == "0000FF"
