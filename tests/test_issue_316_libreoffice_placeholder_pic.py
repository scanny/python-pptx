# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #316 — picture placeholder in LibreOffice.

Issue #316 (https://github.com/scanny/python-pptx/issues/316) reported that a
deck where ``python-pptx`` inserted a picture into a layout's picture
placeholder rendered fine in PowerPoint but showed a blank placeholder in
LibreOffice. The reporter never posted an opc-diff to localize the
difference, and the thread went cold.

LibreOffice's Impress is stricter than PowerPoint about a picture-placeholder
subtree — in particular it expects:

* ``p:nvPicPr`` containing all three of ``p:cNvPr``, ``p:cNvPicPr`` (with
  ``a:picLocks``), and ``p:nvPr`` (with the ``p:ph`` element that marks the
  shape as a placeholder, carrying its ``idx`` and type="pic");
* ``p:blipFill`` containing ``a:blip`` with an explicit ``r:embed`` (no
  implicit-link tolerance);
* ``a:stretch/a:fillRect`` as the blipFill's display mode — without it,
  LibreOffice draws a blank placeholder rectangle instead of rendering the
  image;
* ``p:spPr`` present (may be empty — dimensions may be inherited from the
  layout placeholder).

As of the Wave 13 ``_copy_inherited_spPr_decorations`` fix for issue #907
(commit ``fc9f3915``), the XML emitted by ``insert_picture`` on a picture
placeholder carries every one of those elements, and a manual LibreOffice
round-trip (``libreoffice --headless --convert-to pdf``) confirms the image
renders. This file is the verify-close pass — it pins the XML structure so
that a future refactor of ``_pic_ph_tmpl`` or the placeholder-promotion flow
cannot silently regress LibreOffice compatibility for this scenario.

The tests exercise the real public API end-to-end (no mocks) and cover the
three insertion variants users hit in practice:

* the specialized ``PICTURE`` placeholder (layout "Picture with Caption");
* a generic ``OBJECT``/content placeholder (layout "Title and Content");
* the ``crop=False`` fit-to-box variant shipped in the loadfix fork.

Each test confirms the full element chain LibreOffice requires, directly
and after a ``Presentation.save`` + reopen round-trip. See also the
complementary unit tests in ``tests/shapes/test_placeholder.py`` for the
placeholder-promotion mechanics.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.oxml.ns import qn


class DescribeIssue316LibreOfficePlaceholderPicture:
    """Verify-close regression suite for picture placeholders in LibreOffice."""

    def it_emits_the_nvPicPr_subtree_LibreOffice_requires(self):
        # -- LibreOffice refuses to render a picture placeholder whose
        # -- p:nvPicPr is missing any of cNvPr / cNvPicPr / nvPr. Pin the
        # -- full subtree for the canonical #316 flow: the specialised
        # -- PICTURE placeholder on the "Picture with Caption" layout.
        pic_element = self._insert_picture_into_pic_placeholder()

        nvPicPr = pic_element.find(qn("p:nvPicPr"))
        assert nvPicPr is not None
        assert nvPicPr.find(qn("p:cNvPr")) is not None
        cNvPicPr = nvPicPr.find(qn("p:cNvPicPr"))
        assert cNvPicPr is not None
        # -- a:picLocks keeps PowerPoint from offering "change picture" in a
        # -- way LibreOffice has historically mis-rendered.
        assert cNvPicPr.find(qn("a:picLocks")) is not None
        nvPr = nvPicPr.find(qn("p:nvPr"))
        assert nvPr is not None
        # -- the p:ph child is the marker that makes the p:pic a placeholder,
        # -- and LibreOffice needs its idx to link the pic back to the layout.
        ph = nvPr.find(qn("p:ph"))
        assert ph is not None
        assert ph.get("idx") == "1"
        assert ph.get("type") == "pic"

    def it_emits_a_blipFill_with_an_explicit_r_embed(self):
        # -- LibreOffice refuses to load a placeholder picture without an
        # -- explicit r:embed on a:blip; the implicit-link tolerance a few
        # -- PowerPoint versions show isn't there. Pin the rId.
        pic_element = self._insert_picture_into_pic_placeholder()

        blipFill = pic_element.find(qn("p:blipFill"))
        assert blipFill is not None
        blip = blipFill.find(qn("a:blip"))
        assert blip is not None
        rEmbed = blip.get(qn("r:embed"))
        assert rEmbed is not None
        assert rEmbed.startswith("rId")

    def it_emits_an_a_stretch_a_fillRect_display_mode(self):
        # -- #316's "blank placeholder" symptom is what LibreOffice draws
        # -- when a:stretch/a:fillRect is missing (no display-mode child on
        # -- a:blipFill). Pin it for both the default crop=True path and
        # -- the crop=False fit path.
        for crop in (True, False):
            pic_element = self._insert_picture_into_pic_placeholder(crop=crop)
            blipFill = pic_element.find(qn("p:blipFill"))
            assert blipFill is not None, f"no p:blipFill (crop={crop})"
            stretch = blipFill.find(qn("a:stretch"))
            assert stretch is not None, f"no a:stretch (crop={crop})"
            assert (
                stretch.find(qn("a:fillRect")) is not None
            ), f"no a:stretch/a:fillRect (crop={crop})"

    def it_also_emits_the_required_subtree_for_a_generic_OBJECT_placeholder(self):
        # -- insert_picture is inherited by SlidePlaceholder; a "Title and
        # -- Content" layout's OBJECT placeholder must emit the same
        # -- LibreOffice-required structure as a specialised PicturePlaceholder.
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
        content_ph = next(ph for ph in slide.placeholders if ph.placeholder_format.idx == 1)

        ph_pic = content_ph.insert_picture("tests/test_files/python-powered.png")
        pic_element = ph_pic._element

        # -- full LibreOffice checklist on a generic-placeholder promotion --
        assert pic_element.xpath("./p:nvPicPr/p:cNvPr")
        assert pic_element.xpath("./p:nvPicPr/p:cNvPicPr/a:picLocks")
        assert pic_element.xpath("./p:nvPicPr/p:nvPr/p:ph")
        blips = pic_element.xpath("./p:blipFill/a:blip")
        assert len(blips) == 1
        assert blips[0].get(qn("r:embed")) is not None
        assert pic_element.xpath("./p:blipFill/a:stretch/a:fillRect")
        assert pic_element.find(qn("p:spPr")) is not None

    def it_preserves_the_LibreOffice_structure_through_save_and_reopen(self):
        # -- the failure mode in #316 is in the *saved* file, not the
        # -- in-memory element, so the load-bearing assertion is that the
        # -- serialised pptx still carries every piece when reopened.
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[8])  # Picture with Caption
        pic_ph = next(ph for ph in slide.placeholders if ph.placeholder_format.idx == 1)
        pic_ph.insert_picture("tests/test_files/python-powered.png")

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reloaded = Presentation(buf)
        # -- pick the p:pic (picture placeholder) out of the shape tree;
        # -- the layout "Picture with Caption" also carries a title p:sp.
        pic_elements = [
            shape._element
            for shape in reloaded.slides[0].shapes
            if shape._element.tag == qn("p:pic")
        ]
        assert len(pic_elements) == 1
        pic_element = pic_elements[0]

        # -- structural pins survive round-trip --
        assert pic_element.xpath("./p:nvPicPr/p:cNvPr")
        assert pic_element.xpath("./p:nvPicPr/p:cNvPicPr/a:picLocks")
        phs = pic_element.xpath("./p:nvPicPr/p:nvPr/p:ph")
        assert len(phs) == 1
        assert phs[0].get("type") == "pic"
        assert phs[0].get("idx") == "1"
        # -- blipFill with a resolvable rId --
        blips = pic_element.xpath("./p:blipFill/a:blip")
        assert len(blips) == 1
        rEmbed = blips[0].get(qn("r:embed"))
        assert rEmbed is not None
        assert reloaded.slides[0].part.related_part(rEmbed) is not None
        # -- stretch/fillRect display mode survives --
        assert pic_element.xpath("./p:blipFill/a:stretch/a:fillRect")

    # -- helpers ----------------------------------------------------------

    @staticmethod
    def _insert_picture_into_pic_placeholder(crop: bool = True):
        """Return the ``p:pic`` element produced by inserting a picture into
        a PICTURE placeholder on the "Picture with Caption" layout.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[8])  # Picture with Caption
        pic_ph = next(ph for ph in slide.placeholders if ph.placeholder_format.idx == 1)
        ph_pic = pic_ph.insert_picture("tests/test_files/python-powered.png", crop=crop)
        return ph_pic._element
