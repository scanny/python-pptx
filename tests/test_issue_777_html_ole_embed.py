# pyright: reportPrivateUsage=false

"""Regression test for issue #777 — embed an HTML file as an OLE object.

Issue #777 (https://github.com/scanny/python-pptx/issues/777) asked for
the ability to embed an HTML file as an OLE object inside a slide, so
that PowerPoint displays an icon that launches the default HTML handler
when double-clicked.

The generic OLE-object embedding path delivered by
``feat/issue-752-ole-embed-generic`` (merged on master at
``270972c3``) already accepts an arbitrary ``progId`` string plus a
caller-supplied ``extension`` and wires the bytes through to a generic
``EmbeddedPackagePart`` carrying the ``OFC_OLE_OBJECT`` content-type.
``MSHtml.MHT`` is the progId PowerPoint itself uses for embedded HTML
documents, so calling ``SlideShapes.add_ole_object(html_path,
prog_id="MSHtml.MHT", ..., extension="html")`` is sufficient today.

This test exercises the #777 flow end-to-end to pin the behavior:

* author the embed from a str path to an on-disk ``.html`` file,
* author the embed from a ``BytesIO`` of the same bytes,
* round-trip the package through save + reopen,
* verify the ``_OleFormat`` proxy reports the ``MSHtml.MHT`` progId and
  returns the original HTML bytes byte-for-byte, and
* verify the package contains an ``/ppt/embeddings/oleObject*.html``
  part with the generic OLE content-type.
"""

from __future__ import annotations

import io
import os
import tempfile
import zipfile

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.util import Inches


_HTML_BYTES = (
    b"<!DOCTYPE html>\n"
    b"<html><head><title>#777</title></head>"
    b"<body><p>Hello, embedded HTML!</p></body></html>\n"
)


class DescribeIssue777EmbedHtmlOleObject(object):
    """The #777 flow: embed an HTML file as an OLE object (via ``MSHtml.MHT``)."""

    def it_round_trips_an_html_file_path_as_an_MSHtml_OLE_object(
        self, tmp_path
    ):
        """``add_ole_object(html_path, prog_id="MSHtml.MHT", ..., extension="html")``
        embeds the file and survives save / reopen byte-for-byte.
        """
        html_path = tmp_path / "page.html"
        html_path.write_bytes(_HTML_BYTES)

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        shape = slide.shapes.add_ole_object(
            str(html_path),
            prog_id="MSHtml.MHT",
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
        )

        # -- authored shape exposes the OLE metadata up front --
        assert shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        assert shape.ole_format.prog_id == "MSHtml.MHT"
        assert shape.ole_format.blob == _HTML_BYTES
        assert shape.ole_format.show_as_icon is True

        # -- save + reopen round-trip preserves progId and bytes --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]
        ole_shapes = [s for s in slide2.shapes if s.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT]
        assert len(ole_shapes) == 1
        reopened = ole_shapes[0]
        assert reopened.ole_format.prog_id == "MSHtml.MHT"
        assert reopened.ole_format.blob == _HTML_BYTES
        assert reopened.ole_format.show_as_icon is True

    def it_also_accepts_a_file_like_object_with_explicit_extension(
        self
    ):
        """Callers that only hold HTML bytes in memory can still embed
        via ``BytesIO``, provided they pass ``extension="html"`` so the
        embedded part-name carries a meaningful suffix.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        shape = slide.shapes.add_ole_object(
            io.BytesIO(_HTML_BYTES),
            prog_id="MSHtml.MHT",
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
            extension="html",
        )

        assert shape.ole_format.prog_id == "MSHtml.MHT"
        assert shape.ole_format.blob == _HTML_BYTES

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        # -- package contains a generic OLE embedding part with an
        #    .html suffix and the OLE content-type. This is the #1070
        #    content-type whitelist working in concert with #752's
        #    generic embedding path. --
        with zipfile.ZipFile(buf) as zf:
            names = zf.namelist()
            ole_names = [n for n in names if n.startswith("ppt/embeddings/oleObject") and n.endswith(".html")]
            assert len(ole_names) == 1
            # -- [Content_Types].xml declares the OLE content-type for
            #    the .html part, not an HTML content-type --
            content_types_xml = zf.read("[Content_Types].xml").decode("utf-8")
            assert CT.OFC_OLE_OBJECT in content_types_xml
            # -- and the embedded bytes match exactly --
            assert zf.read(ole_names[0]) == _HTML_BYTES

    def it_accepts_the_PROG_ID_HTML_enum_convenience_member(self):
        """The ``PROG_ID.HTML`` convenience member embeds the same way
        via progId ``"htmlfile"`` (the other PowerPoint-emitted HTML
        progId). The enum-member variant is tested here alongside the
        #777-named ``"MSHtml.MHT"`` progId to pin both working paths.
        """
        from pptx.enum.shapes import PROG_ID

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as f:
            f.write(_HTML_BYTES)
            html_path = f.name
        try:
            shape = slide.shapes.add_ole_object(
                html_path,
                prog_id=PROG_ID.HTML,
                left=Inches(1),
                top=Inches(1),
            )
        finally:
            os.unlink(html_path)

        assert shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        assert shape.ole_format.prog_id == "htmlfile"
        assert shape.ole_format.blob == _HTML_BYTES

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]
        ole_shapes = [s for s in slide2.shapes if s.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT]
        assert len(ole_shapes) == 1
        assert ole_shapes[0].ole_format.prog_id == "htmlfile"
        assert ole_shapes[0].ole_format.blob == _HTML_BYTES
