# pyright: reportPrivateUsage=false

"""Regression test for issue #1047 — box-and-whisker chartex passthrough.

Issue #1047 (https://github.com/scanny/python-pptx/issues/1047) asks for
support of PowerPoint's Office 2016+ *box-and-whisker* chart
(``cx:boxAndWhiskerChart``/``cx:series/@layoutId="boxWhisker"``). Box-and-
whisker is one of seven chart kinds that live in the Microsoft
``cx:`` (chartex) namespace rather than the legacy ``c:`` grammar;
structured authoring of chartex charts is still blocked on the F4 chartex
foundation (see ``docs/dev/analysis/chartex-foundation.rst``) and the per-
kind design note ``docs/dev/analysis/chartex-box-whisker.rst``.

What this test exercises is the **passthrough MVP** available today:

* :attr:`GraphicFrame.has_chart` / :attr:`GraphicFrame.has_chartex` are
  |True| for a box-and-whisker graphic-frame, with
  :attr:`GraphicFrame.chart_type` reporting
  :attr:`XL_CHART_TYPE.UNSUPPORTED_CHARTEX`.
* :attr:`GraphicFrame.chartex_type` returns the raw
  ``cx:series/@layoutId`` value — ``"boxWhisker"`` — so downstream
  consumers can distinguish the chart kind without parsing the chartex
  part themselves.
* The ``cx:chartSpace`` part and its declared ``layoutId`` survive
  save / reopen round-trip verbatim.

The fixture is a minimal in-process ``.pptx`` built by zip surgery on a
fresh :class:`~pptx.Presentation`: a single slide with one
``mc:AlternateContent``-wrapped chartex graphic-frame referencing a
hand-authored box-and-whisker ``cx:chartSpace``.
"""

from __future__ import annotations

import io
import zipfile

from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.shapes.graphfrm import GraphicFrame

_CHARTEX_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex"
_CHARTEX_CT = "application/vnd.ms-office.chartex+xml"
_CHARTEX_RELTYPE = "http://schemas.microsoft.com/office/2014/relationships/chartEx"

# -- a minimal box-and-whisker cx:chartSpace. The distinguishing feature
# -- is `cx:series/@layoutId="boxWhisker"`. Two numeric dimensions are
# -- used (one "cat"-role numDim for group labels, one "val"-role numDim
# -- for observations) which is how PowerPoint writes a box-and-whisker
# -- with a single grouped series. A companion `cx:layoutPr/cx:statistics`
# -- block selects quartile method + outlier/mean-marker visibility; it
# -- is deliberately retained verbatim here so the round-trip assertion
# -- can confirm non-MVP sub-elements survive as opaque payload. --
_BOX_WHISKER_CX_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    "<cx:chartData>"
    '<cx:data id="0">'
    '<cx:strDim type="cat">'
    "<cx:f>Sheet1!$A$2:$A$5</cx:f>"
    '<cx:lvl ptCount="4">'
    '<cx:pt idx="0">A</cx:pt><cx:pt idx="1">A</cx:pt>'
    '<cx:pt idx="2">B</cx:pt><cx:pt idx="3">B</cx:pt>'
    "</cx:lvl>"
    "</cx:strDim>"
    '<cx:numDim type="val">'
    "<cx:f>Sheet1!$B$2:$B$5</cx:f>"
    '<cx:lvl ptCount="4" formatCode="General">'
    '<cx:pt idx="0">10</cx:pt><cx:pt idx="1">12</cx:pt>'
    '<cx:pt idx="2">7</cx:pt><cx:pt idx="3">9</cx:pt>'
    "</cx:lvl>"
    "</cx:numDim>"
    "</cx:data>"
    "</cx:chartData>"
    "<cx:chart>"
    "<cx:plotArea>"
    "<cx:plotAreaRegion>"
    "<cx:plotSurface/>"
    '<cx:series layoutId="boxWhisker" hidden="0" ownerIdx="0">'
    "<cx:tx><cx:txData><cx:v>Values</cx:v></cx:txData></cx:tx>"
    "<cx:layoutPr>"
    '<cx:statistics quartileMethod="exclusive"/>'
    '<cx:visibility meanMarker="1" meanLine="0" nonoutliers="0" outliers="1"/>'
    "</cx:layoutPr>"
    '<cx:dataId val="0"/>'
    "</cx:series>"
    "</cx:plotAreaRegion>"
    '<cx:axis id="0"><cx:catScaling gapWidth="0.5"/></cx:axis>'
    '<cx:axis id="1"><cx:valScaling/></cx:axis>'
    "</cx:plotArea>"
    "</cx:chart>"
    "</cx:chartSpace>"
)

# -- slide1.xml wrapping the chartex graphic-frame in `mc:AlternateContent`
# -- with a minimal `mc:Fallback` `p:sp`. The `cx:chart r:id="rIdCx"`
# -- references the chartEx part via the slide's rels file. --
_SLIDE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">'
    "<p:cSld><p:spTree>"
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    "<p:grpSpPr/>"
    "<mc:AlternateContent>"
    '<mc:Choice xmlns:cx14="http://schemas.microsoft.com/office/drawing/2014/chartex"'
    ' Requires="cx14">'
    "<p:graphicFrame>"
    "<p:nvGraphicFramePr>"
    '<p:cNvPr id="2" name="Chart 1"/>'
    "<p:cNvGraphicFramePr/><p:nvPr/>"
    "</p:nvGraphicFramePr>"
    '<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm>'
    "<a:graphic>"
    f'<a:graphicData uri="{_CHARTEX_URI}">'
    '<cx:chart r:id="rIdCx"/>'
    "</a:graphicData>"
    "</a:graphic>"
    "</p:graphicFrame>"
    "</mc:Choice>"
    "<mc:Fallback>"
    "<p:sp>"
    '<p:nvSpPr><p:cNvPr id="2" name="Chart 1 Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
    "<p:spPr>"
    '<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    "</p:spPr>"
    "</p:sp>"
    "</mc:Fallback>"
    "</mc:AlternateContent>"
    "</p:spTree></p:cSld>"
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
    "</p:sld>"
)


def _build_box_whisker_pptx() -> io.BytesIO:
    """Return an in-memory ``.pptx`` containing one chartex box-and-whisker shape.

    Starts from a fresh :class:`~pptx.Presentation`, adds a blank slide, writes
    to a BytesIO, then unzips and surgically rewrites ``ppt/slides/slide1.xml``
    and its rels to point at a new hand-authored ``ppt/charts/chartEx1.xml``
    box-and-whisker part. The chartex content-type override is appended to
    ``[Content_Types].xml``.
    """
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    base = io.BytesIO()
    prs.save(base)
    base.seek(0)

    out = io.BytesIO()
    with zipfile.ZipFile(base, "r") as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            name = item.filename
            data = src.read(name)
            if name == "ppt/slides/slide1.xml":
                data = _SLIDE_XML.encode("utf-8")
            elif name == "ppt/slides/_rels/slide1.xml.rels":
                rels_root = data.decode("utf-8")
                # -- append rIdCx pointing at the chartEx part; insert before </Relationships>
                extra = (
                    f'<Relationship Id="rIdCx" Type="{_CHARTEX_RELTYPE}"'
                    f' Target="../charts/chartEx1.xml"/>'
                )
                rels_root = rels_root.replace("</Relationships>", extra + "</Relationships>")
                data = rels_root.encode("utf-8")
            elif name == "[Content_Types].xml":
                extra = (
                    f'<Override PartName="/ppt/charts/chartEx1.xml"'
                    f' ContentType="{_CHARTEX_CT}"/>'
                ).encode("utf-8")
                data = data.replace(b"</Types>", extra + b"</Types>")
            dst.writestr(item, data)
        # -- add the chartEx part itself --
        dst.writestr("ppt/charts/chartEx1.xml", _BOX_WHISKER_CX_XML)

    out.seek(0)
    return out


class DescribeIssue1047BoxWhiskerPassthrough:
    """Exercises box-and-whisker chartex detection and round-trip preservation."""

    def it_detects_a_box_whisker_chartex_graphic_frame(self):
        pkg = _build_box_whisker_pptx()
        prs = Presentation(pkg)
        slide = prs.slides[0]

        graphic_frames = [s for s in slide.shapes if isinstance(s, GraphicFrame)]
        assert len(graphic_frames) == 1, "expected exactly one graphic-frame shape"
        frame = graphic_frames[0]

        assert frame.has_chart is True
        assert frame.has_chartex is True
        assert frame.chart_type is XL_CHART_TYPE.UNSUPPORTED_CHARTEX

    def and_it_reports_the_boxWhisker_layoutId_via_chartex_type(self):
        pkg = _build_box_whisker_pptx()
        prs = Presentation(pkg)
        frame = next(s for s in prs.slides[0].shapes if isinstance(s, GraphicFrame))

        assert frame.chartex_type == "boxWhisker"

    def it_round_trips_the_box_whisker_chart_through_save_and_reopen(self):
        pkg = _build_box_whisker_pptx()
        prs1 = Presentation(pkg)
        buf = io.BytesIO()
        prs1.save(buf)
        buf.seek(0)

        # -- detection survives round-trip --
        prs2 = Presentation(buf)
        frame = next(s for s in prs2.slides[0].shapes if isinstance(s, GraphicFrame))
        assert frame.has_chartex is True
        assert frame.chartex_type == "boxWhisker"
        assert frame.chart_type is XL_CHART_TYPE.UNSUPPORTED_CHARTEX

        # -- chartEx part is still in the package --
        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            assert "ppt/charts/chartEx1.xml" in zf.namelist()
            chartEx_xml = zf.read("ppt/charts/chartEx1.xml").decode("utf-8")
            slide_xml = zf.read("ppt/slides/slide1.xml").decode("utf-8")

        # -- layoutId and non-MVP sub-elements are preserved verbatim --
        assert 'layoutId="boxWhisker"' in chartEx_xml
        assert "cx:statistics" in chartEx_xml
        assert 'quartileMethod="exclusive"' in chartEx_xml

        # -- the slide still carries the chartex graphicData wrapper --
        assert _CHARTEX_URI in slide_xml
        assert "AlternateContent" in slide_xml

    def but_chartex_type_is_None_on_a_non_chartex_graphic_frame(self):
        """A table-carrying graphic-frame (not chartex) reports `chartex_type is None`."""
        from pptx.util import Emu

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_table(
            rows=2,
            cols=2,
            left=Emu(0),
            top=Emu(0),
            width=Emu(1_000_000),
            height=Emu(500_000),
        )

        frame = next(s for s in slide.shapes if isinstance(s, GraphicFrame))
        assert frame.has_chartex is False
        assert frame.chartex_type is None
