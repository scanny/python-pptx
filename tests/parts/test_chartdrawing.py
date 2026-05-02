# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.chartdrawing` (issue #351)."""

from __future__ import annotations

from unittest.mock import Mock

import pytest
from lxml import etree

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import nsdecls, qn
from pptx.parts.chartdrawing import ChartDrawingPart

CDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"


def _parse(xml: str) -> etree._Element:
    """Parse `xml` into a raw lxml element without the custom class-lookup."""
    return etree.fromstring(xml.encode("utf-8"))


def _user_shapes(*anchor_fragments: str) -> etree._Element:
    """Return a `cdr:userShapes` element containing `anchor_fragments` as XML."""
    inner = "".join(anchor_fragments)
    xml = f"<cdr:userShapes {nsdecls('cdr', 'a', 'r')}>{inner}</cdr:userShapes>"
    return _parse(xml)


_REL_SIZE_ANCHOR = (
    "<cdr:relSizeAnchor>"
    "<cdr:from><cdr:x>0.1</cdr:x><cdr:y>0.1</cdr:y></cdr:from>"
    "<cdr:to><cdr:x>0.4</cdr:x><cdr:y>0.2</cdr:y></cdr:to>"
    "<cdr:sp/>"
    "</cdr:relSizeAnchor>"
)

_ABS_SIZE_ANCHOR = (
    "<cdr:absSizeAnchor>"
    "<cdr:from><cdr:x>0.5</cdr:x><cdr:y>0.6</cdr:y></cdr:from>"
    '<cdr:ext cx="914400" cy="228600"/>'
    "<cdr:cxnSp/>"
    "</cdr:absSizeAnchor>"
)


class DescribeChartDrawingPart:
    """Unit-test suite for `pptx.parts.chartdrawing.ChartDrawingPart` objects."""

    def it_constructs_a_new_empty_part_with_a_caller_supplied_partname(self):
        package = Mock(name="package")
        partname = PackURI("/ppt/charts/drawings/drawing1.xml")

        part = ChartDrawingPart.new(package, partname)

        assert isinstance(part, ChartDrawingPart)
        assert part.partname == partname
        assert part.content_type == CT.DML_CHARTSHAPES
        # -- the root element is an empty `cdr:userShapes`
        assert part._element.tag == qn("cdr:userShapes")
        assert len(part._element) == 0

    def it_allocates_a_fresh_partname_when_none_is_supplied(self):
        package = Mock(name="package")
        package.next_partname.return_value = PackURI("/ppt/charts/drawings/drawing42.xml")

        part = ChartDrawingPart.new(package)

        package.next_partname.assert_called_once_with("/ppt/charts/drawings/drawing%d.xml")
        assert part.partname == PackURI("/ppt/charts/drawings/drawing42.xml")

    def it_yields_each_anchor_child_in_document_order(self):
        element = _user_shapes(_REL_SIZE_ANCHOR, _ABS_SIZE_ANCHOR, _REL_SIZE_ANCHOR)
        part = ChartDrawingPart(
            PackURI("/ppt/charts/drawings/drawing1.xml"),
            CT.DML_CHARTSHAPES,
            Mock(name="package"),
            element,  # pyright: ignore[reportArgumentType]
        )

        anchors = list(part.iter_anchor_elements())

        assert [a.tag for a in anchors] == [
            f"{{{CDR_NS}}}relSizeAnchor",
            f"{{{CDR_NS}}}absSizeAnchor",
            f"{{{CDR_NS}}}relSizeAnchor",
        ]

    def it_skips_non_anchor_children_when_iterating_anchors(self):
        # -- `cdr:userShapes` in the XSD allows only anchors as direct children,
        # -- but real-world PowerPoint files occasionally carry `mc:AlternateContent`
        # -- wrappers emitted by the markup-compatibility extension. The iterator
        # -- must filter those out so callers get a stream of true anchors. --
        element = _user_shapes(
            _REL_SIZE_ANCHOR,
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/'
            'markup-compatibility/2006"/>',
            _ABS_SIZE_ANCHOR,
        )
        part = ChartDrawingPart(
            PackURI("/ppt/charts/drawings/drawing1.xml"),
            CT.DML_CHARTSHAPES,
            Mock(name="package"),
            element,  # pyright: ignore[reportArgumentType]
        )

        anchor_tags = [a.tag for a in part.iter_anchor_elements()]

        assert anchor_tags == [
            f"{{{CDR_NS}}}relSizeAnchor",
            f"{{{CDR_NS}}}absSizeAnchor",
        ]

    @pytest.mark.parametrize(
        ("fragments", "expected"),
        [
            ((), 0),
            ((_REL_SIZE_ANCHOR,), 1),
            ((_REL_SIZE_ANCHOR, _ABS_SIZE_ANCHOR), 2),
            ((_REL_SIZE_ANCHOR, _REL_SIZE_ANCHOR, _ABS_SIZE_ANCHOR), 3),
        ],
    )
    def it_knows_its_anchor_count(self, fragments, expected):
        element = _user_shapes(*fragments)
        part = ChartDrawingPart(
            PackURI("/ppt/charts/drawings/drawing1.xml"),
            CT.DML_CHARTSHAPES,
            Mock(name="package"),
            element,  # pyright: ignore[reportArgumentType]
        )

        assert part.anchor_count == expected

    def it_serializes_back_to_bytes_via_the_XmlPart_blob_property(self):
        # -- the part is an `XmlPart`, so round-trip fidelity must flow through
        # -- `.blob`. This guards against a regression where `ChartDrawingPart`
        # -- gets promoted to a non-XmlPart and silently drops annotation XML. --
        element = _user_shapes(_REL_SIZE_ANCHOR)
        part = ChartDrawingPart(
            PackURI("/ppt/charts/drawings/drawing1.xml"),
            CT.DML_CHARTSHAPES,
            Mock(name="package"),
            element,  # pyright: ignore[reportArgumentType]
        )

        blob = part.blob

        assert b"userShapes" in blob
        assert b"relSizeAnchor" in blob
