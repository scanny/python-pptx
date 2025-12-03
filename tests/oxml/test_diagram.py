"""Unit-test suite for pptx.oxml.diagram module."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.diagram import (
    CT_DiagramConnection,
    CT_DiagramConnectionList,
    CT_DiagramData,
    CT_DiagramPoint,
    CT_DiagramPointList,
    CT_DiagramText,
)


class DescribeCT_DiagramData:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramData` object."""

    def it_has_ptLst_property(self):
        """Test that CT_DiagramData has ptLst property."""
        xml = (
            '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
            "  <dgm:ptLst/>"
            "  <dgm:cxnLst/>"
            "</dgm:dataModel>"
        )
        data_model = parse_xml(xml)
        assert isinstance(data_model.ptLst, CT_DiagramPointList)

    def it_has_cxnLst_property(self):
        """Test that CT_DiagramData has cxnLst property."""
        xml = (
            '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
            "  <dgm:ptLst/>"
            "  <dgm:cxnLst/>"
            "</dgm:dataModel>"
        )
        data_model = parse_xml(xml)
        assert isinstance(data_model.cxnLst, CT_DiagramConnectionList)


class DescribeCT_DiagramPointList:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramPointList` object."""

    def it_can_iterate_points(self):
        """Test iterating over diagram points."""
        xml = (
            '<dgm:ptLst xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
            '  <dgm:pt modelId="{id1}"/>'
            '  <dgm:pt modelId="{id2}"/>'
            '  <dgm:pt modelId="{id3}"/>'
            "</dgm:ptLst>"
        )
        pt_lst = parse_xml(xml)
        points = list(pt_lst.iter_pts())

        assert len(points) == 3
        assert all(isinstance(pt, CT_DiagramPoint) for pt in points)


class DescribeCT_DiagramPoint:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramPoint` object."""

    def it_has_modelId_attribute(self):
        """Test that point has modelId attribute."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' modelId="{TEST-ID}"/>'
        )
        pt = parse_xml(xml)
        assert pt.modelId == "{TEST-ID}"

    def it_has_type_attribute(self):
        """Test that point has type attribute."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' type="doc"/>'
        )
        pt = parse_xml(xml)
        assert pt.type == "doc"

    def it_can_extract_text_content(self):
        """Test extracting text from a point."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' modelId="{test}">'
            "  <dgm:t>"
            "    <a:bodyPr/>"
            "    <a:lstStyle/>"
            "    <a:p>"
            "      <a:r>"
            "        <a:t>Hello World</a:t>"
            "      </a:r>"
            "    </a:p>"
            "  </dgm:t>"
            "</dgm:pt>"
        )
        pt = parse_xml(xml)
        assert pt.text == "Hello World"

    def it_can_extract_multiple_text_runs(self):
        """Test extracting text from multiple runs."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' modelId="{test}">'
            "  <dgm:t>"
            "    <a:bodyPr/>"
            "    <a:lstStyle/>"
            "    <a:p>"
            "      <a:r><a:t>Hello </a:t></a:r>"
            "      <a:r><a:t>World</a:t></a:r>"
            "    </a:p>"
            "  </dgm:t>"
            "</dgm:pt>"
        )
        pt = parse_xml(xml)
        assert pt.text == "Hello World"

    def it_returns_empty_string_when_no_text(self):
        """Test that point without text returns empty string."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' modelId="{test}"/>'
        )
        pt = parse_xml(xml)
        assert pt.text == ""

    def it_returns_empty_string_when_t_element_is_empty(self):
        """Test that point with empty t element returns empty string."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' modelId="{test}">'
            "  <dgm:t>"
            "    <a:bodyPr/>"
            "    <a:lstStyle/>"
            "    <a:p>"
            "      <a:endParaRPr/>"
            "    </a:p>"
            "  </dgm:t>"
            "</dgm:pt>"
        )
        pt = parse_xml(xml)
        assert pt.text == ""

    def it_can_set_text_content(self):
        """Test setting text content in a point."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' modelId="{test}">'
            "  <dgm:t>"
            "    <a:bodyPr/>"
            "    <a:lstStyle/>"
            "    <a:p>"
            "      <a:r><a:t>Old Text</a:t></a:r>"
            "    </a:p>"
            "  </dgm:t>"
            "</dgm:pt>"
        )
        pt = parse_xml(xml)
        pt.text = "New Text"
        assert pt.text == "New Text"

    def it_can_replace_text_in_multiple_runs(self):
        """Test setting text when multiple runs exist."""
        xml = (
            '<dgm:pt xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' modelId="{test}">'
            "  <dgm:t>"
            "    <a:bodyPr/>"
            "    <a:lstStyle/>"
            "    <a:p>"
            "      <a:r><a:t>Hello </a:t></a:r>"
            "      <a:r><a:t>World</a:t></a:r>"
            "    </a:p>"
            "  </dgm:t>"
            "</dgm:pt>"
        )
        pt = parse_xml(xml)
        pt.text = "Replaced"
        assert pt.text == "Replaced"


class DescribeCT_DiagramText:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramText` object."""

    def it_is_a_valid_element_class(self):
        """Test that CT_DiagramText can be instantiated."""
        xml = (
            '<dgm:t xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "  <a:bodyPr/>"
            "  <a:lstStyle/>"
            "  <a:p/>"
            "</dgm:t>"
        )
        t = parse_xml(xml)
        assert isinstance(t, CT_DiagramText)


class DescribeCT_DiagramConnectionList:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramConnectionList` object."""

    def it_can_iterate_connections(self):
        """Test iterating over connections."""
        xml = (
            '<dgm:cxnLst xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
            '  <dgm:cxn modelId="{c1}" srcId="{s1}" destId="{d1}"/>'
            '  <dgm:cxn modelId="{c2}" srcId="{s2}" destId="{d2}"/>'
            "</dgm:cxnLst>"
        )
        cxn_lst = parse_xml(xml)
        connections = list(cxn_lst.iter_cxns())

        assert len(connections) == 2
        assert all(isinstance(cxn, CT_DiagramConnection) for cxn in connections)


class DescribeCT_DiagramConnection:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramConnection` object."""

    def it_has_connection_attributes(self):
        """Test that connection has required attributes."""
        xml = (
            '<dgm:cxn xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
            ' modelId="{MODEL}" srcId="{SRC}" destId="{DEST}" type="presOf"/>'
        )
        cxn = parse_xml(xml)

        assert cxn.modelId == "{MODEL}"
        assert cxn.srcId == "{SRC}"
        assert cxn.destId == "{DEST}"
        assert cxn.type == "presOf"
