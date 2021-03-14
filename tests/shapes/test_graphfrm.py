# encoding: utf-8

"""Unit-test suite for pptx.shapes.graphfrm module."""

import pytest

from pptx.chart.chart import Chart
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.parts.slide import SlidePart
from pptx.shapes.graphfrm import GraphicFrame, _OleFormat
from pptx.shapes.shapetree import SlideShapes
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_TABLE,
)

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock, property_mock


class DescribeGraphicFrame(object):
    """Unit-test suite for `pptx.shapes.graphfrm.GraphicFrame` object."""

    def it_provides_access_to_the_chart_it_contains(
        self, request, has_chart_prop_, chart_part_, chart_
    ):
        has_chart_prop_.return_value = True
        property_mock(request, GraphicFrame, "chart_part", return_value=chart_part_)
        chart_part_.chart = chart_

        assert GraphicFrame(None, None).chart is chart_

    def but_it_raises_on_chart_if_there_isnt_one(self, has_chart_prop_):
        has_chart_prop_.return_value = False

        with pytest.raises(ValueError) as e:
            GraphicFrame(None, None).chart
        assert str(e.value) == "shape does not contain a chart"

    def it_provides_access_to_its_chart_part(self, request, chart_part_):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData/c:chart{r:id=rId42}"
        )
        property_mock(
            request,
            GraphicFrame,
            "part",
            return_value=instance_mock(
                request, SlidePart, related_parts={"rId42": chart_part_}
            ),
        )
        graphic_frame = GraphicFrame(graphicFrame, None)

        assert graphic_frame.chart_part is chart_part_

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, True),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_TABLE, False),
        ),
    )
    def it_knows_whether_it_contains_a_chart(self, graphicData_uri, expected_value):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri
        )
        assert GraphicFrame(graphicFrame, None).has_chart is expected_value

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, False),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_TABLE, True),
        ),
    )
    def it_knows_whether_it_contains_a_table(self, graphicData_uri, expected_value):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri
        )
        assert GraphicFrame(graphicFrame, None).has_table is expected_value

    def it_provides_access_to_the_OleFormat_object(self, request):
        ole_format_ = instance_mock(request, _OleFormat)
        _OleFormat_ = class_mock(
            request, "pptx.shapes.graphfrm._OleFormat", return_value=ole_format_
        )
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=http://schemas.openxmlformats"
            ".org/presentationml/2006/ole}"
        )
        parent_ = instance_mock(request, SlideShapes)
        graphic_frame = GraphicFrame(graphicFrame, parent_)

        ole_format = graphic_frame.ole_format

        _OleFormat_.assert_called_once_with(graphicFrame.graphicData, parent_)
        assert ole_format is ole_format_

    def but_it_raises_on_ole_format_when_this_is_not_an_OLE_object(self):
        graphic_frame = GraphicFrame(
            element(
                "p:graphicFrame/a:graphic/a:graphicData{uri=http://schemas.openxmlfor"
                "mats.org/drawingml/2006/table}"
            ),
            None,
        )
        with pytest.raises(ValueError) as e:
            graphic_frame.ole_format
        assert str(e.value) == "not an OLE-object shape"

    def it_raises_on_shadow(self):
        graphic_frame = GraphicFrame(None, None)
        with pytest.raises(NotImplementedError):
            graphic_frame.shadow

    @pytest.mark.parametrize(
        "uri, oleObj_child, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, None, MSO_SHAPE_TYPE.CHART),
            (GRAPHIC_DATA_URI_OLEOBJ, "embed", MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT),
            (GRAPHIC_DATA_URI_OLEOBJ, "link", MSO_SHAPE_TYPE.LINKED_OLE_OBJECT),
            (GRAPHIC_DATA_URI_TABLE, None, MSO_SHAPE_TYPE.TABLE),
            ("foobar", None, None),
        ),
    )
    def it_knows_its_shape_type(self, uri, oleObj_child, expected_value):
        graphicFrame = element(
            (
                "p:graphicFrame/a:graphic/a:graphicData{uri=%s}/p:oleObj/p:%s"
                % (uri, oleObj_child)
            )
            if oleObj_child
            else "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % uri
        )
        assert GraphicFrame(graphicFrame, None).shape_type is expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def chart_(self, request):
        return instance_mock(request, Chart)

    @pytest.fixture
    def chart_part_(self, request, chart_):
        return instance_mock(request, ChartPart, chart=chart_)

    @pytest.fixture
    def has_chart_prop_(self, request):
        return property_mock(request, GraphicFrame, "has_chart")


class Describe_OleFormat(object):
    """Unit-test suite for `pptx.shapes.graphfrm._OleFormat` object."""

    def it_provides_access_to_the_OLE_object_blob(self, request):
        ole_obj_part_ = instance_mock(request, EmbeddedPackagePart, blob=b"0123456789")
        property_mock(
            request,
            _OleFormat,
            "part",
            return_value=instance_mock(
                request, SlidePart, related_parts={"rId7": ole_obj_part_}
            ),
        )
        graphicData = element("a:graphicData/p:oleObj{r:id=rId7}")

        assert _OleFormat(graphicData, None).blob == b"0123456789"

    def it_knows_the_OLE_object_prog_id(self):
        graphicData = element("a:graphicData/p:oleObj{progId=Excel.Sheet.12}")
        assert _OleFormat(graphicData, None).prog_id == "Excel.Sheet.12"

    def it_knows_whether_to_show_the_OLE_object_as_an_icon(self):
        graphicData = element("a:graphicData/p:oleObj{showAsIcon=1}")
        assert _OleFormat(graphicData, None).show_as_icon is True
