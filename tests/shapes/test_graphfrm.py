"""Unit-test suite for pptx.shapes.graphfrm module."""

from __future__ import annotations

import pytest

from pptx.chart.chart import Chart
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part
from pptx.parts.chart import ChartPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.parts.slide import SlidePart
from pptx.shapes.graphfrm import GraphicFrame, SmartArt, _OleFormat
from pptx.shapes.model3d import Model3D
from pptx.shapes.shapetree import SlideShapes
from pptx.spec import (
    GRAPHIC_DATA_URI_CHART,
    GRAPHIC_DATA_URI_CHARTEX,
    GRAPHIC_DATA_URI_MODEL_3D,
    GRAPHIC_DATA_URI_OLEOBJ,
    GRAPHIC_DATA_URI_SMART_ART,
    GRAPHIC_DATA_URI_TABLE,
)

from ..unitutil.cxml import element
from ..unitutil.mock import class_mock, instance_mock, property_mock


class DescribeGraphicFrame(object):
    """Unit-test suite for `pptx.shapes.graphfrm.GraphicFrame` object."""

    def it_provides_access_to_the_chart_it_contains(
        self, request, has_chart_prop_, has_chartex_prop_, chart_part_, chart_
    ):
        has_chart_prop_.return_value = True
        has_chartex_prop_.return_value = False
        property_mock(request, GraphicFrame, "chart_part", return_value=chart_part_)
        chart_part_.chart = chart_

        assert GraphicFrame(None, None).chart is chart_

    def but_it_raises_on_chart_if_there_isnt_one(self, has_chart_prop_):
        has_chart_prop_.return_value = False

        with pytest.raises(ValueError) as e:
            GraphicFrame(None, None).chart
        assert str(e.value) == "shape does not contain a chart"

    def and_it_raises_NotImplementedError_on_chart_for_a_chartex_chart(
        self, has_chart_prop_, has_chartex_prop_
    ):
        has_chart_prop_.return_value = True
        has_chartex_prop_.return_value = True

        with pytest.raises(NotImplementedError) as e:
            GraphicFrame(None, None).chart
        assert "UNSUPPORTED_CHARTEX" in str(e.value)

    def it_provides_access_to_its_chart_part(self, request, chart_part_):
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.related_part.return_value = chart_part_
        property_mock(request, GraphicFrame, "part", return_value=slide_part_)
        graphic_frame = GraphicFrame(
            element("p:graphicFrame/a:graphic/a:graphicData/c:chart{r:id=rId42}"), None
        )

        chart_part = graphic_frame.chart_part

        slide_part_.related_part.assert_called_once_with("rId42")
        assert chart_part is chart_part_

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, True),
            (GRAPHIC_DATA_URI_CHARTEX, True),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_TABLE, False),
        ),
    )
    def it_knows_whether_it_contains_a_chart(self, graphicData_uri, expected_value):
        graphicFrame = element("p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri)
        assert GraphicFrame(graphicFrame, None).has_chart is expected_value

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, False),
            (GRAPHIC_DATA_URI_CHARTEX, True),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_TABLE, False),
        ),
    )
    def it_knows_whether_it_contains_a_chartex_chart(self, graphicData_uri, expected_value):
        graphicFrame = element("p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri)
        assert GraphicFrame(graphicFrame, None).has_chartex is expected_value

    def it_reports_UNSUPPORTED_CHARTEX_for_chartex_chart_type(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_CHARTEX
        )

        chart_type = GraphicFrame(graphicFrame, None).chart_type

        assert chart_type is XL_CHART_TYPE.UNSUPPORTED_CHARTEX

    def but_it_raises_on_chart_type_if_not_a_chart(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_TABLE
        )
        with pytest.raises(ValueError) as e:
            GraphicFrame(graphicFrame, None).chart_type
        assert str(e.value) == "shape does not contain a chart"

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, False),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_TABLE, True),
        ),
    )
    def it_knows_whether_it_contains_a_table(self, graphicData_uri, expected_value):
        graphicFrame = element("p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri)
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
            (GRAPHIC_DATA_URI_CHARTEX, None, MSO_SHAPE_TYPE.CHART),
            (GRAPHIC_DATA_URI_OLEOBJ, "embed", MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT),
            (GRAPHIC_DATA_URI_OLEOBJ, "link", MSO_SHAPE_TYPE.LINKED_OLE_OBJECT),
            (GRAPHIC_DATA_URI_SMART_ART, None, MSO_SHAPE_TYPE.IGX_GRAPHIC),
            (GRAPHIC_DATA_URI_TABLE, None, MSO_SHAPE_TYPE.TABLE),
            ("foobar", None, None),
        ),
    )
    def it_knows_its_shape_type(self, uri, oleObj_child, expected_value):
        graphicFrame = element(
            ("p:graphicFrame/a:graphic/a:graphicData{uri=%s}/p:oleObj/p:%s" % (uri, oleObj_child))
            if oleObj_child
            else "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % uri
        )
        assert GraphicFrame(graphicFrame, None).shape_type is expected_value

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, False),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_SMART_ART, True),
            (GRAPHIC_DATA_URI_TABLE, False),
        ),
    )
    def it_knows_whether_it_contains_SmartArt(self, graphicData_uri, expected_value):
        graphicFrame = element("p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri)
        assert GraphicFrame(graphicFrame, None).has_smart_art is expected_value

    def it_provides_access_to_a_SmartArt_object(self, request):
        smart_art_ = instance_mock(request, SmartArt)
        SmartArt_ = class_mock(request, "pptx.shapes.graphfrm.SmartArt", return_value=smart_art_)
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_SMART_ART
        )
        parent_ = instance_mock(request, SlideShapes)
        graphic_frame = GraphicFrame(graphicFrame, parent_)

        smart_art = graphic_frame.smart_art

        SmartArt_.assert_called_once_with(graphicFrame.graphicData, parent_)
        assert smart_art is smart_art_

    def but_it_raises_on_smart_art_when_not_a_SmartArt_shape(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_TABLE
        )
        with pytest.raises(ValueError) as e:
            GraphicFrame(graphicFrame, None).smart_art
        assert str(e.value) == "shape does not contain SmartArt"

    @pytest.mark.parametrize(
        "graphicData_uri, expected_value",
        (
            (GRAPHIC_DATA_URI_CHART, False),
            (GRAPHIC_DATA_URI_MODEL_3D, True),
            (GRAPHIC_DATA_URI_OLEOBJ, False),
            (GRAPHIC_DATA_URI_SMART_ART, False),
            (GRAPHIC_DATA_URI_TABLE, False),
        ),
    )
    def it_knows_whether_it_contains_a_3D_model(self, graphicData_uri, expected_value):
        graphicFrame = element("p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % graphicData_uri)
        assert GraphicFrame(graphicFrame, None).has_model_3d is expected_value

    def it_provides_access_to_a_Model3D_object(self, request):
        model_3d_ = instance_mock(request, Model3D)
        Model3D_ = class_mock(request, "pptx.shapes.graphfrm.Model3D", return_value=model_3d_)
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}/am3d:model3D{r:embed=rId9,ext=glb}"
            % GRAPHIC_DATA_URI_MODEL_3D
        )
        parent_ = instance_mock(request, SlideShapes)
        graphic_frame = GraphicFrame(graphicFrame, parent_)

        model_3d = graphic_frame.model_3d

        Model3D_.assert_called_once_with(graphicFrame.graphicData.model_3d, parent_)
        assert model_3d is model_3d_

    def but_it_raises_on_model_3d_when_not_a_3D_model_shape(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_TABLE
        )
        with pytest.raises(ValueError) as e:
            GraphicFrame(graphicFrame, None).model_3d
        assert str(e.value) == "shape does not contain a 3D model"

    def and_it_raises_on_model_3d_when_the_am3d_model3D_child_is_absent(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_MODEL_3D
        )
        with pytest.raises(ValueError) as e:
            GraphicFrame(graphicFrame, None).model_3d
        assert str(e.value) == "shape does not contain a 3D model"

    def it_provides_the_raw_model_3d_xml(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}/am3d:model3D{r:embed=rId9,ext=glb}"
            % GRAPHIC_DATA_URI_MODEL_3D
        )

        xml_str = GraphicFrame(graphicFrame, None).model_3d_xml

        assert xml_str is not None
        assert "model3D" in xml_str
        assert 'ext="glb"' in xml_str
        assert "rId9" in xml_str

    def but_model_3d_xml_is_None_when_not_a_3D_model_shape(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_TABLE
        )
        assert GraphicFrame(graphicFrame, None).model_3d_xml is None

    def and_model_3d_xml_is_None_when_the_am3d_model3D_child_is_absent(self):
        graphicFrame = element(
            "p:graphicFrame/a:graphic/a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_MODEL_3D
        )
        assert GraphicFrame(graphicFrame, None).model_3d_xml is None

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

    @pytest.fixture
    def has_chartex_prop_(self, request):
        return property_mock(request, GraphicFrame, "has_chartex")


class Describe_OleFormat(object):
    """Unit-test suite for `pptx.shapes.graphfrm._OleFormat` object."""

    def it_provides_access_to_the_OLE_object_blob(self, request):
        ole_obj_part_ = instance_mock(request, EmbeddedPackagePart, blob=b"0123456789")
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.related_part.return_value = ole_obj_part_
        property_mock(request, _OleFormat, "part", return_value=slide_part_)
        ole_format = _OleFormat(element("a:graphicData/p:oleObj{r:id=rId7}"), None)

        blob = ole_format.blob

        slide_part_.related_part.assert_called_once_with("rId7")
        assert blob == b"0123456789"

    def it_knows_the_OLE_object_prog_id(self):
        graphicData = element("a:graphicData/p:oleObj{progId=Excel.Sheet.12}")
        assert _OleFormat(graphicData, None).prog_id == "Excel.Sheet.12"

    def it_knows_whether_to_show_the_OLE_object_as_an_icon(self):
        graphicData = element("a:graphicData/p:oleObj{showAsIcon=1}")
        assert _OleFormat(graphicData, None).show_as_icon is True


class DescribeSmartArt(object):
    """Unit-test suite for `pptx.shapes.graphfrm.SmartArt` object."""

    @pytest.mark.parametrize(
        "attr, rId_attr, rId_value",
        (
            ("data_xml", "r:dm", "rId2"),
            ("layout_xml", "r:lo", "rId3"),
            ("colors_xml", "r:cs", "rId5"),
            ("quick_style_xml", "r:qs", "rId4"),
        ),
    )
    def it_provides_read_only_access_to_each_of_the_four_SmartArt_parts(
        self, request, attr, rId_attr, rId_value
    ):
        diagram_part_ = instance_mock(request, Part, blob=b"<dgm:payload/>")
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.related_part.return_value = diagram_part_
        property_mock(request, SmartArt, "part", return_value=slide_part_)
        graphicData = element(
            "a:graphicData{uri=%s}/dgm:relIds{%s=%s}"
            % (GRAPHIC_DATA_URI_SMART_ART, rId_attr, rId_value)
        )

        value = getattr(SmartArt(graphicData, None), attr)

        slide_part_.related_part.assert_called_once_with(rId_value)
        assert value == b"<dgm:payload/>"

    @pytest.mark.parametrize("attr", ("data_xml", "layout_xml", "colors_xml", "quick_style_xml"))
    def but_it_returns_None_when_the_dgm_relIds_element_is_missing(self, attr):
        graphicData = element("a:graphicData{uri=%s}" % GRAPHIC_DATA_URI_SMART_ART)
        assert getattr(SmartArt(graphicData, None), attr) is None

    @pytest.mark.parametrize(
        "attr, rId_attr",
        (
            ("data_xml", "r:dm"),
            ("layout_xml", "r:lo"),
            ("colors_xml", "r:cs"),
            ("quick_style_xml", "r:qs"),
        ),
    )
    def and_it_returns_None_when_the_named_rId_is_absent(self, attr, rId_attr):
        # -- use a different rId attr so the named one is missing --
        other_attr = "r:dm" if rId_attr != "r:dm" else "r:lo"
        graphicData = element(
            "a:graphicData{uri=%s}/dgm:relIds{%s=rId7}" % (GRAPHIC_DATA_URI_SMART_ART, other_attr)
        )
        assert getattr(SmartArt(graphicData, None), attr) is None

    def and_it_returns_None_when_the_rId_cannot_be_resolved(self, request):
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.related_part.side_effect = KeyError("rId2")
        property_mock(request, SmartArt, "part", return_value=slide_part_)
        graphicData = element(
            "a:graphicData{uri=%s}/dgm:relIds{r:dm=rId2}" % GRAPHIC_DATA_URI_SMART_ART
        )

        assert SmartArt(graphicData, None).data_xml is None


class DescribeModel3D(object):
    """Unit-test suite for `pptx.shapes.model3d.Model3D` object."""

    def it_exposes_the_embedded_relationship_id(self):
        model3D = element("am3d:model3D{r:embed=rId9,ext=glb}")
        assert Model3D(model3D, None).embedded_rel_id == "rId9"

    def but_embedded_rel_id_is_None_when_the_attribute_is_absent(self):
        model3D = element("am3d:model3D{ext=glb}")
        assert Model3D(model3D, None).embedded_rel_id is None

    @pytest.mark.parametrize("ext_value", ("glb", "obj", "fbx"))
    def it_exposes_the_model_file_extension(self, ext_value):
        model3D = element("am3d:model3D{r:embed=rId9,ext=%s}" % ext_value)
        assert Model3D(model3D, None).ext == ext_value

    def but_ext_is_None_when_the_attribute_is_absent(self):
        model3D = element("am3d:model3D{r:embed=rId9}")
        assert Model3D(model3D, None).ext is None

    def it_provides_the_embedded_media_blob(self, request):
        media_part_ = instance_mock(request, Part, blob=b"glTF-binary-payload")
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.related_part.return_value = media_part_
        property_mock(request, Model3D, "part", return_value=slide_part_)
        model3D = element("am3d:model3D{r:embed=rId9,ext=glb}")

        blob = Model3D(model3D, None).media_blob

        slide_part_.related_part.assert_called_once_with("rId9")
        assert blob == b"glTF-binary-payload"

    def but_media_blob_is_None_when_the_embed_rId_is_absent(self):
        model3D = element("am3d:model3D{ext=glb}")
        assert Model3D(model3D, None).media_blob is None

    def and_media_blob_is_None_when_the_rId_cannot_be_resolved(self, request):
        slide_part_ = instance_mock(request, SlidePart)
        slide_part_.related_part.side_effect = KeyError("rId9")
        property_mock(request, Model3D, "part", return_value=slide_part_)
        model3D = element("am3d:model3D{r:embed=rId9,ext=glb}")

        assert Model3D(model3D, None).media_blob is None
