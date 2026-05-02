"""Unit-test suite for the `pptx.chart.point` module."""

from __future__ import annotations

import pytest

from pptx.chart.datalabel import DataLabel
from pptx.chart.marker import Marker
from pptx.chart.point import BubblePoints, CategoryPoints, Point, XyPoints
from pptx.dml.chtfmt import ChartFormat

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, instance_mock


class Describe_BasePoints(object):
    def it_supports_indexed_access(self, getitem_fixture):
        points, idx, Point_, ser, point_ = getitem_fixture
        point = points[idx]
        Point_.assert_called_once_with(ser, idx)
        assert point is point_

    def it_raises_on_indexed_access_out_of_range(self):
        points = XyPoints(
            element(
                "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:num"
                "Ref/c:numCache/c:ptCount{val=3})"
            )
        )
        with pytest.raises(IndexError):
            points[-1]
        with pytest.raises(IndexError):
            points[3]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def getitem_fixture(self, request, Point_, point_):
        ser = element(
            "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:num"
            "Ref/c:numCache/c:ptCount{val=3})"
        )
        points = XyPoints(ser)
        idx = 2
        return points, idx, Point_, ser, point_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Point_(self, request, point_):
        return class_mock(request, "pptx.chart.point.Point", return_value=point_)

    @pytest.fixture
    def point_(self, request):
        return instance_mock(request, Point)


class DescribeBubblePoints(object):
    def it_supports_len(self, len_fixture):
        points, expected_value = len_fixture
        assert len(points) == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:ser", 0),
            ("c:ser/c:bubbleSize/c:numRef/c:numCache/c:ptCount{val=3}", 0),
            (
                "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=1},c:yVal/c:numRef"
                "/c:numCache/c:ptCount{val=2},c:bubbleSize/c:numRef/c:numCache/c:pt"
                "Count{val=3})",
                1,
            ),
            (
                "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef"
                "/c:numCache/c:ptCount{val=3},c:bubbleSize/c:numRef/c:numCache/c:pt"
                "Count{val=3})",
                3,
            ),
        ]
    )
    def len_fixture(self, request):
        ser_cxml, expected_value = request.param
        points = BubblePoints(element(ser_cxml))
        return points, expected_value


class DescribeCategoryPoints(object):
    def it_supports_len(self, len_fixture):
        points, expected_value = len_fixture
        assert len(points) == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:ser", 0),
            ("c:ser/c:cat/c:numRef/c:numCache/c:ptCount{val=42}", 42),
            ("c:ser/c:cat/c:numRef/c:numCache/c:ptCount{val=24}", 24),
        ]
    )
    def len_fixture(self, request):
        ser_cxml, expected_value = request.param
        points = CategoryPoints(element(ser_cxml))
        return points, expected_value


class DescribePoint(object):
    def it_provides_access_to_its_data_label(self, data_label_fixture):
        point, DataLabel_, ser, idx, data_label_ = data_label_fixture
        data_label = point.data_label
        DataLabel_.assert_called_once_with(ser, idx)
        assert data_label is data_label_

    def it_provides_access_to_its_format(self, format_fixture):
        point, ChartFormat_, ser, chart_format_, expected_xml = format_fixture
        chart_format = point.format
        ChartFormat_.assert_called_once_with(ser.xpath('c:dPt[c:idx/@val="42"]')[0])
        assert chart_format is chart_format_
        assert point._element.xml == expected_xml

    def it_provides_access_to_its_marker(self, marker_fixture):
        point, Marker_, dPt, marker_ = marker_fixture
        marker = point.marker
        Marker_.assert_called_once_with(dPt)
        assert marker is marker_

    @pytest.mark.parametrize(
        ("ser_cxml", "expected_value"),
        [
            ("c:ser", True),
            ("c:ser/c:dPt/c:idx{val=42}", True),
            ("c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative)", True),
            ("c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=1})", True),
            ("c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=0})", False),
        ],
    )
    def it_knows_whether_it_should_invert_if_negative(self, ser_cxml, expected_value):
        point = Point(element(ser_cxml), 42)
        assert point.invert_if_negative is expected_value

    @pytest.mark.parametrize(
        ("ser_cxml", "new_value", "expected_cxml"),
        [
            (
                "c:ser",
                True,
                "c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=1})",
            ),
            (
                "c:ser/c:dPt/c:idx{val=42}",
                False,
                "c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=0})",
            ),
            (
                "c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=0})",
                True,
                "c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=1})",
            ),
            (
                "c:ser/c:dPt/(c:idx{val=42},c:spPr)",
                False,
                "c:ser/c:dPt/(c:idx{val=42},c:invertIfNegative{val=0},c:spPr)",
            ),
        ],
    )
    def it_can_change_whether_it_inverts_if_negative(self, ser_cxml, new_value, expected_cxml):
        point = Point(element(ser_cxml), 42)
        point.invert_if_negative = new_value
        assert point._element.xml == xml(expected_cxml)

    def it_preserves_series_shadow_when_coloring_a_point_issue_450(self):
        """Regression for #450.

        When a series defines its own `c:spPr` with an outer shadow (or any
        other non-fill visual override such as `a:ln`), assigning a solid
        fill color to a single data point must NOT drop those inherited
        visual properties. PowerPoint treats a point-level `c:spPr` as a
        complete override of the series-level shape properties -- anything
        not present on the point's `c:spPr` is rendered without that effect.
        The fix seeds a freshly-created point `c:spPr` with deep copies of
        the series `c:spPr`'s non-fill children.
        """
        from pptx.dml.color import RGBColor

        ser = element(
            "c:ser/c:spPr/("
            "a:solidFill/a:srgbClr{val=2288CC},"
            "a:ln{w=19050}/a:solidFill/a:srgbClr{val=000000},"
            "a:effectLst/a:outerShdw{blurRad=50800,dist=38100,dir=2700000}"
            "/a:srgbClr{val=000000}/a:alpha{val=40000})"
        )
        point = Point(ser, 1)

        point.format.fill.solid()
        point.format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)

        dPt = ser.xpath('c:dPt[c:idx/@val="1"]')[0]
        # -- point fill override present --
        assert dPt.xpath("c:spPr/a:solidFill/a:srgbClr/@val") == ["FF0000"]
        # -- inherited line preserved on the point --
        assert dPt.xpath("c:spPr/a:ln/a:solidFill/a:srgbClr/@val") == ["000000"]
        # -- inherited outer shadow preserved on the point --
        assert dPt.xpath("c:spPr/a:effectLst/a:outerShdw/@blurRad") == ["50800"]
        assert dPt.xpath("c:spPr/a:effectLst/a:outerShdw/a:srgbClr/@val") == ["000000"]
        # -- series spPr is not mutated (deep copy) --
        ser_spPr = ser.xpath("c:spPr")[0]
        assert ser_spPr.xpath("a:solidFill/a:srgbClr/@val") == ["2288CC"]
        assert ser_spPr.xpath("a:effectLst/a:outerShdw/@blurRad") == ["50800"]

    def it_creates_empty_point_spPr_when_series_has_no_spPr_issue_450(self):
        """Baseline behavior when the series has no own `c:spPr`.

        With no series-level overrides to inherit, the newly-created point
        `c:spPr` must remain empty before the fill is added -- i.e. the
        shadow-preservation logic must not fabricate children out of thin
        air.
        """
        from pptx.dml.color import RGBColor

        ser = element("c:ser")
        point = Point(ser, 0)

        point.format.fill.solid()
        point.format.fill.fore_color.rgb = RGBColor(0xAA, 0xBB, 0xCC)

        dPt = ser.xpath('c:dPt[c:idx/@val="0"]')[0]
        assert dPt.xpath("c:spPr/a:solidFill/a:srgbClr/@val") == ["AABBCC"]
        # -- no spurious inherited children --
        assert dPt.xpath("c:spPr/a:ln") == []
        assert dPt.xpath("c:spPr/a:effectLst") == []

    def it_can_preserve_a_solid_fill_color_on_negative_bars_issue_504(self):
        """Regression for #504.

        Setting a solid fill on a data point is not enough to color a
        negative bar: PowerPoint treats the point-level ``invertIfNegative``
        as defaulting to |True|, so the solid fill is inverted to white.
        Assigning ``False`` to ``point.invert_if_negative`` must emit an
        explicit ``c:invertIfNegative val="0"`` sibling on the ``c:dPt`` so
        the authored fill survives.
        """
        from pptx.dml.color import RGBColor

        point = Point(element("c:ser"), 2)
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = RGBColor(0xF3, 0x5D, 0x5D)
        point.invert_if_negative = False

        dPt = point._element.xpath('c:dPt[c:idx/@val="2"]')[0]
        invertIfNegative_vals = dPt.xpath("c:invertIfNegative/@val")
        srgbClr_vals = dPt.xpath("c:spPr/a:solidFill/a:srgbClr/@val")
        assert invertIfNegative_vals == ["0"]
        assert srgbClr_vals == ["F35D5D"]

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def data_label_fixture(self, DataLabel_, data_label_):
        ser, idx = element("c:ser"), 42
        point = Point(ser, idx)
        return point, DataLabel_, ser, idx, data_label_

    @pytest.fixture(
        params=[
            ("c:ser", "c:ser/c:dPt/c:idx{val=42}"),
            ("c:ser/c:dPt/c:idx{val=42}", "c:ser/c:dPt/c:idx{val=42}"),
            (
                "c:ser/c:dPt/c:idx{val=45}",
                "c:ser/(c:dPt/c:idx{val=45},c:dPt/c:idx{val=42})",
            ),
        ]
    )
    def format_fixture(self, request, ChartFormat_, chart_format_):
        ser_cxml, expected_cxml = request.param
        ser = element(ser_cxml)
        point = Point(ser, 42)
        expected_xml = xml(expected_cxml)
        return point, ChartFormat_, ser, chart_format_, expected_xml

    @pytest.fixture
    def marker_fixture(self, Marker_, marker_):
        ser = element("c:ser/c:dPt/c:idx{val=42}")
        point = Point(ser, 42)
        dPt = ser[0]
        return point, Marker_, dPt, marker_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, chart_format_):
        return class_mock(request, "pptx.chart.point.ChartFormat", return_value=chart_format_)

    @pytest.fixture
    def chart_format_(self, request):
        return instance_mock(request, ChartFormat)

    @pytest.fixture
    def DataLabel_(self, request, data_label_):
        return class_mock(request, "pptx.chart.point.DataLabel", return_value=data_label_)

    @pytest.fixture
    def data_label_(self, request):
        return instance_mock(request, DataLabel)

    @pytest.fixture
    def Marker_(self, request, marker_):
        return class_mock(request, "pptx.chart.point.Marker", return_value=marker_)

    @pytest.fixture
    def marker_(self, request):
        return instance_mock(request, Marker)


class DescribeXyPoints(object):
    def it_supports_len(self, len_fixture):
        points, expected_value = len_fixture
        assert len(points) == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:ser", 0),
            ("c:ser/c:xVal", 0),
            ("c:ser/c:xVal/c:numRef/c:numCache/c:ptCount{val=3}", 0),
            ("c:ser/c:yVal/c:numRef/c:numCache/c:ptCount{val=3}", 0),
            (
                "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=1},c:yVal/c:numRef"
                "/c:numCache/c:ptCount{val=3})",
                1,
            ),
            (
                "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef"
                "/c:numCache/c:ptCount{val=1})",
                1,
            ),
            (
                "c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef"
                "/c:numCache/c:ptCount{val=3})",
                3,
            ),
        ]
    )
    def len_fixture(self, request):
        ser_cxml, expected_value = request.param
        points = XyPoints(element(ser_cxml))
        return points, expected_value
