# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.chart.series` module."""

from __future__ import annotations

import pytest

from pptx.chart.datalabel import DataLabels
from pptx.chart.marker import Marker
from pptx.chart.point import BubblePoints, CategoryPoints, XyPoints
from pptx.chart.series import (
    AreaSeries,
    BarSeries,
    BubbleSeries,
    ErrorBars,
    LineSeries,
    PieSeries,
    RadarSeries,
    SeriesCollection,
    Trendline,
    XySeries,
    _BaseCategorySeries,
    _BaseSeries,
    _MarkerMixin,
    _SeriesFactory,
)
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import (
    XL_ERROR_BAR_DIRECTION,
    XL_ERROR_BAR_INCLUDE,
    XL_ERROR_BAR_TYPE,
    XL_TRENDLINE_TYPE,
)

from ..unitutil.cxml import element, xml
from ..unitutil.mock import class_mock, function_mock, instance_mock


class Describe_BaseSeries(object):
    def it_knows_its_name(self, name_fixture):
        series, expected_value = name_fixture
        assert series.name == expected_value

    def it_knows_its_position_in_the_series_sequence(self, index_fixture):
        series, expected_value = index_fixture
        assert series.index == expected_value

    def it_provides_access_to_its_format(self, format_fixture):
        series, ChartFormat_, ser, format_ = format_fixture
        format = series.format
        ChartFormat_.assert_called_once_with(ser)
        assert format is format_

    @pytest.mark.parametrize(
        "ser_cxml, expected_value",
        [
            ("c:ser/(c:idx{val=0},c:order{val=0})", False),
            (
                "c:ser/(c:idx{val=0},c:order{val=0},c:errBars/(c:errBarType{val=bo"
                "th},c:errValType{val=fixedVal}))",
                True,
            ),
        ],
    )
    def it_knows_whether_it_has_error_bars(self, ser_cxml, expected_value):
        series = _BaseSeries(element(ser_cxml))
        assert series.has_error_bars is expected_value

    def it_returns_None_for_error_bars_when_absent(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        assert series.error_bars is None

    def it_provides_access_to_its_error_bars(self):
        ser_cxml = (
            "c:ser/(c:idx{val=0},c:order{val=0},c:errBars/(c:errBarType{val=both}"
            ",c:errValType{val=fixedVal},c:val{val=1.5}))"
        )
        series = _BaseSeries(element(ser_cxml))
        error_bars = series.error_bars
        assert isinstance(error_bars, ErrorBars)
        assert error_bars.type == XL_ERROR_BAR_TYPE.FIXED_VALUE
        assert error_bars.include == XL_ERROR_BAR_INCLUDE.BOTH
        assert error_bars.value == 1.5

    def it_can_attach_error_bars_via_set_error_bars(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))

        error_bars = series.set_error_bars(
            type_=XL_ERROR_BAR_TYPE.PERCENT,
            value=10.0,
            include=XL_ERROR_BAR_INCLUDE.BOTH,
        )

        assert isinstance(error_bars, ErrorBars)
        assert series.has_error_bars is True
        assert series.error_bars.type == XL_ERROR_BAR_TYPE.PERCENT
        assert series.error_bars.value == 10.0

    def it_replaces_existing_error_bars_when_set_error_bars_called_twice(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        series.set_error_bars(type_=XL_ERROR_BAR_TYPE.FIXED_VALUE, value=1.0)

        series.set_error_bars(type_=XL_ERROR_BAR_TYPE.PERCENT, value=5.0)

        # -- only one c:errBars child should be present --
        assert len(series._element.xpath("c:errBars")) == 1
        assert series.error_bars.type == XL_ERROR_BAR_TYPE.PERCENT
        assert series.error_bars.value == 5.0

    def it_removes_error_bars_when_set_to_None(self):
        ser_cxml = (
            "c:ser/(c:idx{val=0},c:order{val=0},c:errBars/(c:errBarType{val=both}"
            ",c:errValType{val=fixedVal},c:val{val=1.5}))"
        )
        series = _BaseSeries(element(ser_cxml))

        series.error_bars = None

        assert series.has_error_bars is False

    def it_raises_on_setting_error_bars_to_a_non_ErrorBars_value(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        with pytest.raises(TypeError):
            series.error_bars = 42

    def it_allows_assigning_a_new_ErrorBars_object_to_error_bars(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        # -- build a detached errBars element via the existing setter path --
        new_error_bars = series.set_error_bars(type_=XL_ERROR_BAR_TYPE.STDEV, value=2.0)
        # -- drop it, then reassign --
        series.error_bars = None
        series.error_bars = new_error_bars

        assert series.has_error_bars is True
        assert series.error_bars.type == XL_ERROR_BAR_TYPE.STDEV

    # trendlines -----------------------------------------------------

    def it_returns_an_empty_list_of_trendlines_when_none_present(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        assert series.trendlines == []

    def it_provides_access_to_its_trendlines(self):
        ser_cxml = (
            "c:ser/(c:idx{val=0},c:order{val=0},c:trendline/c:trendlineType{val=linear}"
            ",c:trendline/(c:trendlineType{val=poly},c:order{val=3}))"
        )
        series = _BaseSeries(element(ser_cxml))
        tls = series.trendlines
        assert len(tls) == 2
        assert all(isinstance(tl, Trendline) for tl in tls)
        assert tls[0].trendline_type == XL_TRENDLINE_TYPE.LINEAR
        assert tls[1].trendline_type == XL_TRENDLINE_TYPE.POLYNOMIAL
        assert tls[1].order == 3

    def it_can_attach_a_trendline_via_add_trendline(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))

        tl = series.add_trendline(XL_TRENDLINE_TYPE.LINEAR, display_equation=True)

        assert isinstance(tl, Trendline)
        assert len(series.trendlines) == 1
        assert series.trendlines[0].trendline_type == XL_TRENDLINE_TYPE.LINEAR
        assert series.trendlines[0].display_equation is True

    def it_can_attach_multiple_trendlines(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))

        series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)
        series.add_trendline(XL_TRENDLINE_TYPE.POLYNOMIAL, order=3)
        series.add_trendline(XL_TRENDLINE_TYPE.MOVING_AVG, period=5)

        tls = series.trendlines
        assert len(tls) == 3
        assert tls[0].trendline_type == XL_TRENDLINE_TYPE.LINEAR
        assert tls[1].trendline_type == XL_TRENDLINE_TYPE.POLYNOMIAL
        assert tls[1].order == 3
        assert tls[2].trendline_type == XL_TRENDLINE_TYPE.MOVING_AVG
        assert tls[2].period == 5

    def it_can_delete_a_trendline_via_trendline_delete(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        series.add_trendline(XL_TRENDLINE_TYPE.LINEAR)
        series.add_trendline(XL_TRENDLINE_TYPE.POWER)

        series.trendlines[0].delete()

        tls = series.trendlines
        assert len(tls) == 1
        assert tls[0].trendline_type == XL_TRENDLINE_TYPE.POWER

    def it_raises_on_add_trendline_with_polynomial_order_out_of_range(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        with pytest.raises(ValueError):
            series.add_trendline(XL_TRENDLINE_TYPE.POLYNOMIAL, order=1)
        with pytest.raises(ValueError):
            series.add_trendline(XL_TRENDLINE_TYPE.POLYNOMIAL, order=7)

    def it_raises_on_add_trendline_with_moving_avg_period_too_small(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        with pytest.raises(ValueError):
            series.add_trendline(XL_TRENDLINE_TYPE.MOVING_AVG, period=1)

    # source-range / category-range / name-range ---------------------

    def it_knows_its_source_range(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:f"Sheet1!$B$2:$F$2")'
        series = _BaseSeries(element(ser_cxml))
        assert series.source_range == "Sheet1!$B$2:$F$2"

    def it_returns_None_for_source_range_when_no_numRef_is_present(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        assert series.source_range is None

    def it_returns_None_for_source_range_when_f_is_absent(self):
        ser_cxml = "c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:numCache)"
        series = _BaseSeries(element(ser_cxml))
        assert series.source_range is None

    def it_can_set_its_source_range(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:f"Sheet1!$B$2:$F$2")'
        series = _BaseSeries(element(ser_cxml))

        series.source_range = "Sheet1!$B$2:$F$6"

        assert series.source_range == "Sheet1!$B$2:$F$6"

    def it_adds_an_f_element_when_source_range_setter_runs_with_no_f_present(self):
        ser_cxml = "c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:numCache)"
        series = _BaseSeries(element(ser_cxml))

        series.source_range = "Sheet1!$A$1:$A$3"

        assert series.source_range == "Sheet1!$A$1:$A$3"
        # -- f should be the first child of numRef, before numCache --
        numRef = series._element.xpath("c:val/c:numRef")[0]
        assert numRef[0].tag.endswith("}f")
        assert numRef[1].tag.endswith("}numCache")

    def it_raises_on_setting_source_range_when_no_numRef_container_exists(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        with pytest.raises(ValueError):
            series.source_range = "Sheet1!$A$1:$A$3"

    def it_raises_on_setting_source_range_to_None(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:f"Sheet1!$B$2")'
        series = _BaseSeries(element(ser_cxml))
        with pytest.raises(ValueError):
            series.source_range = None

    def it_knows_its_category_range_from_strRef(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:cat/c:strRef/c:f"Sheet1!$A$2:$A$6")'
        series = _BaseSeries(element(ser_cxml))
        assert series.category_range == "Sheet1!$A$2:$A$6"

    def it_knows_its_category_range_from_numRef(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:cat/c:numRef/c:f"Sheet1!$A$2:$A$6")'
        series = _BaseSeries(element(ser_cxml))
        assert series.category_range == "Sheet1!$A$2:$A$6"

    def it_returns_None_for_category_range_when_absent(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        assert series.category_range is None

    def it_knows_its_name_range(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:tx/c:strRef/c:f"Sheet1!$B$1")'
        series = _BaseSeries(element(ser_cxml))
        assert series.name_range == "Sheet1!$B$1"

    def it_returns_None_for_name_range_when_absent(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        assert series.name_range is None

    def it_parses_values_sheet_reference_from_source_range(self):
        ser_cxml = 'c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:f"Sheet1!$B$2:$F$2")'
        series = _BaseSeries(element(ser_cxml))
        ref = series.values_sheet_reference
        assert ref is not None
        assert ref.sheet_name == "Sheet1"
        assert ref.a1_range == "$B$2:$F$2"
        # -- tuple unpacking works (namedtuple) --
        sheet, a1 = ref
        assert (sheet, a1) == ("Sheet1", "$B$2:$F$2")

    def it_returns_None_for_values_sheet_reference_when_no_source_range(self):
        series = _BaseSeries(element("c:ser/(c:idx{val=0},c:order{val=0})"))
        assert series.values_sheet_reference is None

    def it_unquotes_quoted_sheet_names_in_values_sheet_reference(self):
        ser_cxml = (
            "c:ser/(c:idx{val=0},c:order{val=0},c:val/c:numRef/c:f"
            "\"'My Sheet'!$B$2:$F$2\")"
        )
        series = _BaseSeries(element(ser_cxml))
        ref = series.values_sheet_reference
        assert ref == ("My Sheet", "$B$2:$F$2")

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def format_fixture(self, ChartFormat_, chart_format_):
        ser = element("c:ser")
        series = _BaseSeries(ser)
        return series, ChartFormat_, ser, chart_format_

    @pytest.fixture
    def index_fixture(self):
        ser_cxml, expected_value = "c:ser/c:idx{val=42}", 42
        series = _BaseSeries(element(ser_cxml))
        return series, expected_value

    @pytest.fixture(
        params=[
            ("c:ser", ""),
            ('c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v"foobar"', "foobar"),
        ]
    )
    def name_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = _BaseSeries(element(ser_cxml))
        return series, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartFormat_(self, request, chart_format_):
        return class_mock(request, "pptx.chart.series.ChartFormat", return_value=chart_format_)

    @pytest.fixture
    def chart_format_(self, request):
        return instance_mock(request, ChartFormat)


class Describe_BaseCategorySeries(object):
    def it_is_a_BaseSeries_subclass(self, subclass_fixture):
        base_category_series = subclass_fixture
        assert isinstance(base_category_series, _BaseSeries)

    def it_provides_access_to_its_data_labels(self, data_labels_fixture, DataLabels_, data_labels_):
        ser, expected_dLbls_xml = data_labels_fixture
        DataLabels_.return_value = data_labels_
        series = _BaseCategorySeries(ser)

        data_labels = series.data_labels

        dLbls = ser.xpath("c:dLbls")[0]
        assert dLbls.xml == expected_dLbls_xml
        DataLabels_.assert_called_once_with(dLbls, chart_type=None)
        assert data_labels is data_labels_

    def it_provides_access_to_its_points(self, points_fixture):
        series, CategoryPoints_, ser, points_ = points_fixture
        points = series.points
        CategoryPoints_.assert_called_once_with(ser)
        assert points is points_

    def it_knows_its_values(self, values_get_fixture):
        series, expected_value = values_get_fixture
        assert series.values == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (
                "c:ser",
                "c:dLbls/(c:showLegendKey{val=0},c:showVal{val=0},c:showCatName{val"
                "=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val"
                "=0},c:showLeaderLines{val=1})",
            ),
            ("c:ser/c:dLbls", "c:dLbls"),
        ]
    )
    def data_labels_fixture(self, request):
        ser_cxml, expected_dLbls_cxml = request.param
        ser = element(ser_cxml)
        expected_dLbls_xml = xml(expected_dLbls_cxml)
        return ser, expected_dLbls_xml

    @pytest.fixture
    def points_fixture(self, CategoryPoints_, points_):
        ser = element("c:ser")
        series = _BaseCategorySeries(ser)
        return series, CategoryPoints_, ser, points_

    @pytest.fixture
    def subclass_fixture(self):
        return _BaseCategorySeries(None)

    @pytest.fixture(
        params=[
            ("c:ser", ()),
            ("c:ser/c:val/c:numRef", ()),
            ("c:ser/c:val/c:numLit", ()),
            ("c:ser/c:val/c:numRef/c:numCache", ()),
            ("c:ser/c:val/c:numRef/c:numCache/c:ptCount{val=0}", ()),
            (
                'c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v"' '1.1")',
                (1.1,),
            ),
            (
                'c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=0}/c:v"'
                '1.1",c:pt{idx=2}/c:v"3.3")',
                (1.1, None, 3.3),
            ),
            (
                'c:ser/c:val/c:numLit/(c:ptCount{val=3},c:pt{idx=0}/c:v"1.1",c:pt{i'
                'dx=2}/c:v"3.3")',
                (1.1, None, 3.3),
            ),
            (
                'c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=2}/c:v"'
                '3.3",c:pt{idx=0}/c:v"1.1")',
                (1.1, None, 3.3),
            ),
        ]
    )
    def values_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = _BaseCategorySeries(element(ser_cxml))
        return series, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def CategoryPoints_(self, request, points_):
        return class_mock(request, "pptx.chart.series.CategoryPoints", return_value=points_)

    @pytest.fixture
    def DataLabels_(self, request):
        return class_mock(request, "pptx.chart.series.DataLabels")

    @pytest.fixture
    def data_labels_(self, request):
        return instance_mock(request, DataLabels)

    @pytest.fixture
    def points_(self, request):
        return instance_mock(request, CategoryPoints)


class Describe_MarkerMixin(object):
    def it_provides_access_to_the_series_marker(self, marker_fixture):
        series, Marker_, ser, marker_ = marker_fixture
        marker = series.marker
        Marker_.assert_called_once_with(ser)
        assert marker is marker_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def marker_fixture(self, Marker_, marker_):
        ser = element("c:ser")
        series = LineSeries(ser)
        return series, Marker_, ser, marker_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Marker_(self, request, marker_):
        return class_mock(request, "pptx.chart.series.Marker", return_value=marker_)

    @pytest.fixture
    def marker_(self, request):
        return instance_mock(request, Marker)


class DescribeAreaSeries(object):
    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        area_series = subclass_fixture
        assert isinstance(area_series, _BaseCategorySeries)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def subclass_fixture(self):
        return AreaSeries(None)


class DescribeBarSeries(object):
    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        bar_series = subclass_fixture
        assert isinstance(bar_series, _BaseCategorySeries)

    def it_knows_whether_it_should_invert_if_negative(self, invert_if_negative_get_fixture):
        bar_series, expected_value = invert_if_negative_get_fixture
        assert bar_series.invert_if_negative == expected_value

    def it_can_change_whether_it_inverts_if_negative(self, invert_if_negative_set_fixture):
        bar_series, new_value, expected_xml = invert_if_negative_set_fixture
        bar_series.invert_if_negative = new_value
        assert bar_series._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:ser", True),
            ("c:ser/c:invertIfNegative", True),
            ("c:ser/c:invertIfNegative{val=1}", True),
            ("c:ser/c:invertIfNegative{val=0}", False),
        ]
    )
    def invert_if_negative_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        bar_series = BarSeries(element(ser_cxml))
        return bar_series, expected_value

    @pytest.fixture(
        params=[
            ("c:ser/c:order", True, "c:ser/(c:order,c:invertIfNegative{val=1})"),
            ("c:ser/c:order", False, "c:ser/(c:order,c:invertIfNegative{val=0})"),
            (
                "c:ser/(c:spPr,c:invertIfNegative{val=0})",
                True,
                "c:ser/(c:spPr,c:invertIfNegative{val=1})",
            ),
            (
                "c:ser/(c:tx,c:invertIfNegative{val=1})",
                False,
                "c:ser/(c:tx,c:invertIfNegative{val=0})",
            ),
        ]
    )
    def invert_if_negative_set_fixture(self, request):
        ser_cxml, new_value, expected_ser_cxml = request.param
        bar_series = BarSeries(element(ser_cxml))
        expected_xml = xml(expected_ser_cxml)
        return bar_series, new_value, expected_xml

    @pytest.fixture
    def subclass_fixture(self):
        return BarSeries(None)


class Describe_BubbleSeries(object):
    def it_provides_access_to_its_points(self, points_fixture):
        series, BubblePoints_, ser, points_ = points_fixture
        points = series.points
        BubblePoints_.assert_called_once_with(ser)
        assert points is points_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def points_fixture(self, BubblePoints_, points_):
        ser = element("c:ser")
        series = BubbleSeries(ser)
        return series, BubblePoints_, ser, points_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BubblePoints_(self, request, points_):
        return class_mock(request, "pptx.chart.series.BubblePoints", return_value=points_)

    @pytest.fixture
    def points_(self, request):
        return instance_mock(request, BubblePoints)


class DescribeLineSeries(object):
    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _BaseCategorySeries)

    def it_uses__MarkerMixin(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _MarkerMixin)

    def it_knows_whether_it_should_use_curve_smoothing(self, smooth_get_fixture):
        series, expected_value = smooth_get_fixture
        assert series.smooth == expected_value

    def it_can_change_whether_it_uses_curve_smoothing(self, smooth_set_fixture):
        series, new_value, expected_xml = smooth_set_fixture
        series.smooth = new_value
        assert series._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:ser", True),
            ("c:ser/c:smooth", True),
            ("c:ser/c:smooth{val=1}", True),
            ("c:ser/c:smooth{val=0}", False),
        ]
    )
    def smooth_get_fixture(self, request):
        ser_cxml, expected_value = request.param
        series = LineSeries(element(ser_cxml))
        return series, expected_value

    @pytest.fixture(
        params=[
            ("c:ser", True, "c:ser/c:smooth"),
            ("c:ser/c:smooth", False, "c:ser/c:smooth{val=0}"),
            ("c:ser/c:smooth{val=0}", True, "c:ser/c:smooth"),
        ]
    )
    def smooth_set_fixture(self, request):
        ser_cxml, new_value, expected_ser_cxml = request.param
        series = LineSeries(element(ser_cxml))
        expected_xml = xml(expected_ser_cxml)
        return series, new_value, expected_xml

    @pytest.fixture
    def subclass_fixture(self):
        return LineSeries(None)


class DescribePieSeries(object):
    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        pie_series = subclass_fixture
        assert isinstance(pie_series, _BaseCategorySeries)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def subclass_fixture(self):
        return PieSeries(None)


class DescribeRadarSeries(object):
    def it_is_a_BaseCategorySeries_subclass(self, subclass_fixture):
        radar_series = subclass_fixture
        assert isinstance(radar_series, _BaseCategorySeries)

    def it_uses__MarkerMixin(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _MarkerMixin)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def subclass_fixture(self):
        return RadarSeries(None)


class Describe_XySeries(object):
    def it_uses__MarkerMixin(self, subclass_fixture):
        line_series = subclass_fixture
        assert isinstance(line_series, _MarkerMixin)

    def it_provides_access_to_its_points(self, points_fixture):
        series, XyPoints_, ser, points_ = points_fixture
        points = series.points
        XyPoints_.assert_called_once_with(ser)
        assert points is points_

    def it_knows_its_values(self, values_get_fixture):
        series, expected_values = values_get_fixture
        assert series.values == expected_values

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def points_fixture(self, XyPoints_, points_):
        ser = element("c:ser")
        series = XySeries(ser)
        return series, XyPoints_, ser, points_

    @pytest.fixture
    def subclass_fixture(self):
        return XySeries(None)

    @pytest.fixture(
        params=[
            ("c:ser", ()),
            ("c:ser/c:yVal/c:numRef", ()),
            ("c:ser/c:val/c:numRef/c:numCache", ()),
            (
                "c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v" '"1.1")',
                (1.1,),
            ),
            (
                "c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=0}/c:v"
                '"1.1",c:pt{idx=2}/c:v"3.3")',
                (1.1, None, 3.3),
            ),
            ("c:ser/c:val/c:numLit", ()),
            (
                'c:ser/c:yVal/c:numLit/(c:ptCount{val=3},c:pt{idx=0}/c:v"1.1",c:pt{'
                'idx=2}/c:v"3.3")',
                (1.1, None, 3.3),
            ),
        ]
    )
    def values_get_fixture(self, request):
        ser_cxml, expected_values = request.param
        series = XySeries(element(ser_cxml))
        return series, expected_values

    # fixture components ---------------------------------------------

    @pytest.fixture
    def XyPoints_(self, request, points_):
        return class_mock(request, "pptx.chart.series.XyPoints", return_value=points_)

    @pytest.fixture
    def points_(self, request):
        return instance_mock(request, XyPoints)


class DescribeSeriesCollection(object):
    def it_supports_indexed_access(self, getitem_fixture):
        series_collection, index, _SeriesFactory_, ser, series_ = getitem_fixture
        series = series_collection[index]
        _SeriesFactory_.assert_called_once_with(ser)
        assert series is series_

    def it_supports_len(self, len_fixture):
        series_collection, expected_len = len_fixture
        assert len(series_collection) == expected_len

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:barChart/c:ser/c:order{val=42}", 0, 0),
            (
                "c:barChart/(c:ser/c:order{val=9},c:ser/c:order{val=6},c:ser/c:orde" "r{val=3})",
                2,
                0,
            ),
            (
                "c:plotArea/(c:barChart/(c:ser/(c:idx{val=1},c:order{val=3}),c:ser/"
                "(c:idx{val=0},c:order{val=0})),c:lineChart/c:ser/(c:idx{val=2},c:o"
                "rder{val=1}))",
                1,
                0,
            ),
        ]
    )
    def getitem_fixture(self, request, _SeriesFactory_, series_):
        cxml, index, _offset = request.param
        parent_elm = element(cxml)
        ser = parent_elm.xpath(".//c:ser")[_offset]
        series_collection = SeriesCollection(parent_elm)
        return series_collection, index, _SeriesFactory_, ser, series_

    @pytest.fixture(
        params=[
            ("c:barChart", 0),
            ("c:barChart/c:ser/c:order{val=4}", 1),
            (
                "c:barChart/(c:ser/c:order{val=4},c:ser/c:order{val=1},c:ser/c:orde" "r{val=6})",
                3,
            ),
            ("c:plotArea/c:barChart", 0),
            ("c:plotArea/c:barChart/c:ser/c:order{val=4}", 1),
            (
                "c:plotArea/c:barChart/(c:ser/c:order{val=4},c:ser/c:order{val=1},c"
                ":ser/c:order{val=6})",
                3,
            ),
        ]
    )
    def len_fixture(self, request):
        cxml, expected_len = request.param
        series_collection = SeriesCollection(element(cxml))
        return series_collection, expected_len

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _SeriesFactory_(self, request, series_):
        return function_mock(request, "pptx.chart.series._SeriesFactory", return_value=series_)

    @pytest.fixture
    def series_(self, request):
        return instance_mock(request, _BaseSeries)


class Describe_SeriesFactory(object):
    def it_contructs_a_series_object_from_a_plot_element(self, call_fixture):
        ser, SeriesCls_, series_ = call_fixture
        series = _SeriesFactory(ser)
        SeriesCls_.assert_called_once_with(ser)
        assert series is series_

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:areaChart/c:ser", "AreaSeries"),
            ("c:barChart/c:ser", "BarSeries"),
            ("c:bubbleChart/c:ser", "BubbleSeries"),
            ("c:doughnutChart/c:ser", "PieSeries"),
            ("c:lineChart/c:ser", "LineSeries"),
            ("c:pieChart/c:ser", "PieSeries"),
            ("c:radarChart/c:ser", "RadarSeries"),
            ("c:scatterChart/c:ser", "XySeries"),
        ]
    )
    def call_fixture(self, request):
        xChart_cxml, cls_name = request.param
        ser = element(xChart_cxml).ser_lst[0]
        SeriesCls_ = class_mock(request, "pptx.chart.series.%s" % cls_name)
        series_ = SeriesCls_.return_value
        return ser, SeriesCls_, series_


class DescribeErrorBars(object):
    @pytest.mark.parametrize(
        "errBars_cxml, expected_type",
        [
            ("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal})", "fixedVal"),
            ("c:errBars/(c:errBarType{val=both},c:errValType{val=percentage})", "percentage"),
            ("c:errBars/(c:errBarType{val=both},c:errValType{val=stdDev})", "stdDev"),
            ("c:errBars/(c:errBarType{val=both},c:errValType{val=stdErr})", "stdErr"),
            ("c:errBars/(c:errBarType{val=both},c:errValType{val=cust})", "cust"),
        ],
    )
    def it_knows_its_type(self, errBars_cxml, expected_type):
        error_bars = ErrorBars(element(errBars_cxml))
        assert error_bars.type == XL_ERROR_BAR_TYPE.from_xml(expected_type)

    def it_defaults_type_to_fixed_value_when_errValType_absent(self):
        error_bars = ErrorBars(element("c:errBars/c:errBarType{val=both}"))
        assert error_bars.type == XL_ERROR_BAR_TYPE.FIXED_VALUE

    def it_can_change_its_type(self):
        errBars = element(
            "c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal},c:val{val=1})"
        )
        error_bars = ErrorBars(errBars)

        error_bars.type = XL_ERROR_BAR_TYPE.PERCENT

        assert error_bars.type == XL_ERROR_BAR_TYPE.PERCENT

    @pytest.mark.parametrize(
        "errBars_cxml, expected_include",
        [
            ("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal})", "both"),
            ("c:errBars/(c:errBarType{val=plus},c:errValType{val=fixedVal})", "plus"),
            ("c:errBars/(c:errBarType{val=minus},c:errValType{val=fixedVal})", "minus"),
        ],
    )
    def it_knows_which_sides_are_included(self, errBars_cxml, expected_include):
        error_bars = ErrorBars(element(errBars_cxml))
        assert error_bars.include == XL_ERROR_BAR_INCLUDE.from_xml(expected_include)

    def it_defaults_include_to_BOTH_when_errBarType_absent(self):
        error_bars = ErrorBars(element("c:errBars/c:errValType{val=fixedVal}"))
        assert error_bars.include == XL_ERROR_BAR_INCLUDE.BOTH

    def it_can_change_its_include(self):
        error_bars = ErrorBars(
            element("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal})")
        )

        error_bars.include = XL_ERROR_BAR_INCLUDE.PLUS_VALUES

        assert error_bars.include == XL_ERROR_BAR_INCLUDE.PLUS_VALUES

    def it_knows_its_value(self):
        error_bars = ErrorBars(
            element("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal},c:val{val=2.5})")
        )
        assert error_bars.value == 2.5

    def it_returns_None_for_value_when_c_val_absent(self):
        error_bars = ErrorBars(
            element("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal})")
        )
        assert error_bars.value is None

    def it_can_set_its_value(self):
        error_bars = ErrorBars(
            element("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal})")
        )

        error_bars.value = 3.25

        assert error_bars.value == 3.25

    def it_can_overwrite_its_value(self):
        error_bars = ErrorBars(
            element("c:errBars/(c:errBarType{val=both},c:errValType{val=fixedVal},c:val{val=1})")
        )

        error_bars.value = 7.5

        assert error_bars.value == 7.5

    @pytest.mark.parametrize(
        "errBars_cxml, expected",
        [
            ("c:errBars", None),
            ("c:errBars/c:errDir{val=x}", XL_ERROR_BAR_DIRECTION.X),
            ("c:errBars/c:errDir{val=y}", XL_ERROR_BAR_DIRECTION.Y),
        ],
    )
    def it_knows_its_direction(self, errBars_cxml, expected):
        error_bars = ErrorBars(element(errBars_cxml))
        assert error_bars.direction == expected

    def it_can_change_its_direction(self):
        error_bars = ErrorBars(element("c:errBars"))

        error_bars.direction = XL_ERROR_BAR_DIRECTION.Y

        assert error_bars.direction == XL_ERROR_BAR_DIRECTION.Y

    def it_can_clear_its_direction(self):
        error_bars = ErrorBars(element("c:errBars/c:errDir{val=x}"))

        error_bars.direction = None

        assert error_bars.direction is None

    @pytest.mark.parametrize(
        "errBars_cxml, expected",
        [
            ("c:errBars", True),
            ("c:errBars/c:noEndCap", False),
            ("c:errBars/c:noEndCap{val=1}", False),
            ("c:errBars/c:noEndCap{val=0}", True),
        ],
    )
    def it_knows_whether_end_caps_are_drawn(self, errBars_cxml, expected):
        error_bars = ErrorBars(element(errBars_cxml))
        assert error_bars.end_cap is expected

    def it_can_turn_end_caps_on_and_off(self):
        error_bars = ErrorBars(element("c:errBars"))

        error_bars.end_cap = False
        assert error_bars.end_cap is False

        error_bars.end_cap = True
        assert error_bars.end_cap is True

    def it_provides_access_to_its_format(self):
        error_bars = ErrorBars(element("c:errBars"))
        fmt = error_bars.format
        assert type(fmt).__name__ == "ChartFormat"
        # -- same object returned twice (lazyproperty) --
        assert error_bars.format is fmt


class DescribeTrendline(object):
    @pytest.mark.parametrize(
        ("trendline_cxml", "expected_type"),
        [
            ("c:trendline/c:trendlineType{val=linear}", XL_TRENDLINE_TYPE.LINEAR),
            ("c:trendline/c:trendlineType{val=poly}", XL_TRENDLINE_TYPE.POLYNOMIAL),
            ("c:trendline/c:trendlineType{val=log}", XL_TRENDLINE_TYPE.LOGARITHMIC),
            ("c:trendline/c:trendlineType{val=power}", XL_TRENDLINE_TYPE.POWER),
            ("c:trendline/c:trendlineType{val=exp}", XL_TRENDLINE_TYPE.EXPONENTIAL),
            ("c:trendline/c:trendlineType{val=movingAvg}", XL_TRENDLINE_TYPE.MOVING_AVG),
        ],
    )
    def it_knows_its_trendline_type(self, trendline_cxml, expected_type):
        trendline = Trendline(element(trendline_cxml))
        assert trendline.trendline_type == expected_type

    def it_defaults_to_linear_when_trendline_type_attribute_is_missing(self):
        # -- schema default for @val on c:trendlineType is "linear" --
        trendline = Trendline(element("c:trendline/c:trendlineType"))
        assert trendline.trendline_type == XL_TRENDLINE_TYPE.LINEAR

    def it_can_change_its_trendline_type(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        trendline.trendline_type = XL_TRENDLINE_TYPE.POLYNOMIAL
        assert trendline.trendline_type == XL_TRENDLINE_TYPE.POLYNOMIAL

    def it_reads_order_with_schema_default_when_absent(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=poly}"))
        assert trendline.order == 2

    def it_reads_its_polynomial_order(self):
        trendline = Trendline(element("c:trendline/(c:trendlineType{val=poly},c:order{val=4})"))
        assert trendline.order == 4

    def it_writes_its_polynomial_order(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=poly}"))
        trendline.order = 5
        assert trendline.order == 5

    def it_raises_on_order_out_of_range(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=poly}"))
        with pytest.raises(ValueError):
            trendline.order = 1
        with pytest.raises(ValueError):
            trendline.order = 7

    def it_reads_period_with_schema_default_when_absent(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=movingAvg}"))
        assert trendline.period == 2

    def it_reads_its_moving_average_period(self):
        trendline = Trendline(
            element("c:trendline/(c:trendlineType{val=movingAvg},c:period{val=7})")
        )
        assert trendline.period == 7

    def it_writes_its_moving_average_period(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=movingAvg}"))
        trendline.period = 4
        assert trendline.period == 4

    def it_raises_on_period_less_than_two(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=movingAvg}"))
        with pytest.raises(ValueError):
            trendline.period = 1

    @pytest.mark.parametrize("prop", ["forward", "backward", "intercept"])
    def it_returns_None_for_extrapolation_props_when_absent(self, prop):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        assert getattr(trendline, prop) is None

    @pytest.mark.parametrize(
        ("prop", "child_cxml", "expected"),
        [
            ("forward", "c:forward{val=2.5}", 2.5),
            ("backward", "c:backward{val=1.0}", 1.0),
            ("intercept", "c:intercept{val=0.75}", 0.75),
        ],
    )
    def it_reads_forward_backward_and_intercept(self, prop, child_cxml, expected):
        cxml = "c:trendline/(c:trendlineType{val=linear}," + child_cxml + ")"
        trendline = Trendline(element(cxml))
        assert getattr(trendline, prop) == expected

    @pytest.mark.parametrize("prop", ["forward", "backward", "intercept"])
    def it_writes_and_clears_forward_backward_and_intercept(self, prop):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        setattr(trendline, prop, 3.25)
        assert getattr(trendline, prop) == 3.25
        setattr(trendline, prop, None)
        assert getattr(trendline, prop) is None

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("c:trendline/c:trendlineType{val=linear}", False),
            ("c:trendline/(c:trendlineType{val=linear},c:dispEq{val=1})", True),
            ("c:trendline/(c:trendlineType{val=linear},c:dispEq{val=0})", False),
        ],
    )
    def it_reads_its_display_equation_flag(self, cxml, expected):
        trendline = Trendline(element(cxml))
        assert trendline.display_equation is expected

    def it_writes_and_clears_display_equation(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        trendline.display_equation = True
        assert trendline.display_equation is True
        trendline.display_equation = False
        assert trendline.display_equation is False

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("c:trendline/c:trendlineType{val=linear}", False),
            ("c:trendline/(c:trendlineType{val=linear},c:dispRSqr{val=1})", True),
            ("c:trendline/(c:trendlineType{val=linear},c:dispRSqr{val=0})", False),
        ],
    )
    def it_reads_its_display_r_squared_flag(self, cxml, expected):
        trendline = Trendline(element(cxml))
        assert trendline.display_r_squared is expected

    def it_writes_and_clears_display_r_squared(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        trendline.display_r_squared = True
        assert trendline.display_r_squared is True
        trendline.display_r_squared = False
        assert trendline.display_r_squared is False

    def it_returns_empty_string_for_name_when_absent(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        assert trendline.name == ""

    def it_reads_its_name(self):
        cxml = 'c:trendline/(c:name"My fit",c:trendlineType{val=linear})'
        trendline = Trendline(element(cxml))
        assert trendline.name == "My fit"

    def it_writes_and_clears_its_name(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        trendline.name = "Linear fit"
        assert trendline.name == "Linear fit"
        trendline.name = None
        assert trendline.name == ""

    def it_provides_access_to_its_format(self):
        trendline = Trendline(element("c:trendline/c:trendlineType{val=linear}"))
        fmt = trendline.format
        assert type(fmt).__name__ == "ChartFormat"
        # -- same object returned twice (lazyproperty) --
        assert trendline.format is fmt

    def it_can_be_deleted_from_its_parent(self):
        ser_cxml = (
            "c:ser/(c:idx{val=0},c:order{val=0}"
            ",c:trendline/c:trendlineType{val=linear}"
            ",c:trendline/c:trendlineType{val=poly})"
        )
        ser = element(ser_cxml)
        series = _BaseSeries(ser)
        series.trendlines[0].delete()
        remaining = series.trendlines
        assert len(remaining) == 1
        assert remaining[0].trendline_type == XL_TRENDLINE_TYPE.POLYNOMIAL
