# encoding: utf-8

"""
Test suite for pptx.chart.data module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest


from pptx.chart.data import (
    _BaseChartData, BubbleChartData, BubbleDataPoint, BubbleSeriesData,
    Categories, Category, CategoryChartData, CategoryDataPoint,
    CategorySeriesData, ChartData, _SeriesData, XyChartData, XyDataPoint,
    XySeriesData
)
from pptx.chart.xlsx import CategoryWorkbookWriter
from pptx.enum.base import EnumValue

from ..unitutil.mock import call, class_mock, instance_mock, property_mock


class DescribeChartData(object):

    def it_knows_its_current_categories(self, categories_get_fixture):
        chart_data, expected_value = categories_get_fixture
        assert chart_data.categories == expected_value

    def it_can_change_its_categories(self, categories_set_fixture):
        chart_data, new_value, expected_value = categories_set_fixture
        chart_data.categories = new_value
        assert chart_data.categories == expected_value

    def it_provides_access_to_its_current_series_data(self, series_fixture):
        chart_data, expected_value = series_fixture
        assert chart_data.series == expected_value

    def it_can_add_a_series(self, add_series_fixture):
        chart_data, name, values, _SeriesData_, series_data_ = (
            add_series_fixture
        )
        chart_data.add_series(name, values)
        _SeriesData_.assert_called_once_with(
            0, name, values, chart_data._categories, 0
        )
        assert chart_data._series_lst[0] is series_data_

    def it_can_generate_chart_part_XML_for_its_data(self, xml_bytes_fixture):
        chart_data, chart_type_, ChartXmlWriter_ = xml_bytes_fixture[:3]
        expected_bytes, series_lst_ = xml_bytes_fixture[3:]

        xml_bytes = chart_data.xml_bytes(chart_type_)

        ChartXmlWriter_.assert_called_once_with(chart_type_, series_lst_)
        assert xml_bytes == expected_bytes

    def it_can_provide_its_data_as_an_Excel_workbook(self, xlsx_fixture):
        chart_data, WorkbookWriter_ = xlsx_fixture[:2]
        categories, series_, xlsx_blob_ = xlsx_fixture[2:]
        xlsx_blob = chart_data.xlsx_blob
        WorkbookWriter_.xlsx_blob.assert_called_once_with(
            categories, series_
        )
        assert xlsx_blob is xlsx_blob_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_series_fixture(
            self, request, categories, _SeriesData_, series_data_):
        chart_data = ChartData()
        chart_data.categories = categories
        name = 'Series Foo'
        values = (1.1, 2.2, 3.3)
        return chart_data, name, values, _SeriesData_, series_data_

    @pytest.fixture
    def categories_get_fixture(self, categories):
        chart_data = ChartData()
        chart_data._categories = list(categories)
        expected_value = categories
        return chart_data, expected_value

    @pytest.fixture(params=[
        (['Foo', 'Bar'],       ('Foo', 'Bar')),
        (iter(['Foo', 'Bar']), ('Foo', 'Bar')),
    ])
    def categories_set_fixture(self, request):
        new_value, expected_value = request.param
        chart_data = ChartData()
        return chart_data, new_value, expected_value

    @pytest.fixture
    def series_fixture(self, series_data_):
        chart_data = ChartData()
        chart_data._series_lst = [series_data_, series_data_]
        expected_value = (series_data_, series_data_)
        return chart_data, expected_value

    @pytest.fixture
    def xlsx_fixture(
            self, request, WorkbookWriter_, categories, series_lst_,
            xlsx_blob_):
        chart_data = ChartData()
        chart_data.categories = categories
        chart_data._series_lst = series_lst_
        return (
            chart_data, WorkbookWriter_, categories, series_lst_, xlsx_blob_
        )

    @pytest.fixture
    def xml_bytes_fixture(self, chart_type_, ChartXmlWriter_, series_lst_):
        chart_data = ChartData()
        chart_data._series_lst = series_lst_
        expected_bytes = 'ƒøØßår'.encode('utf-8')
        return (
            chart_data, chart_type_, ChartXmlWriter_, expected_bytes,
            series_lst_
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def categories(self):
        return ('Foo', 'Bar', 'Baz')

    @pytest.fixture
    def ChartXmlWriter_(self, request):
        ChartXmlWriter_ = class_mock(
            request, 'pptx.chart.data.ChartXmlWriter'
        )
        ChartXmlWriter_.return_value.xml = 'ƒøØßår'
        return ChartXmlWriter_

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)

    @pytest.fixture
    def _SeriesData_(self, request, series_data_):
        return class_mock(
            request, 'pptx.chart.data._SeriesData', return_value=series_data_
        )

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, _SeriesData)

    @pytest.fixture
    def series_lst_(self, request):
        return instance_mock(request, list)

    @pytest.fixture
    def WorkbookWriter_(self, request, xlsx_blob_):
        WorkbookWriter_ = class_mock(
            request, 'pptx.chart.data.WorkbookWriter'
        )
        WorkbookWriter_.xlsx_blob.return_value = xlsx_blob_
        return WorkbookWriter_

    @pytest.fixture
    def xlsx_blob_(self, request):
        return instance_mock(request, bytes)


class Describe_BaseChartData(object):

    def it_can_generate_chart_part_XML_for_its_data(self, xml_bytes_fixture):
        chart_data, chart_type_, ChartXmlWriter_, expected_bytes = (
            xml_bytes_fixture
        )
        xml_bytes = chart_data.xml_bytes(chart_type_)

        ChartXmlWriter_.assert_called_once_with(chart_type_, chart_data)
        assert xml_bytes == expected_bytes

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def xml_bytes_fixture(self, chart_type_, ChartXmlWriter_):
        chart_data = _BaseChartData()
        expected_bytes = 'ƒøØßår'.encode('utf-8')
        return chart_data, chart_type_, ChartXmlWriter_, expected_bytes

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ChartXmlWriter_(self, request):
        ChartXmlWriter_ = class_mock(
            request, 'pptx.chart.data.ChartXmlWriter'
        )
        ChartXmlWriter_.return_value.xml = 'ƒøØßår'
        return ChartXmlWriter_

    @pytest.fixture
    def chart_type_(self, request):
        return instance_mock(request, EnumValue)


class DescribeCategoryChartData(object):

    def it_knows_the_categories_range_ref(self, categories_ref_fixture):
        chart_data, expected_value = categories_ref_fixture
        assert chart_data.categories_ref == expected_value

    def it_knows_the_values_range_ref_for_a_series(self, values_ref_fixture):
        chart_data, workbook_writer_, series_, values_ref_ = (
            values_ref_fixture
        )
        values_ref = chart_data.values_ref(series_)
        workbook_writer_.values_ref.assert_called_once_with(series_)
        assert values_ref is values_ref_

    def it_provides_access_to_its_categories(self, categories_fixture):
        chart_data, Categories_, categories_ = categories_fixture
        categories = chart_data.categories
        Categories_.assert_called_once_with()
        assert categories is categories_

    def it_can_add_a_category(self, add_cat_fixture):
        chart_data, name, categories_, category_ = add_cat_fixture
        category = chart_data.add_category(name)
        categories_.add_category.assert_called_once_with(name)
        assert category is category_

    def it_can_add_a_series(self, add_ser_fixture):
        chart_data, name, values, number_format = add_ser_fixture[:4]
        CategorySeriesData_, calls, series_ = add_ser_fixture[4:]
        series = chart_data.add_series(name, values, number_format)
        CategorySeriesData_.assert_called_once_with(
            name, number_format, chart_data
        )
        assert chart_data[-1] is series
        assert series.add_data_point.call_args_list == calls
        assert series is series_

    def it_can_set_its_categories(self, categories_set_fixture):
        chart_data, names, Categories_, categories_, calls = (
            categories_set_fixture
        )
        chart_data.categories = names
        Categories_.assert_called_once_with()
        assert categories_.add_category.call_args_list == calls
        assert chart_data._categories is categories_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_cat_fixture(self, categories_prop_, categories_, category_):
        chart_data = CategoryChartData()
        name = 'foobar'
        categories_.add_category.return_value = category_
        return chart_data, name, categories_, category_

    @pytest.fixture
    def add_ser_fixture(self, CategorySeriesData_, series_):
        chart_data = CategoryChartData()
        name, values, number_format = 'foobar', iter((1, 2, 3)), '0.0'
        calls = [call(1), call(2), call(3)]
        return (
            chart_data, name, values, number_format, CategorySeriesData_,
            calls, series_
        )

    @pytest.fixture
    def categories_fixture(self, Categories_, categories_):
        chart_data = CategoryChartData()
        return chart_data, Categories_, categories_

    @pytest.fixture
    def categories_ref_fixture(
            self, _workbook_writer_prop_, workbook_writer_):
        chart_data = CategoryChartData()
        expected_value = categories_ref = 'Sheet42!$G$24'
        workbook_writer_.categories_ref = categories_ref
        return chart_data, expected_value

    @pytest.fixture
    def categories_set_fixture(self, Categories_, categories_):
        chart_data = CategoryChartData()
        names = iter(('a', 'b', 'c'))
        calls = [call('a'), call('b'), call('c')]
        return chart_data, names, Categories_, categories_, calls

    @pytest.fixture
    def values_ref_fixture(
            self, _workbook_writer_prop_, workbook_writer_, series_):
        chart_data = CategoryChartData()
        values_ref_ = 'Sheet1!$V$24'
        workbook_writer_.values_ref.return_value = values_ref_
        return chart_data, workbook_writer_, series_, values_ref_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Categories_(self, request, categories_):
        return class_mock(
            request, 'pptx.chart.data.Categories', return_value=categories_
        )

    @pytest.fixture
    def CategorySeriesData_(self, request, series_):
        return class_mock(
            request, 'pptx.chart.data.CategorySeriesData',
            return_value=series_
        )

    @pytest.fixture
    def categories_(self, request):
        return instance_mock(request, Categories)

    @pytest.fixture
    def categories_prop_(self, request, categories_):
        return property_mock(
            request, CategoryChartData, 'categories',
            return_value=categories_
        )

    @pytest.fixture
    def category_(self, request):
        return instance_mock(request, Category)

    @pytest.fixture
    def series_(self, request):
        return instance_mock(request, CategorySeriesData)

    @pytest.fixture
    def workbook_writer_(self, request):
        return instance_mock(request, CategoryWorkbookWriter)

    @pytest.fixture
    def _workbook_writer_prop_(self, request, workbook_writer_):
        return property_mock(
            request, CategoryChartData, '_workbook_writer',
            return_value=workbook_writer_
        )


class DescribeCategories(object):

    def it_knows_the_category_hierarchy_depth(self, depth_fixture):
        categories, expected_value = depth_fixture
        assert categories.depth == expected_value

    def it_knows_its_leaf_category_count(self, leaf_fixture):
        categories, expected_value = leaf_fixture
        assert categories.leaf_count == expected_value

    def it_knows_its_levels(self, levels_fixture):
        categories, expected_value = levels_fixture
        assert list(categories.levels) == expected_value

    def it_raises_on_category_depth_not_uniform(self, depth_raises_fixture):
        categories = depth_raises_fixture
        with pytest.raises(ValueError):
            categories.depth

    def it_can_add_a_category(self, add_fixture):
        categories, name, Category_, category_ = add_fixture
        category = categories.add_category(name)
        Category_.assert_called_once_with(name, categories)
        assert categories._categories[-1] is category
        assert category is category_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self, Category_, category_):
        categories = Categories()
        name = 'foobar'
        return categories, name, Category_, category_

    @pytest.fixture(params=[
        ((),        0),
        ((1,),      1),
        ((3,),      3),
        ((1, 1, 1), 1),
        ((3, 3, 3), 3),
    ])
    def depth_fixture(self, request):
        depths, expected_value = request.param
        categories = Categories()
        for depth in depths:
            categories._categories.append(
                instance_mock(request, Category, depth=depth)
            )
        return categories, expected_value

    @pytest.fixture(params=[
        ((),        0),
        ((1,),      1),
        ((1, 2, 3), 6),
    ])
    def leaf_fixture(self, request):
        leaf_counts, expected_value = request.param
        categories = Categories()
        for leaf_count in leaf_counts:
            categories._categories.append(
                instance_mock(request, Category, leaf_count=leaf_count)
            )
        return categories, expected_value

    @pytest.fixture(params=[
        ([(0, 'a', ()), (1, 'b', ())],
         [[(0, 'a'), (1, 'b')]]),
        ([(0, 'WEST', ((0, 'CA', ()), (1, 'LA', ()))),
          (2, 'EAST', ((2, 'NY', ()), (3, 'NJ', ())))],
         [[(0, 'CA'), (1, 'LA'), (2, 'NY'), (3, 'NJ')],
          [(0, 'WEST'), (2, 'EAST')]]),
    ])
    def levels_fixture(self, request):
        cat_data, expected_value = request.param
        categories = Categories()

        def iter_cats(cat_tree):
            for idx, cat_name, sub_cats in cat_tree:
                category_ = instance_mock(request, Category, idx=idx)
                category_.name = cat_name
                category_.sub_categories = list(iter_cats(sub_cats))
                yield category_

        categories._categories = list(iter_cats(cat_data))

        return categories, expected_value

    @pytest.fixture
    def depth_raises_fixture(self, request):
        categories = Categories()
        for depth in (1, 2, 1):
            categories._categories.append(
                instance_mock(request, Category, depth=depth)
            )
        return categories

    # fixture components ---------------------------------------------

    @pytest.fixture
    def Category_(self, request, category_):
        return class_mock(
            request, 'pptx.chart.data.Category', return_value=category_
        )

    @pytest.fixture
    def category_(self, request):
        return instance_mock(request, Category)


class DescribeCategory(object):

    def it_knows_its_depth(self, depth_fixture):
        category, expected_value = depth_fixture
        assert category.depth == expected_value

    def it_knows_its_leaf_category_count(self, leaf_fixture):
        category, expected_value = leaf_fixture
        assert category.leaf_count == expected_value

    def it_raises_on_depth_not_uniform(self, depth_raises_fixture):
        category = depth_raises_fixture
        with pytest.raises(ValueError):
            category.depth

    def it_knows_its_name(self, name_fixture):
        category, expected_value = name_fixture
        assert category.name == expected_value

    def it_provides_access_to_its_sub_categories(self, subs_fixture):
        category, sub_categories_ = subs_fixture
        assert category.sub_categories is sub_categories_

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        ((),        1),
        ((1,),      2),
        ((1, 1, 1), 2),
        ((2, 2, 2), 3),
    ])
    def depth_fixture(self, request):
        depths, expected_value = request.param
        category = Category(None, None)
        for depth in depths:
            category._sub_categories.append(
                instance_mock(request, Category, depth=depth)
            )
        return category, expected_value

    @pytest.fixture
    def depth_raises_fixture(self, request):
        category = Category(None, None)
        for depth in (1, 2, 1):
            category._sub_categories.append(
                instance_mock(request, Category, depth=depth)
            )
        return category

    @pytest.fixture(params=[
        ((),        1),
        ((1,),      1),
        ((1, 2, 3), 6),
    ])
    def leaf_fixture(self, request):
        leaf_counts, expected_value = request.param
        category = Category(None, None)
        for leaf_count in leaf_counts:
            category._sub_categories.append(
                instance_mock(request, Category, leaf_count=leaf_count)
            )
        return category, expected_value

    @pytest.fixture
    def name_fixture(self):
        name = 'foobar'
        category = Category(name, None)
        return category, name

    @pytest.fixture
    def subs_fixture(self):
        sub_categories_ = [42, 24]
        category = Category(None, None)
        category._sub_categories = sub_categories_
        return category, sub_categories_


class DescribeCategorySeriesData(object):

    def it_knows_the_categories_range_ref(self, categories_ref_fixture):
        series_data, expected_value = categories_ref_fixture
        assert series_data.categories_ref == expected_value

    def it_knows_its_values(self, values_fixture):
        series_data, expected_values = values_fixture
        assert series_data.values == expected_values

    def it_knows_its_values_range_ref(self, values_ref_fixture):
        series_data, chart_data_, values_ref_ = values_ref_fixture
        values_ref = series_data.values_ref
        chart_data_.values_ref.assert_called_once_with(series_data)
        assert values_ref is values_ref_

    def it_provides_access_to_the_chart_categories(self, categories_fixture):
        series_data, categories_ = categories_fixture
        assert series_data.categories is categories_

    def it_can_add_a_data_point(self, add_fixture):
        series_data, value, number_format = add_fixture[:3]
        CategoryDataPoint_, data_point_ = add_fixture[3:]
        data_point = series_data.add_data_point(value, number_format)
        CategoryDataPoint_.assert_called_once_with(value, number_format)
        assert series_data[-1] is data_point_
        assert data_point is data_point_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self, request, CategoryDataPoint_, data_point_):
        series_data = CategorySeriesData(None, None, None)
        value, number_format = 42, '0.0'
        return (
            series_data, value, number_format, CategoryDataPoint_,
            data_point_
        )

    @pytest.fixture
    def categories_fixture(self, chart_data_, categories_):
        series_data = CategorySeriesData(None, None, chart_data_)
        chart_data_.categories = categories_
        return series_data, categories_

    @pytest.fixture
    def categories_ref_fixture(self, chart_data_):
        series_data = CategorySeriesData(None, None, chart_data_)
        expected_value = categories_ref = 'Sheet1!$F$42'
        chart_data_.categories_ref = categories_ref
        return series_data, expected_value

    @pytest.fixture
    def values_fixture(self, request):
        series_data = CategorySeriesData(None, None, None)
        expected_values = [1, 2, 3]
        for value in expected_values:
            series_data._data_points.append(
                instance_mock(request, CategoryDataPoint, value=value)
            )
        return series_data, expected_values

    @pytest.fixture
    def values_ref_fixture(self, chart_data_):
        series_data = CategorySeriesData(None, None, chart_data_)
        values_ref_ = 'Sheet1!$V$42'
        chart_data_.values_ref.return_value = values_ref_
        return series_data, chart_data_, values_ref_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def CategoryDataPoint_(self, request, data_point_):
        return class_mock(
            request, 'pptx.chart.data.CategoryDataPoint',
            return_value=data_point_
        )

    @pytest.fixture
    def categories_(self, request):
        return instance_mock(request, Categories)

    @pytest.fixture
    def chart_data_(self, request):
        return instance_mock(request, CategoryChartData)

    @pytest.fixture
    def data_point_(self, request):
        return instance_mock(request, CategoryDataPoint)


class DescribeBubbleChartData(object):

    def it_can_add_a_series(self, add_series_fixture):
        chart_data, name, BubbleSeriesData_, series_data_ = add_series_fixture
        series_data = chart_data.add_series(name)
        BubbleSeriesData_.assert_called_once_with(chart_data, name)
        assert chart_data[-1] is series_data_
        assert series_data is series_data_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_series_fixture(self, request, BubbleSeriesData_, series_data_):
        chart_data = BubbleChartData()
        name = 'Series Name'
        return chart_data, name, BubbleSeriesData_, series_data_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BubbleSeriesData_(self, request, series_data_):
        return class_mock(
            request, 'pptx.chart.data.BubbleSeriesData',
            return_value=series_data_
        )

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, BubbleSeriesData)


class DescribeXyChartData(object):

    def it_can_add_a_series(self, add_series_fixture):
        chart_data, label, XySeriesData_, series_data_ = add_series_fixture
        series_data = chart_data.add_series(label)
        XySeriesData_.assert_called_once_with(chart_data, label)
        assert chart_data[-1] is series_data_
        assert series_data is series_data_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_series_fixture(self, request, XySeriesData_, series_data_):
        chart_data = XyChartData()
        label = 'Series Label'
        return chart_data, label, XySeriesData_, series_data_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def XySeriesData_(self, request, series_data_):
        return class_mock(
            request, 'pptx.chart.data.XySeriesData',
            return_value=series_data_
        )

    @pytest.fixture
    def series_data_(self, request):
        return instance_mock(request, XySeriesData)


class DescribeBubbleSeriesData(object):

    def it_can_add_a_data_point(self, add_data_point_fixture):
        series_data, x, y, size, BubbleDataPoint_, data_point_ = (
            add_data_point_fixture
        )
        data_point = series_data.add_data_point(x, y, size)
        BubbleDataPoint_.assert_called_once_with(x, y, size)
        assert series_data[-1] is data_point_
        assert data_point is data_point_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_data_point_fixture(self, request, BubbleDataPoint_, data_point_):
        series_data = BubbleSeriesData(None, None)
        x, y, size = 42, 24, 17
        return series_data, x, y, size, BubbleDataPoint_, data_point_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def BubbleDataPoint_(self, request, data_point_):
        return class_mock(
            request, 'pptx.chart.data.BubbleDataPoint',
            return_value=data_point_
        )

    @pytest.fixture
    def data_point_(self, request):
        return instance_mock(request, BubbleDataPoint)


class DescribeXySeriesData(object):

    def it_can_add_a_data_point(self, add_data_point_fixture):
        series_data, x, y, XyDataPoint_, data_point_ = add_data_point_fixture
        data_point = series_data.add_data_point(x, y)
        XyDataPoint_.assert_called_once_with(x, y)
        assert series_data[-1] is data_point_
        assert data_point is data_point_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_data_point_fixture(self, request, XyDataPoint_, data_point_):
        series_data = XySeriesData(None, None)
        x, y = 42, 24
        return series_data, x, y, XyDataPoint_, data_point_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def XyDataPoint_(self, request, data_point_):
        return class_mock(
            request, 'pptx.chart.data.XyDataPoint', return_value=data_point_
        )

    @pytest.fixture
    def data_point_(self, request):
        return instance_mock(request, XyDataPoint)


class DescribeCategoryDataPoint(object):

    def it_knows_its_value(self, value_fixture):
        data_point, expected_value = value_fixture
        assert data_point.value == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def value_fixture(self):
        value = 42
        data_point = CategoryDataPoint(value)
        return data_point, value
