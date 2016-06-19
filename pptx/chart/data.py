# encoding: utf-8

"""
ChartData and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence

from ..util import lazyproperty
from .xlsx import (
    BubbleWorkbookWriter, CategoryWorkbookWriter, XyWorkbookWriter
)
from .xmlwriter import ChartXmlWriter


class _BaseChartData(Sequence):
    """
    Base class providing common members for chart data objects. A chart data
    object serves as a proxy for the chart data table that will be written to
    an Excel worksheet; operating as a sequence of series as well as
    providing access to chart-level attributes. A chart data object is used
    as a parameter in :meth:`shapes.add_chart` and
    :meth:`Chart.replace_data`. The data structure varies between major chart
    categories such as category charts and XY charts.
    """
    def __init__(self, number_format='General'):
        super(_BaseChartData, self).__init__()
        self._number_format = number_format
        self._series = []

    def __getitem__(self, index):
        return self._series.__getitem__(index)

    def __len__(self):
        return self._series.__len__()

    def append(self, series):
        return self._series.append(series)

    def data_point_offset(self, series):
        """
        The total integer number of data points appearing in the series of
        this chart that are prior to *series* in this sequence.
        """
        count = 0
        for this_series in self:
            if series is this_series:
                return count
            count += len(this_series)
        raise ValueError('series not in chart data object')

    @property
    def number_format(self):
        """
        The formatting template string, e.g. '#,##0.0', that determines how
        X and Y values are formatted in this chart and in the Excel
        spreadsheet. A number format specified on a series will override this
        value for that series. Likewise, a distinct number format can be
        specified for a particular data point within a series.
        """
        return self._number_format

    def series_index(self, series):
        """
        Return the integer index of *series* in this sequence.
        """
        for idx, s in enumerate(self):
            if series is s:
                return idx
        raise ValueError('series not in chart data object')

    def series_name_ref(self, series):
        """
        Return the Excel worksheet reference to the cell containing the name
        for *series*.
        """
        return self._workbook_writer.series_name_ref(series)

    def x_values_ref(self, series):
        """
        The Excel worksheet reference to the X values for *series* (not
        including the column label).
        """
        return self._workbook_writer.x_values_ref(series)

    @property
    def xlsx_blob(self):
        """
        Return a blob containing an Excel workbook file populated with the
        contents of this chart data object.
        """
        return self._workbook_writer.xlsx_blob

    def xml_bytes(self, chart_type):
        """
        Return a blob containing the XML for a chart of *chart_type*
        containing the series in this chart data object, as bytes suitable
        for writing directly to a file.
        """
        return self._xml(chart_type).encode('utf-8')

    def y_values_ref(self, series):
        """
        The Excel worksheet reference to the Y values for *series* (not
        including the column label).
        """
        return self._workbook_writer.y_values_ref(series)

    @property
    def _workbook_writer(self):
        """
        The worksheet writer object to which layout and writing of the Excel
        worksheet for this chart will be delegated.
        """
        raise NotImplementedError('must be implemented by all subclasses')

    def _xml(self, chart_type):
        """
        Return (as unicode text) the XML for a chart of *chart_type*
        populated with the values in this chart data object. The XML is
        a complete XML document, including an XML declaration specifying
        UTF-8 encoding.
        """
        return ChartXmlWriter(chart_type, self).xml


class _BaseSeriesData(Sequence):
    """
    Base class providing common members for series data objects. A series
    data object serves as proxy for a series data column in the Excel
    worksheet. It operates as a sequence of data points, as well as providing
    access to series-level attributes like the series label.
    """
    def __init__(self, chart_data, name, number_format):
        super(_BaseSeriesData, self).__init__()
        self._chart_data = chart_data
        self._name = name
        self._number_format = number_format
        self._data_points = []

    def __getitem__(self, index):
        return self._data_points.__getitem__(index)

    def __len__(self):
        return self._data_points.__len__()

    def append(self, data_point):
        return self._data_points.append(data_point)

    @property
    def data_point_offset(self):
        """
        The integer count of data points that appear in all chart series
        prior to this one.
        """
        return self._chart_data.data_point_offset(self)

    @property
    def index(self):
        """
        Zero-based integer indicating the sequence position of this series in
        its chart. For example, the second of three series would return `1`.
        """
        return self._chart_data.series_index(self)

    @property
    def name(self):
        """
        The name of this series, e.g. 'Series 1'. This name is used as the
        column heading for the y-values of this series and may also appear in
        the chart legend and perhaps other chart locations.
        """
        return self._name

    @property
    def name_ref(self):
        """
        The Excel worksheet reference to the cell containing the name for
        this series.
        """
        return self._chart_data.series_name_ref(self)

    @property
    def number_format(self):
        """
        The formatting template string that determines how a number in this
        series is formatted, both in the chart and in the Excel spreadsheet;
        for example '#,##0.0'. If not specified for this series, it is
        inherited from the parent chart data object.
        """
        number_format = self._number_format
        if number_format is None:
            return self._chart_data.number_format
        return number_format

    @property
    def x_values(self):
        """
        A sequence containing the X value of each datapoint in this series,
        in data point order.
        """
        return [dp.x for dp in self._data_points]

    @property
    def x_values_ref(self):
        """
        The Excel worksheet reference to the X values for this chart (not
        including the column heading).
        """
        return self._chart_data.x_values_ref(self)

    @property
    def y_values(self):
        """
        A sequence containing the Y value of each datapoint in this series,
        in data point order.
        """
        return [dp.y for dp in self._data_points]

    @property
    def y_values_ref(self):
        """
        The Excel worksheet reference to the Y values for this chart (not
        including the column heading).
        """
        return self._chart_data.y_values_ref(self)


class _BaseDataPoint(object):
    """
    Base class providing common members for data point objects.
    """
    def __init__(self, series_data, number_format):
        super(_BaseDataPoint, self).__init__()
        self._series_data = series_data
        self._number_format = number_format

    @property
    def number_format(self):
        """
        The formatting template string that determines how the value of this
        data point is formatted, both in the chart and in the Excel
        spreadsheet; for example '#,##0.0'. If not specified for this data
        point, it is inherited from the parent series data object.
        """
        number_format = self._number_format
        if number_format is None:
            return self._series_data.number_format
        return number_format


class CategoryChartData(_BaseChartData):
    """
    A ChartData object suitable for use with category charts, all those
    having a discrete set of string values (categories) as the range of their
    independent axis (X-axis) values. Unlike the continuous ChartData types
    such as XyChartData, CategoryChartData has a single category sequence in
    lieu of X values for each data point and its data points have only the
    Y value.
    """
    def add_category(self, name):
        """
        Return a newly created |Category| object having *name* and appended
        to the end of the category sequence for this chart.
        """
        return self.categories.add_category(name)

    def add_series(self, name, values=(), number_format=None):
        """
        Add a series to this data set entitled *name* and having the data
        points specified by *values*, an iterable of numeric values.
        *number_format* specifies how the series values will be displayed,
        and may be a string, e.g. '#,##0', or an integer in the range 0-22 or
        37-49, signifying one of the built-in Excel number formats. The valid
        integer values and their meaning are documented on the
        :ref:`ExcelNumFormat` page.
        """
        series_data = CategorySeriesData(self, name, number_format)
        self.append(series_data)
        for value in values:
            series_data.add_data_point(value)
        return series_data

    @lazyproperty
    def categories(self):
        """
        A |Categories| object providing access to the hierarchy of category
        objects for this chart data. Assigning an iterable of category names
        (strings) replaces the |Categories| object with a new one containing
        a category for each name.
        """
        return Categories()

    @categories.setter
    def categories(self, category_names):
        categories = Categories()
        for name in category_names:
            categories.add_category(name)
        self._categories = categories

    @property
    def categories_ref(self):
        """
        The Excel worksheet reference to the categories for this chart (not
        including the column heading).
        """
        return self._workbook_writer.categories_ref

    def values_ref(self, series):
        """
        The Excel worksheet reference to the values for *series* (not
        including the column heading).
        """
        return self._workbook_writer.values_ref(series)

    @lazyproperty
    def _workbook_writer(self):
        """
        The worksheet writer object to which layout and writing of the Excel
        worksheet for this chart will be delegated.
        """
        return CategoryWorkbookWriter(self)


class Categories(Sequence):
    """
    A sequence of |Category| objects, also having certain hierarchical graph
    behaviors for support of multi-level (nested) categories.
    """
    def __init__(self):
        super(Categories, self).__init__()
        self._categories = []

    def __getitem__(self, idx):
        return self._categories.__getitem__(idx)

    def __len__(self):
        """
        Return the count of the highest level of category in this sequence.
        If it contains hierarchical (multi-level) categories, this number
        will differ from :attr:`category_count`, which is the number of leaf
        nodes.
        """
        return self._categories.__len__()

    def add_category(self, name):
        """
        Return a newly created |Category| object having *name* and appended
        to the end of this category sequence.
        """
        category = Category(name, self)
        self._categories.append(category)
        return category

    @property
    def depth(self):
        """
        The number of hierarchy levels in this category graph. Returns 0 if
        it contains no categories.
        """
        categories = self._categories
        if not categories:
            return 0
        first_depth = categories[0].depth
        for category in categories[1:]:
            if category.depth != first_depth:
                raise ValueError('category depth not uniform')
        return first_depth

    def index(self, category):
        """
        The offset of *category* in the overall sequence of leaf categories.
        A non-leaf category gets the index of its first sub-category.
        """
        index = 0
        for this_category in self._categories:
            if category is this_category:
                return index
            index += this_category.leaf_count
        raise ValueError('category not in top-level categories')

    @property
    def leaf_count(self):
        """
        The number of leaf-level categories in this hierarchy. The return
        value is the same as that of `len()` only when the hierarchy is
        single level.
        """
        return sum(c.leaf_count for c in self._categories)

    @property
    def levels(self):
        """
        A generator of (idx, name) sequences representing the category
        hierarchy from the bottom up. The first level contains all leaf
        categories, and each subsequent is the next level up.
        """
        def levels(categories):
            # yield all lower levels
            sub_categories = [
                sc for c in categories for sc in c.sub_categories
            ]
            if sub_categories:
                for level in levels(sub_categories):
                    yield level
            # yield this level
            yield [(cat.idx, cat.name) for cat in categories]

        for level in levels(self):
            yield level


class Category(object):
    """
    A chart category, primarily having a name to be displayed in the category
    axis, but also able to be configured in a hierarchy for support of
    multi-level category charts.
    """
    def __init__(self, name, parent):
        super(Category, self).__init__()
        self._name = name
        self._parent = parent
        self._sub_categories = []

    def add_sub_category(self, name):
        """
        Return a newly created |Category| object having *name* and appended
        to the end of the sub-category sequence for this category.
        """
        category = Category(name, self)
        self._sub_categories.append(category)
        return category

    @property
    def depth(self):
        """
        The number of hierarchy levels rooted at this category node. Returns
        1 if this category has no sub-categories.
        """
        sub_categories = self._sub_categories
        if not sub_categories:
            return 1
        first_depth = sub_categories[0].depth
        for category in sub_categories[1:]:
            if category.depth != first_depth:
                raise ValueError('category depth not uniform')
        return first_depth + 1

    @property
    def idx(self):
        """
        The offset of this category in the overall sequence of leaf
        categories. A non-leaf category gets the index of its first
        sub-category.
        """
        return self._parent.index(self)

    def index(self, sub_category):
        """
        The offset of *sub_category* in the overall sequence of leaf
        categories.
        """
        index = self._parent.index(self)
        for this_sub_category in self._sub_categories:
            if sub_category is this_sub_category:
                return index
            index += this_sub_category.leaf_count
        raise ValueError('sub_category not in this category')

    @property
    def leaf_count(self):
        """
        The number of leaf category nodes under this category. Returns
        1 if this category has no sub-categories.
        """
        if not self._sub_categories:
            return 1
        return sum(category.leaf_count for category in self._sub_categories)

    @property
    def name(self):
        """
        The string that appears on the axis for this category.
        """
        return self._name

    @property
    def sub_categories(self):
        """
        The sequence of child categories for this category.
        """
        return self._sub_categories


class ChartData(CategoryChartData):
    """
    Accumulates data specifying the categories and series values for a plot
    and acts as a proxy for the chart data table that will be written to an
    Excel worksheet. Used as a parameter in :meth:`shapes.add_chart` and
    :meth:`Chart.replace_data`.
    """


class CategorySeriesData(_BaseSeriesData):
    """
    The data specific to a particular category chart series. It provides
    access to the series label, the series data points, and an optional
    number format to be applied to each data point not having a specified
    number format.
    """
    def add_data_point(self, value, number_format=None):
        """
        Return a CategoryDataPoint object newly created with value *value*,
        an optional *number_format*, and appended to this sequence.
        """
        data_point = CategoryDataPoint(self, value, number_format)
        self.append(data_point)
        return data_point

    @property
    def categories(self):
        """
        The |Categories| object that provides access to the category objects
        for this series.
        """
        return self._chart_data.categories

    @property
    def categories_ref(self):
        """
        The Excel worksheet reference to the categories for this chart (not
        including the column heading).
        """
        return self._chart_data.categories_ref

    @property
    def values(self):
        """
        A sequence containing the (Y) value of each datapoint in this series,
        in data point order.
        """
        return [dp.value for dp in self._data_points]

    @property
    def values_ref(self):
        """
        The Excel worksheet reference to the (Y) values for this series (not
        including the column heading).
        """
        return self._chart_data.values_ref(self)


class XyChartData(_BaseChartData):
    """
    A specialized ChartData object suitable for use with an XY (aka. scatter)
    chart. Unlike ChartData, it has no category sequence. Rather, each data
    point of each series specifies both an X and a Y value.
    """
    def add_series(self, name, number_format=None):
        """
        Return an |XySeriesData| object newly created and added at the end of
        this sequence, identified by *name* and values formatted with
        *number_format*.
        """
        series_data = XySeriesData(self, name, number_format)
        self.append(series_data)
        return series_data

    @lazyproperty
    def _workbook_writer(self):
        """
        The worksheet writer object to which layout and writing of the Excel
        worksheet for this chart will be delegated.
        """
        return XyWorkbookWriter(self)


class BubbleChartData(XyChartData):
    """
    A specialized ChartData object suitable for use with a bubble chart.
    A bubble chart is essentially an XY chart where the markers are scaled to
    provide a third quantitative dimension to the exhibit.
    """
    def add_series(self, name, number_format=None):
        """
        Return a |BubbleSeriesData| object newly created and added at the end
        of this sequence, and having series named *name* and values formatted
        with *number_format*.
        """
        series_data = BubbleSeriesData(self, name, number_format)
        self.append(series_data)
        return series_data

    def bubble_sizes_ref(self, series):
        """
        The Excel worksheet reference for the range containing the bubble
        sizes for *series*.
        """
        return self._workbook_writer.bubble_sizes_ref(series)

    @lazyproperty
    def _workbook_writer(self):
        """
        The worksheet writer object to which layout and writing of the Excel
        worksheet for this chart will be delegated.
        """
        return BubbleWorkbookWriter(self)


class XySeriesData(_BaseSeriesData):
    """
    The data specific to a particular XY chart series. It provides access to
    the series label, the series data points, and an optional number format
    to be applied to each data point not having a specified number format.

    The sequence of data points in an XY series is significant; lines are
    plotted following the sequence of points, even if that causes a line
    segment to "travel backward" (implying a multi-valued function). The data
    points are not automatically sorted into increasing order by X value.
    """
    def add_data_point(self, x, y, number_format=None):
        """
        Return an XyDataPoint object newly created with values *x* and *y*,
        and appended to this sequence.
        """
        data_point = XyDataPoint(self, x, y, number_format)
        self.append(data_point)
        return data_point


class BubbleSeriesData(XySeriesData):
    """
    The data specific to a particular Bubble chart series. It provides access
    to the series label, the series data points, and an optional number
    format to be applied to each data point not having a specified number
    format.

    The sequence of data points in a bubble chart series is maintained
    throughout the chart building process because a data point has no unique
    identifier and can only be retrieved by index.
    """
    def add_data_point(self, x, y, size, number_format=None):
        """
        Append a new BubbleDataPoint object having the values *x*, *y*, and
        *size*. The optional *number_format* is used to format the Y value.
        If not provided, the number format is inherited from the series data.
        """
        data_point = BubbleDataPoint(self, x, y, size, number_format)
        self.append(data_point)
        return data_point

    @property
    def bubble_sizes(self):
        """
        A sequence containing the bubble size for each datapoint in this
        series, in data point order.
        """
        return [dp.bubble_size for dp in self._data_points]

    @property
    def bubble_sizes_ref(self):
        """
        The Excel worksheet reference for the range containing the bubble
        sizes for this series.
        """
        return self._chart_data.bubble_sizes_ref(self)


class CategoryDataPoint(_BaseDataPoint):
    """
    A data point in a category chart series. Provides access to the value of
    the datapoint and the number format with which it should appear in the
    Excel file.
    """
    def __init__(self, series_data, value, number_format):
        super(CategoryDataPoint, self).__init__(series_data, number_format)
        self._value = value

    @property
    def value(self):
        """
        The (Y) value for this category data point.
        """
        return self._value


class XyDataPoint(_BaseDataPoint):
    """
    A data point in an XY chart series. Provides access to the x and y values
    of the datapoint.
    """
    def __init__(self, series_data, x, y, number_format):
        super(XyDataPoint, self).__init__(series_data, number_format)
        self._x = x
        self._y = y

    @property
    def x(self):
        """
        The X value for this XY data point.
        """
        return self._x

    @property
    def y(self):
        """
        The Y value for this XY data point.
        """
        return self._y


class BubbleDataPoint(XyDataPoint):
    """
    A data point in a bubble chart series. Provides access to the x, y, and
    size values of the datapoint.
    """
    def __init__(self, series_data, x, y, size, number_format):
        super(BubbleDataPoint, self).__init__(series_data, x, y, number_format)
        self._size = size

    @property
    def bubble_size(self):
        """
        The value representing the size of the bubble for this data point.
        """
        return self._size
