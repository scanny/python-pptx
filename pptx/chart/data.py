# encoding: utf-8

"""
ChartData and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from collections import Sequence
from xml.sax.saxutils import escape

from ..oxml import parse_xml
from ..oxml.ns import nsdecls
from ..util import lazyproperty
from .xlsx import BubbleWorkbookWriter, WorkbookWriter, XyWorkbookWriter
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
    def __init__(self):
        super(_BaseChartData, self).__init__()
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
    def __init__(self, chart_data, name):
        super(_BaseSeriesData, self).__init__()
        self._chart_data = chart_data
        self._name = name
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

    @property
    def name(self):
        """
        The string that appears on the axis for this category.
        """
        return self._name


class ChartData(object):
    """
    Accumulates data specifying the categories and series values for a plot
    and acts as a proxy for the chart data table that will be written to an
    Excel worksheet. Used as a parameter in :meth:`shapes.add_chart` and
    :meth:`Chart.replace_data`.
    """
    def __init__(self):
        super(ChartData, self).__init__()
        self._categories = []
        self._series_lst = []

    def add_series(self, name, values, number_format=0):
        """
        Add a series to this data set entitled *name* and having the data
        points specified by *values*, an iterable of numeric values.
        *num_fmt* specifies how the series values will be displayed, and may
        be a string, e.g. '#,##0', or an integer in the range 0-22 or 37-49,
        signifying one of the built-in Excel number formats. The valid
        integer values and their meaning are documented on the
        :ref:`ExcelNumFormat` page.
        """
        series_idx = len(self._series_lst)
        series = _SeriesData(
            series_idx, name, values, self._categories, number_format
        )
        self._series_lst.append(series)

    @property
    def categories(self):
        """
        Read-write. The sequence of category label strings to use in the
        chart. Any type that is iterable over a sequence of strings can be
        assigned, e.g. a list, tuple, or iterator.
        """
        return tuple(self._categories)

    @categories.setter
    def categories(self, categories):
        # Contents need to be replaced in-place so reference sent to
        # _SeriesData objects retain access to latest values
        self._categories[:] = categories

    @property
    def series(self):
        """
        An snapshot of the current |_SeriesData| objects in this chart data
        contained in an immutable sequence.
        """
        return tuple(self._series_lst)

    @property
    def xlsx_blob(self):
        """
        Return a blob containing an Excel workbook file populated with the
        categories and series in this chart data object.
        """
        return WorkbookWriter.xlsx_blob(self.categories, self._series_lst)

    def xml_bytes(self, chart_type):
        """
        Return a blob containing the XML for a chart of *chart_type*
        containing the series in this chart data object, as bytes suitable
        for writing directly to a file.
        """
        return self._xml(chart_type).encode('utf-8')

    def _xml(self, chart_type):
        """
        Return (as unicode text) the XML for a chart of *chart_type* having
        the categories and series in this chart data object. The XML is
        a complete XML document, including an XML declaration specifying
        UTF-8 encoding.
        """
        return ChartXmlWriter(chart_type, self._series_lst).xml


class XyChartData(_BaseChartData):
    """
    A specialized ChartData object suitable for use with an XY (aka. scatter)
    chart. Unlike ChartData, it has no category sequence. Rather, each data
    point of each series specifies both an X and a Y value.
    """
    def add_series(self, name):
        """
        Return an |XySeriesData| object newly created and added at the end of
        this sequence, and having series named *name*.
        """
        series_data = XySeriesData(self, name)
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
    def add_series(self, name):
        """
        Return a |BubbleSeriesData| object newly created and added at the end
        of this sequence, and having series named *name*.
        """
        series_data = BubbleSeriesData(self, name)
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
    def add_data_point(self, x, y):
        """
        Return an XyDataPoint object newly created with values *x* and *y*,
        and appended to this sequence.
        """
        data_point = XyDataPoint(x, y)
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
    def add_data_point(self, x, y, size):
        """
        Append a new BubbleDataPoint object having the values *x*, *y*, and
        *size*.
        """
        data_point = BubbleDataPoint(x, y, size)
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


class XyDataPoint(object):
    """
    A data point in an XY chart series. Provides access to the x and y values
    of the datapoint.
    """
    def __init__(self, x, y):
        super(XyDataPoint, self).__init__()
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
    def __init__(self, x, y, size):
        super(BubbleDataPoint, self).__init__(x, y)
        self._size = size

    @property
    def bubble_size(self):
        """
        The value representing the size of the bubble for this data point.
        """
        return self._size


class _SeriesData(object):
    """
    Like |ChartData|, a data transfer object, but specific to the data
    specifying a series. In addition, this object also provides XML
    generation for the ``<c:ser>`` element subtree.
    """
    def __init__(self, series_idx, name, values, categories, number_format):
        super(_SeriesData, self).__init__()
        self._series_idx = series_idx
        self._name = name
        self._values = values
        self._categories = categories
        self._number_format = number_format

    def __len__(self):
        """
        The number of values this series contains.
        """
        return len(self._values)

    @property
    def cat(self):
        """
        The ``<c:cat>`` element XML for this series, as an oxml element.
        """
        xml = self._cat_tmpl.format(
            wksht_ref=self._categories_ref, cat_count=len(self._categories),
            cat_pt_xml=self._cat_pt_xml, nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def cat_xml(self):
        """
        The unicode XML snippet for the ``<c:cat>`` element for this series,
        containing the category labels and spreadsheet reference.
        """
        return self._cat_tmpl.format(
            wksht_ref=self._categories_ref, cat_count=len(self._categories),
            cat_pt_xml=self._cat_pt_xml, nsdecls=''
        )

    @property
    def index(self):
        """
        The zero-based index of this series within the plot.
        """
        return self._series_idx

    @property
    def name(self):
        """
        The name of this series.
        """
        return self._name

    @property
    def number_format(self):
        """
        The Excel number format to be used to display the values in this
        series. May be a string such as '0.00%' or an integer specifying one
        of the built-in Excel number formats.
        """
        return self._number_format

    @property
    def tx(self):
        """
        Return a ``<c:tx>`` oxml element for this series, containing the
        series name.
        """
        name = escape(self.name)
        xml = self._tx_tmpl.format(
            wksht_ref=self._series_name_ref, series_name=name,
            nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def tx_xml(self):
        """
        Return the ``<c:tx>`` element for this series as unicode text. This
        element contains the series name.
        """
        name = escape(self.name)
        return self._tx_tmpl.format(
            wksht_ref=self._series_name_ref, series_name=name, nsdecls=''
        )

    @property
    def val(self):
        """
        The ``<c:val>`` XML for this series, as an oxml element.
        """
        xml = self._val_tmpl.format(
            wksht_ref=self._values_ref, val_count=len(self),
            val_pt_xml=self._val_pt_xml, nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def val_xml(self):
        """
        Return the unicode XML snippet for the ``<c:val>`` element describing
        this series.
        """
        return self._val_tmpl.format(
            wksht_ref=self._values_ref, val_count=len(self),
            val_pt_xml=self._val_pt_xml, nsdecls=''
        )

    @property
    def values(self):
        """
        The values in this series as a sequence of float.
        """
        return self._values

    @property
    def _categories_ref(self):
        """
        The Excel worksheet reference to the categories for this series.
        """
        end_row_number = len(self._categories) + 1
        return "Sheet1!$A$2:$A$%d" % end_row_number

    @property
    def _cat_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements containing the
        category names for this series.
        """
        xml = ''
        for idx, name in enumerate(self._categories):
            xml += (
                '                <c:pt idx="%d">\n'
                '                  <c:v>%s</c:v>\n'
                '                </c:pt>\n'
            ) % (idx, escape(str(name)))
        return xml

    @property
    def _cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series, containing
        the category labels and spreadsheet reference.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{cat_pt_xml}'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:cat>\n'
        )

    @property
    def _col_letter(self):
        """
        The letter of the Excel worksheet column in which the data for this
        series appears.
        """
        return chr(ord('B') + self._series_idx)

    @property
    def _series_name_ref(self):
        """
        The Excel worksheet reference to the name for this series.
        """
        return "Sheet1!$%s$1" % self._col_letter

    @property
    def _tx_tmpl(self):
        """
        The string formatting template for the ``<c:tx>`` element for this
        series, containing the series title and spreadsheet range reference.
        """
        return (
            '          <c:tx{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="1"/>\n'
            '                <c:pt idx="0">\n'
            '                  <c:v>{series_name}</c:v>\n'
            '                </c:pt>\n'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:tx>\n'
        )

    @property
    def _val_pt_xml(self):
        """
        The unicode XML snippet containing the ``<c:pt>`` elements for this
        series.
        """
        xml = ''
        for idx, value in enumerate(self._values):
            xml += (
                '                <c:pt idx="%d">\n'
                '                  <c:v>%s</c:v>\n'
                '                </c:pt>\n'
            ) % (idx, value)
        return xml

    @property
    def _val_tmpl(self):
        """
        The template for the ``<c:val>`` element for this series, containing
        the series values and their spreadsheet range reference.
        """
        return (
            '          <c:val{nsdecls}>\n'
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>General</c:formatCode>\n'
            '                <c:ptCount val="{val_count}"/>\n'
            '{val_pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
            '          </c:val>\n'
        )

    @property
    def _values_ref(self):
        """
        The Excel worksheet reference to the values for this series (not
        including the series name).
        """
        return "Sheet1!$%s$2:$%s$%d" % (
            self._col_letter, self._col_letter, len(self._values)+1
        )
