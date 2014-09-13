# encoding: utf-8

"""
ChartData and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..oxml import parse_xml
from ..oxml.ns import nsdecls
from .xlsx import WorkbookWriter
from .xmlwriter import ChartXmlWriter


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

    def add_series(self, name, values):
        """
        Add a series to this data set entitled *name* and the data points
        specified by *values*, an iterable of numeric values.
        """
        series_idx = len(self._series_lst)
        series = _SeriesData(series_idx, name, values, self._categories)
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


class _SeriesData(object):
    """
    Like |ChartData|, a data transfer object, but specific to the data
    specifying a series. In addition, this object also provides XML
    generation for the ``<c:ser>`` element subtree.
    """
    def __init__(self, series_idx, name, values, categories):
        super(_SeriesData, self).__init__()
        self._series_idx = series_idx
        self._name = name
        self._values = values
        self._categories = categories

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
    def tx(self):
        """
        Return a ``<c:tx>`` oxml element for this series, containing the
        series name.
        """
        xml = self._tx_tmpl.format(
            wksht_ref=self._series_name_ref, series_name=self.name,
            nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def tx_xml(self):
        """
        Return the ``<c:tx>`` element for this series as unicode text. This
        element contains the series name.
        """
        return self._tx_tmpl.format(
            wksht_ref=self._series_name_ref, series_name=self.name,
            nsdecls=''
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
            ) % (idx, name)
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
