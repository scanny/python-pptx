# encoding: utf-8

"""
Chart part objects, including Chart and Charts
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.package import XmlPart
from ..chart.chart import Chart
from ..util import lazyproperty


class ChartPart(XmlPart):
    """
    A chart part; corresponds to parts having partnames matching
    ppt/charts/chart[1-9][0-9]*.xml
    """
    partname_template = '/ppt/charts/chart%d.xml'

    @classmethod
    def new(cls, chart_type, chart_data, package):
        """
        Return a new |ChartPart| instance added to *package* containing
        a chart of *chart_type* and depicting *chart_data*.
        """
        chart_blob = chart_data.xml_bytes(chart_type)
        partname = package.next_partname(cls.partname_template)
        content_type = CT.DML_CHART
        chart_part = cls.load(partname, content_type, chart_blob, package)
        xlsx_blob = chart_data.xlsx_blob
        chart_part.chart_workbook.update_from_xlsx_blob(xlsx_blob)
        return chart_part

    @lazyproperty
    def chart(self):
        """
        The |Chart| object representing the chart in this part.
        """
        return Chart(self._element, self)

    @lazyproperty
    def chart_workbook(self):
        """
        The |ChartWorkbook| object providing access to the external chart
        data in a linked or embedded Excel workbook.
        """
        raise NotImplementedError


class ChartWorkbook(object):
    """
    Provides access to the external chart data in a linked or embedded Excel
    workbook.
    """
    def update_from_xlsx_blob(self, xlsx_blob):
        """
        Replace the Excel spreadsheet in the related |EmbeddedXlsxPart| with
        the Excel binary in *xlsx_blob*, adding a new |EmbeddedXlsxPart| if
        there isn't one.
        """
        raise NotImplementedError
