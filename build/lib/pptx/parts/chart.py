# encoding: utf-8

"""
Chart part objects, including Chart and Charts
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..chart.chart import Chart
from .embeddedpackage import EmbeddedXlsxPart
from ..opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from ..opc.package import XmlPart
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
        return ChartWorkbook(self._element, self)


class ChartWorkbook(object):
    """
    Provides access to the external chart data in a linked or embedded Excel
    workbook.
    """
    def __init__(self, chartSpace, chart_part):
        super(ChartWorkbook, self).__init__()
        self._chartSpace = chartSpace
        self._chart_part = chart_part

    def update_from_xlsx_blob(self, xlsx_blob):
        """
        Replace the Excel spreadsheet in the related |EmbeddedXlsxPart| with
        the Excel binary in *xlsx_blob*, adding a new |EmbeddedXlsxPart| if
        there isn't one.
        """
        xlsx_part = self.xlsx_part
        if xlsx_part is None:
            self.xlsx_part = EmbeddedXlsxPart.new(xlsx_blob, self._package)
            return
        xlsx_part.blob = xlsx_blob

    @property
    def xlsx_part(self):
        """
        Return the related |EmbeddedXlsxPart| object having its rId at
        `c:chartSpace/c:externalData/@rId` or |None| if there is no
        `<c:externalData>` element.
        """
        xlsx_part_rId = self._chartSpace.xlsx_part_rId
        if xlsx_part_rId is None:
            return None
        return self._chart_part.related_parts[xlsx_part_rId]

    @xlsx_part.setter
    def xlsx_part(self, xlsx_part):
        """
        Set the related |EmbeddedXlsxPart| to *xlsx_part*. Assume one does
        not already exist.
        """
        rId = self._chart_part.relate_to(xlsx_part, RT.PACKAGE)
        externalData = self._chartSpace.get_or_add_externalData()
        externalData.rId = rId

    @property
    def _package(self):
        return self._chart_part.package
