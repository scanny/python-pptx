# encoding: utf-8

"""
Embedded Package part objects, including EmbeddedPackagePart and
EmbeddedXlsxPart
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.package import Part


class EmbeddedPackagePart(Part):
    """
    A distinct OPC package, e.g. an Excel file, stored (embedded) in the main
    package as a part having partname like
    ``ppt/embeddings/Microsoft_Excel_Sheet1.xlsx``.
    """


class EmbeddedXlsxPart(EmbeddedPackagePart):
    """
    An Excel file stored in a part, typically used as a data source for
    a chart.
    """
    partname_template = '/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx'

    @classmethod
    def new(cls, xlsx_blob, package):
        """
        Return a new |EmbeddedXlsxPart| instance added to *package* and
        containing *xlsx_blob*.
        """
        partname = package.next_partname(cls.partname_template)
        content_type = CT.SML_SHEET
        xlsx_part = cls(partname, content_type, xlsx_blob, package)
        return xlsx_part
