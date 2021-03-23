# encoding: utf-8

"""Embedded Package part objects.

"Package" in this context means another OPC package, i.e. a DOCX, PPTX, or XLSX "file".
"""

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part


class EmbeddedPackagePart(Part):
    """A distinct OPC package, e.g. an Excel file, embedded in this PPTX package.

    Has a partname like: `ppt/embeddings/Microsoft_Excel_Sheet1.xlsx`.
    """

    @classmethod
    def factory(cls, prog_id, object_blob, package):
        """Return a new |EmbeddedPackagePart| subclass instance added to *package*.

        The subclass is determined by `prog_id` which corresponds to the "application"
        used to open the "file-type" of `object_blob`. The returned part contains the
        bytes of `object_blob` and has the content-type also determined by `prog_id`.
        """
        raise NotImplementedError


class EmbeddedXlsxPart(EmbeddedPackagePart):
    """An Excel file stored in a part.

    This part-type arises as the data source for a chart, but may also be the OLE-object
    for an embedded object shape.
    """

    partname_template = "/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx"

    @classmethod
    def new(cls, xlsx_blob, package):
        """Return new |EmbeddedXlsxPart| object added to *package* from *xlsx_blob*."""
        return cls(
            package.next_partname(cls.partname_template),
            CT.SML_SHEET,
            xlsx_blob,
            package,
        )
