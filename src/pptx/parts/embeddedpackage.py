"""Embedded Package part objects.

"Package" in this context means another OPC package, i.e. a DOCX, PPTX, or XLSX "file".
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part

if TYPE_CHECKING:
    from pptx.package import Package


# -- regex for validating a "safe" part-name file extension. Must match what would
# -- legally appear in a PackURI path segment; keep it ASCII letters/digits only.
_EXT_RE = re.compile(r"^[A-Za-z0-9]{1,8}$")


class EmbeddedPackagePart(Part):
    """A distinct OPC package, e.g. an Excel file, embedded in this PPTX package.

    Has a partname like: `ppt/embeddings/Microsoft_Excel_Sheet1.xlsx`.
    """

    @classmethod
    def factory(
        cls,
        prog_id: PROG_ID | str,
        object_blob: bytes,
        package: Package,
        extension: str | None = None,
    ):
        """Return a new |EmbeddedPackagePart| subclass instance added to *package*.

        The subclass is determined by `prog_id` which corresponds to the "application"
        used to open the "file-type" of `object_blob`. The returned part contains the
        bytes of `object_blob` and has the content-type also determined by `prog_id`.

        `extension` is an optional file-extension hint (without the leading dot, e.g.
        ``"zip"``) that controls the part-name of a generic OLE-object part. When
        `prog_id` is a member of |PROG_ID|, the extension carried by the enum member
        is preferred over this argument for non-Office types. When `prog_id` is a str
        and `extension` is omitted, the part-name defaults to ``.bin``.
        """
        # -- MS-Office "package" file-types have a distinct subclass with its own
        # -- content-type and part-name template.
        if isinstance(prog_id, PROG_ID) and prog_id.is_office_package:
            EmbeddedPartCls = {
                PROG_ID.DOCX: EmbeddedDocxPart,
                PROG_ID.PPTX: EmbeddedPptxPart,
                PROG_ID.XLSX: EmbeddedXlsxPart,
            }[prog_id]
            return EmbeddedPartCls.new(object_blob, package)

        # -- generic OLE-object: choose a part-name extension. Prefer the enum-member
        # -- extension over the caller-supplied hint; fall back to "bin" if neither is
        # -- provided or the hint fails validation.
        ext = prog_id.extension if isinstance(prog_id, PROG_ID) else extension
        if not ext or not _EXT_RE.match(ext):
            ext = "bin"
        partname_tmpl = f"/ppt/embeddings/oleObject%d.{ext}"
        return cls(
            package.next_partname(partname_tmpl),
            CT.OFC_OLE_OBJECT,
            package,
            object_blob,
        )

    @classmethod
    def new(cls, blob: bytes, package: Package):
        """Return new |EmbeddedPackagePart| subclass object.

        The returned part object contains `blob` and is added to `package`.
        """
        return cls(
            package.next_partname(cls.partname_template),
            cls.content_type,
            package,
            blob,
        )


class EmbeddedDocxPart(EmbeddedPackagePart):
    """A Word .docx file stored in a part.

    This part-type arises when a Word document appears as an embedded OLE-object shape.
    """

    partname_template = "/ppt/embeddings/Microsoft_Word_Document%d.docx"
    content_type = CT.WML_DOCUMENT


class EmbeddedPptxPart(EmbeddedPackagePart):
    """A PowerPoint file stored in a part.

    This part-type arises when a PowerPoint presentation (.pptx file) appears as an
    embedded OLE-object shape.
    """

    partname_template = "/ppt/embeddings/Microsoft_PowerPoint_Presentation%d.pptx"
    content_type = CT.PML_PRESENTATION


class EmbeddedXlsxPart(EmbeddedPackagePart):
    """An Excel file stored in a part.

    This part-type arises as the data source for a chart, but may also be the OLE-object
    for an embedded object shape.
    """

    partname_template = "/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx"
    content_type = CT.SML_SHEET


def clone_embedded_xlsx(source_chart_part, target_chart_part):
    """Copy `source_chart_part`'s embedded xlsx workbook to `target_chart_part`.

    Returns the new :class:`EmbeddedXlsxPart` attached to `target_chart_part`,
    or ``None`` when the source chart has no embedded workbook (e.g. the
    chart is linked to an external workbook or the source ``c:externalData``
    rel was dropped).

    This is the F5 cross-part helper used by chart-copy and combo-chart
    workflows: the raw xlsx bytes are duplicated into a *new* package part
    owned by `target_chart_part`'s package and then related back via a
    ``PACKAGE`` relationship. The source part is never reused across charts —
    PowerPoint expects each chart's ``c:externalData/@r:id`` to resolve to a
    part it alone owns, and sharing a single ``EmbeddedXlsxPart`` between
    charts breaks PowerPoint's "Edit Data" dialog (see downstream items
    #877 cross-slide chart copy and #239 replace-data preserving formulas).

    When `target_chart_part` already has an embedded workbook, the relation
    is replaced: the new part becomes the active embedded workbook, and the
    old relationship is overwritten on the existing ``c:externalData``
    element.
    """
    # -- reach through to the source's ChartWorkbook accessor rather than
    # -- calling `related_part` directly so that issue #490's "broken rel"
    # -- recovery path kicks in transparently: the source's workbook is
    # -- treated as absent when the rel is stale, and the clone becomes a
    # -- safe no-op. --
    source_workbook = source_chart_part.chart_workbook
    source_xlsx_part = source_workbook.xlsx_part
    if source_xlsx_part is None:
        return None
    new_part = EmbeddedXlsxPart.new(source_xlsx_part.blob, target_chart_part.package)
    target_workbook = target_chart_part.chart_workbook
    target_workbook.xlsx_part = new_part
    return new_part
