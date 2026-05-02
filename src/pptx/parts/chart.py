"""Chart part objects, including Chart and Charts."""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from pptx.chart.chart import Chart
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartRelationshipCloner, XmlPart
from pptx.parts.embeddedpackage import EmbeddedXlsxPart, clone_embedded_xlsx
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.package import Package


class ChartPart(XmlPart):
    """A chart part.

    Corresponds to parts having partnames matching ppt/charts/chart[1-9][0-9]*.xml
    """

    partname_template = "/ppt/charts/chart%d.xml"

    @classmethod
    def new(cls, chart_type: XL_CHART_TYPE, chart_data: ChartData, package: Package):
        """Return new |ChartPart| instance added to `package`.

        Returned chart-part contains a chart of `chart_type` depicting `chart_data`.
        """
        chart_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_CHART,
            package,
            chart_data.xml_bytes(chart_type),
        )
        chart_part.chart_workbook.update_from_xlsx_blob(chart_data.xlsx_blob)
        return chart_part

    @classmethod
    def clone_from(cls, source_chart_part: ChartPart, package: Package) -> ChartPart:
        """Return a new |ChartPart| in `package` that duplicates `source_chart_part`.

        Implements the cross-slide chart-copy primitive (issue #877). The
        returned chart part is a deep copy of the source: its ``c:chartSpace``
        element is cloned (with all ``r:id`` / ``r:embed`` / ``r:link``
        attributes rewritten against freshly-allocated relationships on the
        new part) and its embedded ``.xlsx`` workbook is cloned into a
        *distinct* :class:`EmbeddedXlsxPart` owned by `package`, so each
        chart continues to own its own "Edit Data" workbook.

        `package` is normally the package the caller will eventually relate
        the new chart into (same-package for an intra-presentation copy,
        cross-package for copying a chart between two open presentations).
        In the same-package case, :class:`PartRelationshipCloner` reuses
        non-xlsx relationship targets (images, theme-override, ...) because
        they already live in the package; only the embedded workbook is
        forcibly duplicated. In the cross-package case, the cloner
        materialises non-xlsx targets in the destination package so no
        cross-package references escape.
        """
        # -- Seed a placeholder chart part; `PartRelationshipCloner.clone` below
        # -- produces the real (rId-remapped) chartSpace element and replaces it. --
        partname = package.next_partname(cls.partname_template)
        placeholder_cs = copy.deepcopy(source_chart_part._element)
        new_chart_part = cls(partname, CT.DML_CHART, package, placeholder_cs)

        # -- Delegate per-rId cloning of non-xlsx relationships to F1. This walks
        # -- every `r:id` / `r:embed` / `r:link` on the source chartSpace, creates
        # -- a matching relationship on `new_chart_part` (reusing existing target
        # -- parts for same-package clones, materialising fresh duplicates for
        # -- cross-package clones), and rewrites the rId attributes on the clone. --
        new_chart_part._element = PartRelationshipCloner.clone(
            source_chart_part, new_chart_part, source_chart_part._element
        )

        # -- Explicitly clone the embedded workbook into a fresh EmbeddedXlsxPart
        # -- owned by the target and rewire `c:externalData/@r:id` to point at it.
        # -- Sharing an xlsx part across charts breaks PowerPoint's "Edit Data"
        # -- dialog, so we must override `PartRelationshipCloner`'s same-package
        # -- "reuse" semantics for this specific relationship (see #877 / F5). --
        clone_embedded_xlsx(source_chart_part, new_chart_part)

        return new_chart_part

    @lazyproperty
    def chart(self):
        """|Chart| object representing the chart in this part."""
        return Chart(self._element, self)

    @lazyproperty
    def chart_workbook(self):
        """
        The |ChartWorkbook| object providing access to the external chart
        data in a linked or embedded Excel workbook.
        """
        return ChartWorkbook(self._element, self)


class ChartWorkbook(object):
    """Provides access to external chart data in a linked or embedded Excel workbook."""

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
            self.xlsx_part = EmbeddedXlsxPart.new(xlsx_blob, self._chart_part.package)
            return
        xlsx_part.blob = xlsx_blob

    @property
    def xlsx_part(self):
        """Optional |EmbeddedXlsxPart| object containing data for this chart.

        This related part has its rId at `c:chartSpace/c:externalData/@rId`. This value
        is |None| if there is no `<c:externalData>` element, or if the rId referenced
        by that element is not present in the chart-part's relationships (e.g. the
        chart was pasted from a pre-2007 `.xls` workbook, the embedded-package rel was
        stripped by another client, or the chart uses externally-linked data). Prior
        to the fix for issue #490, a stale rId caused ``KeyError`` to propagate out of
        :meth:`Chart.replace_data`.
        """
        xlsx_part_rId = self._chartSpace.xlsx_part_rId
        if xlsx_part_rId is None:
            return None
        try:
            return self._chart_part.related_part(xlsx_part_rId)
        except KeyError:
            return None

    @xlsx_part.setter
    def xlsx_part(self, xlsx_part):
        """
        Set the related |EmbeddedXlsxPart| to *xlsx_part*. Assume one does
        not already exist.
        """
        rId = self._chart_part.relate_to(xlsx_part, RT.PACKAGE)
        externalData = self._chartSpace.get_or_add_externalData()
        externalData.rId = rId
