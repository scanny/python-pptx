# pyright: reportPrivateUsage=false

"""Regression test for issue #596 — AttributeError: 'NoneType' has no attribute 'idx'.

Issue #596 (https://github.com/scanny/python-pptx/issues/596) reports an
``AttributeError`` raised from
``_BaseSeriesXmlRewriter._add_cloned_sers`` at the line
``new_ser.idx.val = plotArea.next_idx``. The root cause is that
``plotArea.last_ser`` returned ``None`` — the last ``c:xChart`` element
in the plotArea had no ``c:ser`` children — so ``deepcopy(None)`` fed a
``None`` ``new_ser`` into the ``.idx.val`` assignment.

A related (though not the reporter's) failure mode is a ``c:ser`` that
is missing its (spec-required) ``c:idx`` or ``c:order`` child. The
oxml descriptor layer declares both as ``OneAndOnlyOne``, so accessing
``.idx`` or ``.order`` on such a series raises ``InvalidXmlError``
rather than the friendlier ``AttributeError`` the reporter saw — but it
still breaks series cloning in the wild.

These tests pin the fix: series cloning now (1) walks backwards through
``plotArea.xCharts`` to find a non-empty xChart when the last one is
empty, (2) synthesizes missing ``c:idx`` / ``c:order`` children on the
clone, and (3) raises a clear ``ValueError`` only when no xChart
anywhere in the plotArea contains a template ``c:ser``.
"""

from __future__ import annotations

import pytest

from pptx.chart.xmlwriter import _BaseSeriesXmlRewriter
from pptx.oxml import parse_xml

_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _plotArea(inner_xml):
    """Return a ``c:plotArea`` element wrapping *inner_xml*."""
    return parse_xml(
        '<c:plotArea xmlns:c="%s" xmlns:a="%s">%s</c:plotArea>' % (_CHART_NS, _DML_NS, inner_xml)
    )


class DescribeIssue596CloneSerNoIdx:
    """Issue #596 regression pins."""

    def it_clones_from_earlier_xChart_when_last_xChart_is_empty(self):
        """Reporter's exact failure mode.

        A plotArea whose final xChart has no ``c:ser`` children used to
        raise ``AttributeError: 'NoneType' object has no attribute 'idx'``
        from ``_add_cloned_sers``. Now it locates the last ``c:ser`` in
        an earlier xChart, clones it into the empty (last) xChart, and
        returns without error.
        """
        plotArea = _plotArea(
            "<c:barChart>"
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            "</c:barChart>"
            "<c:lineChart></c:lineChart>"
        )
        rewriter = _BaseSeriesXmlRewriter(None)

        rewriter._add_cloned_sers(plotArea, 1)

        lineChart = plotArea.xpath("./c:lineChart")[0]
        assert lineChart.xpath(
            "./c:ser"
        ), "expected the new c:ser to be appended into the last (previously empty) xChart"
        new_ser = lineChart.xpath("./c:ser")[0]
        assert new_ser.xpath("./c:idx/@val") == ["1"]
        assert new_ser.xpath("./c:order/@val") == ["1"]

    def it_clones_multiple_sers_into_last_empty_xChart(self):
        """Growing from 1 to 4 series when the last xChart is empty.

        All three new sers should land inside the (originally empty)
        last xChart, with ``c:idx`` / ``c:order`` values 1, 2, 3.
        """
        plotArea = _plotArea(
            "<c:barChart>"
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            "</c:barChart>"
            "<c:lineChart></c:lineChart>"
        )
        rewriter = _BaseSeriesXmlRewriter(None)

        rewriter._add_cloned_sers(plotArea, 3)

        lineChart = plotArea.xpath("./c:lineChart")[0]
        new_sers = lineChart.xpath("./c:ser")
        assert len(new_sers) == 3
        assert [s.xpath("./c:idx/@val")[0] for s in new_sers] == ["1", "2", "3"]
        assert [s.xpath("./c:order/@val")[0] for s in new_sers] == ["1", "2", "3"]
        # -- barChart is untouched --
        assert len(plotArea.xpath("./c:barChart/c:ser")) == 1

    def it_synthesizes_missing_idx_and_order_on_clones(self):
        """A template ser lacking c:idx / c:order is tolerated.

        The spec requires both, but PowerPoint sometimes emits decks
        that omit them; the clone has the children inserted and
        populated so the resulting XML is spec-valid.
        """
        plotArea = _plotArea(
            "<c:barChart>" "<c:ser>" "<c:tx><c:v>Series 1</c:v></c:tx>" "</c:ser>" "</c:barChart>"
        )
        rewriter = _BaseSeriesXmlRewriter(None)

        rewriter._add_cloned_sers(plotArea, 1)

        sers = plotArea.xpath("./c:barChart/c:ser")
        assert len(sers) == 2
        # -- original series left alone (still has no c:idx / c:order) --
        assert sers[0].xpath("./c:idx") == []
        assert sers[0].xpath("./c:order") == []
        # -- cloned series has both synthesized (defaulting to next_idx==0, next_order==0) --
        assert sers[1].xpath("./c:idx/@val") == ["0"]
        assert sers[1].xpath("./c:order/@val") == ["0"]

    def it_raises_value_error_when_plotArea_has_no_ser_anywhere(self):
        """A truly seriesless plotArea now raises a clear ValueError.

        Previously this path fell through to the ``None.idx`` access
        and raised the obscure ``AttributeError`` from the reporter.
        """
        plotArea = _plotArea("<c:barChart></c:barChart>")
        rewriter = _BaseSeriesXmlRewriter(None)

        with pytest.raises(ValueError, match="no existing series to clone from"):
            rewriter._add_cloned_sers(plotArea, 1)

    def it_tolerates_missing_idx_in_next_idx_computation(self):
        """``plotArea.next_idx`` skips sers that lack a c:idx child.

        Before the fix, the list-comp ``[s.idx.val for s in self.sers]``
        raised ``InvalidXmlError`` when any ser omitted ``c:idx``.
        """
        plotArea = _plotArea(
            "<c:barChart>"
            '<c:ser><c:idx val="3"/><c:order val="3"/></c:ser>'
            '<c:ser><c:order val="4"/></c:ser>'
            "</c:barChart>"
        )

        # -- next_idx skips the ser with no c:idx and returns max(3) + 1 --
        assert plotArea.next_idx == 4

    def it_tolerates_missing_order_in_next_order_computation(self):
        """``plotArea.next_order`` skips sers that lack a c:order child."""
        plotArea = _plotArea(
            "<c:barChart>"
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            '<c:ser><c:idx val="1"/></c:ser>'
            "</c:barChart>"
        )

        assert plotArea.next_order == 1

    def it_treats_empty_plotArea_last_ser_as_None(self):
        """``plotArea.last_ser`` is |None| only when no xChart has any ser."""
        plotArea = _plotArea("<c:barChart></c:barChart>" "<c:lineChart></c:lineChart>")
        assert plotArea.last_ser is None

    def it_returns_last_ser_from_earlier_xChart_when_last_xChart_empty(self):
        """``plotArea.last_ser`` walks back past empty xCharts."""
        plotArea = _plotArea(
            "<c:barChart>"
            '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
            '<c:ser><c:idx val="1"/><c:order val="1"/></c:ser>'
            "</c:barChart>"
            "<c:lineChart></c:lineChart>"
        )
        last_ser = plotArea.last_ser
        assert last_ser is not None
        assert last_ser.xpath("./c:idx/@val") == ["1"]
