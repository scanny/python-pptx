# pyright: reportPrivateUsage=false

"""Regression test for issue #579 — ``series.values`` on string-typed cache.

Issue #579 (https://github.com/scanny/python-pptx/issues/579) reported that
reading ``series.values`` raises ``ValueError`` when the chart XML contains
a ``c:pt/c:v`` whose text is not parseable as a float. Some authoring
tools emit string labels (or bare blanks that survive as non-numeric
text) inside a ``c:numCache`` even though the ECMA-376 schema calls for
``xsd:double``. PowerPoint happily renders such decks; python-pptx
should not crash on read.

The fix is lenient: when ``float(c:v.text)`` raises ``ValueError``, the
point is treated as a blank — the same way a missing ``c:pt`` element
or empty ``c:v`` is already handled — so iteration yields ``None`` for
that point instead of raising.
"""

from __future__ import annotations

from pptx.chart.series import BubbleSeries, XySeries, _BaseCategorySeries
from pptx.oxml.chart.series import CT_StrVal_NumVal_Composite

from .unitutil.cxml import element


class DescribeIssue579StringValueInNumCache(object):
    """Non-numeric ``c:v`` text in a numeric cache must not raise."""

    def it_returns_None_for_a_non_numeric_cached_value(self):
        # -- a real-world pattern: an author-side tool dropped a string
        # -- label ("N/A") into what is schema-wise a numeric cache.
        pt = element('c:pt{idx=0}/c:v"N/A"')
        assert isinstance(pt, CT_StrVal_NumVal_Composite)

        assert pt.value is None

    def it_returns_None_for_an_empty_v_element(self):
        # -- pre-existing behaviour, pinned here so the fix doesn't
        # -- regress the empty-text branch.
        pt = element("c:pt{idx=0}/c:v")
        assert pt.value is None

    def it_still_parses_a_normal_numeric_value(self):
        pt = element('c:pt{idx=0}/c:v"3.14"')
        assert pt.value == 3.14

    def it_yields_None_for_string_points_via_BaseCategorySeries_values(self):
        # -- reporter's failure mode: ``series.values`` raised on a
        # -- ``c:numCache`` with a string point. Now it yields ``None``.
        ser_cxml = (
            "c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},"
            'c:pt{idx=0}/c:v"1.1",c:pt{idx=1}/c:v"N/A",c:pt{idx=2}/c:v"3.3")'
        )
        series = _BaseCategorySeries(element(ser_cxml))

        assert series.values == (1.1, None, 3.3)

    def it_yields_None_for_string_points_via_XySeries_values(self):
        ser_cxml = (
            "c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=2},"
            'c:pt{idx=0}/c:v"2.5",c:pt{idx=1}/c:v"oops")'
        )
        series = XySeries(element(ser_cxml))

        assert series.values == (2.5, None)

    def it_yields_None_for_string_points_via_XySeries_x_values(self):
        ser_cxml = (
            "c:ser/c:xVal/c:numRef/c:numCache/(c:ptCount{val=2},"
            'c:pt{idx=0}/c:v"not-a-number",c:pt{idx=1}/c:v"7.0")'
        )
        series = XySeries(element(ser_cxml))

        assert series.x_values == (None, 7.0)

    def it_yields_None_for_string_points_via_BubbleSeries_sizes(self):
        ser_cxml = (
            "c:ser/c:bubbleSize/c:numRef/c:numCache/(c:ptCount{val=2},"
            'c:pt{idx=0}/c:v"bad",c:pt{idx=1}/c:v"4.2")'
        )
        series = BubbleSeries(element(ser_cxml))

        assert tuple(series.iter_bubble_sizes()) == (None, 4.2)
