"""Regression test for issue #1127 — ``ERCENT_40`` typo in ``MSO_PATTERN_TYPE``.

Issue #1127 (https://github.com/scanny/python-pptx/issues/1127) reports that
the ``MSO_PATTERN_TYPE`` enumeration (in :mod:`pptx.enum.dml`) exposes a
member named ``ERCENT_40`` — obviously a typo of ``PERCENT_40``. The
correctly-named member ``PERCENT_40`` is now the canonical name; the
original misspelling is preserved as a backwards-compatible alias so that
callers referencing ``MSO_PATTERN_TYPE.ERCENT_40`` continue to work.

Both names resolve to the same enum member, the same int value (``6``),
and the same OOXML attribute value (``"pct40"``).
"""

from __future__ import annotations

from pptx.enum.dml import MSO_PATTERN, MSO_PATTERN_TYPE


class DescribeIssue1127PercentForty:
    """Regression suite for issue #1127."""

    def it_exposes_PERCENT_40_as_the_canonical_member_name(self):
        assert hasattr(MSO_PATTERN_TYPE, "PERCENT_40")
        assert MSO_PATTERN_TYPE.PERCENT_40.name == "PERCENT_40"

    def it_preserves_ERCENT_40_as_a_backwards_compat_alias(self):
        # -- the misspelled name must continue to resolve --
        assert hasattr(MSO_PATTERN_TYPE, "ERCENT_40")

    def it_resolves_both_names_to_the_same_member(self):
        assert MSO_PATTERN_TYPE.ERCENT_40 is MSO_PATTERN_TYPE.PERCENT_40

    def it_assigns_the_same_integer_value_to_both_names(self):
        assert int(MSO_PATTERN_TYPE.PERCENT_40) == 6
        assert int(MSO_PATTERN_TYPE.ERCENT_40) == 6

    def it_maps_both_names_to_the_pct40_OOXML_attribute_value(self):
        assert MSO_PATTERN_TYPE.to_xml(MSO_PATTERN_TYPE.PERCENT_40) == "pct40"
        assert MSO_PATTERN_TYPE.to_xml(MSO_PATTERN_TYPE.ERCENT_40) == "pct40"

    def it_round_trips_pct40_back_to_PERCENT_40_by_canonical_name(self):
        member = MSO_PATTERN_TYPE.from_xml("pct40")
        assert member is MSO_PATTERN_TYPE.PERCENT_40
        # -- .name reports the canonical (non-alias) name --
        assert member.name == "PERCENT_40"

    def it_also_works_via_the_MSO_PATTERN_convenience_alias(self):
        assert MSO_PATTERN.PERCENT_40 is MSO_PATTERN_TYPE.PERCENT_40
        assert MSO_PATTERN.ERCENT_40 is MSO_PATTERN_TYPE.PERCENT_40
