"""Unit-test suite for `pptx.oxml.dml.shape_style` module."""

from __future__ import annotations

import pytest

from pptx.oxml.dml.shape_style import (
    CT_FontReference,
    CT_ShapeStyle,
    CT_StyleMatrixReference,
)

from ...unitutil.cxml import element


class DescribeCT_ShapeStyle(object):
    """Unit-test suite for `pptx.oxml.dml.shape_style.CT_ShapeStyle`."""

    def it_parses_a_p_style_element_to_the_custom_class(self):
        style = element(
            "p:style/(a:lnRef{idx=1},a:fillRef{idx=3},a:effectRef{idx=2}," "a:fontRef{idx=minor})"
        )
        assert isinstance(style, CT_ShapeStyle)

    def it_provides_typed_access_to_its_four_child_refs(self):
        style = element(
            "p:style/(a:lnRef{idx=1},a:fillRef{idx=3},a:effectRef{idx=2}," "a:fontRef{idx=major})"
        )
        assert isinstance(style.lnRef, CT_StyleMatrixReference)
        assert isinstance(style.fillRef, CT_StyleMatrixReference)
        assert isinstance(style.effectRef, CT_StyleMatrixReference)
        assert isinstance(style.fontRef, CT_FontReference)


class DescribeCT_StyleMatrixReference(object):
    """Unit-test suite for `pptx.oxml.dml.shape_style.CT_StyleMatrixReference`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_idx"),
        [
            ("a:lnRef{idx=0}", 0),
            ("a:lnRef{idx=1}", 1),
            ("a:fillRef{idx=7}", 7),
            ("a:effectRef{idx=2}", 2),
        ],
    )
    def it_reads_its_idx_as_an_int(self, cxml, expected_idx):
        ref = element(cxml)
        assert ref.idx == expected_idx


class DescribeCT_FontReference(object):
    """Unit-test suite for `pptx.oxml.dml.shape_style.CT_FontReference`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_idx"),
        [
            ("a:fontRef{idx=major}", "major"),
            ("a:fontRef{idx=minor}", "minor"),
            ("a:fontRef{idx=none}", "none"),
        ],
    )
    def it_reads_its_idx_as_the_collection_token(self, cxml, expected_idx):
        ref = element(cxml)
        assert ref.idx == expected_idx
