"""Unit-test suite for transparency functionality in `pptx.dml` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import ColorFormat, RGBColor
from pptx.dml.fill import FillFormat, _SolidFill
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL

from ..oxml.unitdata.dml import (
    a_alpha,
    a_solidFill,
    an_srgbClr,
    a_schemeClr,
)
from ..unitutil.cxml import element


class DescribeColorFormatTransparency(object):
    """Unit-test suite for ColorFormat transparency property."""

    def it_knows_its_transparency_value(self, transparency_get_fixture):
        color_format, expected_transparency = transparency_get_fixture
        assert color_format.transparency == expected_transparency

    def it_can_set_its_transparency_value(self, transparency_set_fixture):
        color_format, transparency, expected_xml = transparency_set_fixture
        color_format.transparency = transparency
        assert color_format._xFill.xml == expected_xml

    def it_raises_on_transparency_get_for_NoneColor(self, _NoneColor_color_format):
        # This should not raise - NoneColor should return 0.0 transparency
        transparency = _NoneColor_color_format.transparency
        assert transparency == 0.0

    def it_raises_on_transparency_set_for_NoneColor(self, _NoneColor_color_format):
        with pytest.raises(ValueError):
            _NoneColor_color_format.transparency = 0.5

    def it_raises_on_assign_invalid_transparency_value(self, rgb_color_format):
        color_format = rgb_color_format
        with pytest.raises(ValueError):
            color_format.transparency = 1.1
        with pytest.raises(ValueError):
            color_format.transparency = -0.1

    def it_can_set_transparency_to_zero_removes_alpha_element(self, rgb_color_format):
        """Setting transparency to 0.0 should remove the alpha element."""
        color_format = rgb_color_format
        # First set some transparency
        color_format.transparency = 0.5
        assert color_format._color._xClr.alpha is not None

        # Then set to 0.0 - should remove alpha element
        color_format.transparency = 0.0
        assert color_format._color._xClr.alpha is None
        assert color_format.transparency == 0.0

    def it_can_set_transparency_to_one_creates_zero_alpha(self, rgb_color_format):
        """Setting transparency to 1.0 should create alpha element with val=0."""
        color_format = rgb_color_format
        color_format.transparency = 1.0

        alpha_element = color_format._color._xClr.alpha
        assert alpha_element is not None
        assert alpha_element.val == 0.0
        assert color_format.transparency == 1.0

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            # no alpha element - fully opaque
            (lambda: an_srgbClr().with_val("FF0000"), 0.0),
            # alpha = 100000 (100% opaque) - 0% transparent
            (lambda: an_srgbClr().with_val("FF0000").with_child(
                a_alpha().with_val(100000)), 0.0),
            # alpha = 50000 (50% opaque) - 50% transparent
            (lambda: an_srgbClr().with_val("FF0000").with_child(
                a_alpha().with_val(50000)), 0.5),
            # alpha = 0 (0% opaque) - 100% transparent
            (lambda: an_srgbClr().with_val("FF0000").with_child(
                a_alpha().with_val(0)), 1.0),
        ]
    )
    def transparency_get_fixture(self, request):
        xClr_bldr_fn, expected_transparency = request.param
        xClr_bldr = xClr_bldr_fn()
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)
        return color_format, expected_transparency

    @pytest.fixture(
        params=[
            # Set to 0.0 (fully opaque) - should remove alpha element
            (lambda: an_srgbClr().with_val("FF0000"), 0.0,
             lambda: an_srgbClr().with_val("FF0000")),
            # Set to 0.5 (50% transparent) - should add alpha=50000
            (lambda: an_srgbClr().with_val("FF0000"), 0.5,
             lambda: an_srgbClr().with_val("FF0000").with_child(
                 a_alpha().with_val(50000))),
            # Set to 1.0 (fully transparent) - should add alpha=0
            (lambda: an_srgbClr().with_val("FF0000"), 1.0,
             lambda: an_srgbClr().with_val("FF0000").with_child(
                 a_alpha().with_val(0))),
            # Set from existing alpha to different value
            (lambda: an_srgbClr().with_val("FF0000").with_child(
                 a_alpha().with_val(75000)), 0.3,
             lambda: an_srgbClr().with_val("FF0000").with_child(
                 a_alpha().with_val(70000))),
        ]
    )
    def transparency_set_fixture(self, request):
        xClr_bldr_fn, transparency, expected_xClr_bldr_fn = request.param

        xClr_bldr = xClr_bldr_fn()
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)

        expected_xClr_bldr = expected_xClr_bldr_fn()
        expected_xml = a_solidFill().with_nsdecls().with_child(expected_xClr_bldr).xml()

        return color_format, transparency, expected_xml

    @pytest.fixture
    def rgb_color_format(self):
        solidFill = a_solidFill().with_nsdecls().with_child(
            an_srgbClr().with_val("FF0000")
        ).element
        return ColorFormat.from_colorchoice_parent(solidFill)

    @pytest.fixture
    def _NoneColor_color_format(self):
        solidFill = a_solidFill().with_nsdecls().element
        return ColorFormat.from_colorchoice_parent(solidFill)

    # Additional validation tests
    def it_validates_transparency_value_range(self):
        """Test that _validate_transparency_value works correctly."""
        from pptx.dml.color import ColorFormat

        # Create a dummy ColorFormat instance for testing validation
        xClr_bldr = an_srgbClr().with_val("FF0000")
        solidFill = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        color_format = ColorFormat.from_colorchoice_parent(solidFill)

        # Valid values should not raise
        color_format._validate_transparency_value(0.0)
        color_format._validate_transparency_value(0.5)
        color_format._validate_transparency_value(1.0)

        # Invalid values should raise ValueError
        with pytest.raises(ValueError, match="transparency must be number in range 0.0 to 1.0"):
            color_format._validate_transparency_value(-0.1)
        with pytest.raises(ValueError, match="transparency must be number in range 0.0 to 1.0"):
            color_format._validate_transparency_value(1.1)
        with pytest.raises(ValueError, match="transparency must be number in range 0.0 to 1.0"):
            color_format._validate_transparency_value(2.0)


class DescribeFillFormatTransparency(object):
    """Unit-test suite for FillFormat transparency property."""

    def it_delegates_transparency_to_solid_fill_object(self):
        """Test that FillFormat delegates transparency to its _SolidFill object."""
        # Create a _SolidFill directly
        xClr_bldr = an_srgbClr().with_val("FF0000")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element
        solid_fill = _SolidFill(solidFill_elm)

        # Create a FillFormat that wraps our _SolidFill
        from pptx.dml.fill import FillFormat

        # Create a simple mock parent
        class MockParent:
            def __init__(self):
                self.eg_fillProperties = solidFill_elm

        parent_mock = MockParent()
        fill_format = FillFormat(parent_mock, solid_fill)

        # Set transparency through FillFormat
        fill_format.transparency = 0.3

        # Verify it's delegated to the underlying solid fill
        assert abs(fill_format._fill.transparency - 0.3) < 0.001
        assert abs(fill_format.transparency - 0.3) < 0.001

    def it_raises_on_transparency_access_for_non_solid_fills(self):
        """Test that non-solid fills raise TypeError on transparency access."""
        from pptx.dml.fill import FillFormat, _NoFill

        # Create a FillFormat with NoFill
        class MockParent:
            def __init__(self):
                self.eg_fillProperties = None

        parent_mock = MockParent()
        no_fill = _NoFill(None)
        fill_format = FillFormat(parent_mock, no_fill)

        with pytest.raises(TypeError, match="fill type .* has no transparency"):
            fill_format.transparency

    def it_raises_on_transparency_set_for_non_solid_fills(self):
        """Test that non-solid fills raise TypeError on transparency set."""
        from pptx.dml.fill import FillFormat, _NoFill

        # Create a FillFormat with NoFill
        class MockParent:
            def __init__(self):
                self.eg_fillProperties = None

        parent_mock = MockParent()
        no_fill = _NoFill(None)
        fill_format = FillFormat(parent_mock, no_fill)

        with pytest.raises(TypeError, match="fill type .* has no transparency"):
            fill_format.transparency = 0.5


class DescribeSolidFillTransparency(object):
    """Unit-test suite for _SolidFill transparency property."""

    def it_provides_access_to_transparency(self, solid_fill_transparency_fixture):
        solid_fill, expected_transparency = solid_fill_transparency_fixture
        assert abs(solid_fill.transparency - expected_transparency) < 0.001

    def it_can_set_transparency(self, solid_fill_transparency_set_fixture):
        solid_fill, transparency = solid_fill_transparency_set_fixture
        solid_fill.transparency = transparency
        assert abs(solid_fill.transparency - transparency) < 0.001

    def it_delegates_transparency_to_fore_color(self, solid_fill_obj):
        """Test that _SolidFill delegates transparency to its fore_color."""
        solid_fill = solid_fill_obj

        # Set transparency through _SolidFill
        solid_fill.transparency = 0.6

        # Verify it's delegated to fore_color
        assert solid_fill.fore_color.transparency == 0.6
        assert solid_fill.transparency == 0.6

    def it_has_correct_fill_type(self, solid_fill_obj):
        """Test that _SolidFill reports correct fill type."""
        solid_fill = solid_fill_obj
        assert solid_fill.type == MSO_FILL.SOLID

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            # No alpha element - fully opaque
            (lambda: an_srgbClr().with_val("00FF00"), 0.0),
            # Alpha = 80000 (80% opaque) - 20% transparent
            (lambda: an_srgbClr().with_val("00FF00").with_child(
                a_alpha().with_val(80000)), 0.2),
            # Alpha = 30000 (30% opaque) - 70% transparent
            (lambda: an_srgbClr().with_val("00FF00").with_child(
                a_alpha().with_val(30000)), 0.7),
        ]
    )
    def solid_fill_transparency_fixture(self, request):
        xClr_bldr_fn, expected_transparency = request.param
        xClr_bldr = xClr_bldr_fn()
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        solid_fill = _SolidFill(solidFill_elm)
        return solid_fill, expected_transparency

    @pytest.fixture(
        params=[0.0, 0.1, 0.33, 0.67, 0.9, 1.0]
    )
    def solid_fill_transparency_set_fixture(self, request):
        transparency = request.param
        xClr_bldr = an_srgbClr().with_val("00FF00")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        solid_fill = _SolidFill(solidFill_elm)
        return solid_fill, transparency

    @pytest.fixture
    def solid_fill_obj(self):
        """Create a _SolidFill object for testing."""
        xClr_bldr = an_srgbClr().with_val("0000FF")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        return _SolidFill(solidFill_elm)


class DescribeTransparencyIntegration(object):
    """Integration tests for transparency across the entire stack."""

    def it_works_end_to_end_with_solid_fill_and_color_format(self):
        """Test transparency from ColorFormat through _SolidFill."""
        # Create a complete fill hierarchy using _SolidFill directly
        xClr_bldr = an_srgbClr().with_val("FF00FF")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        solid_fill = _SolidFill(solidFill_elm)

        # Test setting transparency at _SolidFill level
        solid_fill.transparency = 0.4

        # Verify it propagates through all levels
        assert abs(solid_fill.transparency - 0.4) < 0.001
        assert abs(solid_fill.fore_color.transparency - 0.4) < 0.001

        # Verify the underlying XML has the alpha element
        alpha_elm = solid_fill.fore_color._color._xClr.alpha
        assert alpha_elm is not None
        assert abs(alpha_elm.val - 0.6) < 0.001  # 1.0 - 0.4 = 0.6 (60% opaque)

    def it_handles_transparency_removal_correctly(self):
        """Test that setting transparency to 0.0 removes alpha element."""
        xClr_bldr = an_srgbClr().with_val("FFFF00").with_child(a_alpha().with_val(40000))
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        solid_fill = _SolidFill(solidFill_elm)

        # Verify initial transparency (40000/100000 = 0.4 opaque = 0.6 transparent)
        assert abs(solid_fill.transparency - 0.6) < 0.001
        assert solid_fill.fore_color._color._xClr.alpha is not None

        # Set transparency to 0.0
        solid_fill.transparency = 0.0

        # Verify alpha element is removed
        assert solid_fill.transparency == 0.0
        assert solid_fill.fore_color._color._xClr.alpha is None

    def it_handles_color_format_transparency_directly(self):
        """Test ColorFormat transparency manipulation directly."""
        xClr_bldr = an_srgbClr().with_val("00FFFF")
        solidFill_elm = a_solidFill().with_nsdecls().with_child(xClr_bldr).element

        color_format = ColorFormat.from_colorchoice_parent(solidFill_elm)

        # Test setting various transparency values
        test_values = [0.0, 0.25, 0.5, 0.75, 1.0]
        for transparency in test_values:
            color_format.transparency = transparency
            assert abs(color_format.transparency - transparency) < 0.001

            if transparency == 0.0:
                # No alpha element for fully opaque
                assert color_format._color._xClr.alpha is None
            else:
                # Alpha element should exist with correct value
                alpha_elm = color_format._color._xClr.alpha
                assert alpha_elm is not None
                expected_alpha = 1.0 - transparency
                assert abs(alpha_elm.val - expected_alpha) < 0.001