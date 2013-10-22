# encoding: utf-8

"""Test suite for pptx.coreprops module."""

from __future__ import absolute_import

from datetime import datetime, timedelta

from hamcrest import assert_that, instance_of, is_, less_than

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.oxml import CT_CoreProperties
from pptx.parts.coreprops import _CoreProperties

from ..testing import TestCase


class Test_CoreProperties(TestCase):
    """Test _CoreProperties"""
    def test_default_constructs_default_core_props(self):
        """_CoreProperties.default() returns new default core props part"""
        # exercise ---------------------
        core_props = _CoreProperties._default()
        # verify -----------------------
        assert_that(core_props, is_(instance_of(_CoreProperties)))
        assert_that(core_props._content_type, is_(CT.OPC_CORE_PROPERTIES))
        assert_that(core_props.partname, is_('/docProps/core.xml'))
        assert_that(core_props._element, is_(instance_of(CT_CoreProperties)))
        assert_that(core_props.title, is_('PowerPoint Presentation'))
        assert_that(core_props.last_modified_by, is_('python-pptx'))
        assert_that(core_props.revision, is_(1))
        # core_props.modified only stores time with seconds resolution, so
        # comparison needs to be a little loose (within two seconds)
        modified_timedelta = datetime.utcnow() - core_props.modified
        max_expected_timedelta = timedelta(seconds=2)
        assert_that(modified_timedelta, less_than(max_expected_timedelta))
