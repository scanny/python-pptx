# encoding: utf-8

"""Test suite for pptx.oxml.coreprops module."""

from __future__ import absolute_import

from datetime import datetime

from hamcrest import assert_that, equal_to, is_

from pptx.oxml.coreprops import CT_CoreProperties

from .unitdata.coreprops import a_coreProperties
from ..unitutil import TestCase


class TestCT_CoreProperties(TestCase):
    """Test CT_CoreProperties"""
    _cases = (
        ('author',            'Creator'),
        ('category',          'Category'),
        ('comments',          'Description'),
        ('content_status',    'Content Status'),
        ('identifier',        'Identifier'),
        ('keywords',          'Key Words'),
        ('language',          'Language'),
        ('last_modified_by',  'Last Modified By'),
        ('subject',           'Subject'),
        ('title',             'Title'),
        ('version',           'Version'),
    )

    def test_str_getters_are_empty_string_for_missing_element(self):
        """CT_CoreProperties str props empty str ('') for missing element"""
        # setup ------------------------
        childless_core_prop_builder = a_coreProperties()
        # verify -----------------------
        for attr_name, value in self._cases:
            childless_coreProperties = childless_core_prop_builder.element
            attr_value = getattr(childless_coreProperties, attr_name)
            reason = ("attr '%s' did not return '' for this XML:\n\n%s" %
                      (attr_name, childless_core_prop_builder.xml))
            assert_that(attr_value, is_(equal_to('')), reason)

    def test_str_getter_values_match_xml(self):
        """CT_CoreProperties string property values match parsed XML"""
        # verify -----------------------
        for attr_name, value in self._cases:
            builder = a_coreProperties().with_child(attr_name, value)
            coreProperties = builder.element
            attr_value = getattr(coreProperties, attr_name)
            reason = ("failed for property '%s' with this XML:\n\n%s" %
                      (attr_name, builder.xml))
            assert_that(attr_value, is_(equal_to(value)), reason)

    def test_str_setters_produce_correct_xml(self):
        """Assignment to CT_CoreProperties str property yields correct XML"""
        for attr_name, value in self._cases:
            # setup --------------------
            coreProperties = a_coreProperties().element  # no child elements
            # exercise -----------------
            setattr(coreProperties, attr_name, value)
            # verify -------------------
            expected_xml = a_coreProperties().with_child(attr_name, value).xml
            self.assertEqualLineByLine(expected_xml, coreProperties)

    def test_str_setter_raises_on_str_longer_than_255_chars(self):
        """Raises on assign len(str) > 255 to CT_CoreProperties str prop"""
        # setup ------------------------
        coreProperties = a_coreProperties().element
        # verify -----------------------
        with self.assertRaises(ValueError):
            coreProperties.comments = 'foobar 123 ' * 50

    def test_date_parser_recognizes_W3CDTF_strings(self):
        """date parser recognizes W3CDTF formatted date strings"""
        # valid W3CDTF date cases:
        # yyyy e.g. '2003'
        # yyyy-mm e.g. '2003-12'
        # yyyy-mm-dd e.g. '2003-12-31'
        # UTC timezone e.g. '2003-12-31T10:14:55Z'
        # numeric timezone e.g. '2003-12-31T10:14:55-08:00'
        cases = (
            ('1999', datetime(1999, 1, 1, 0, 0, 0)),
            ('2000-02', datetime(2000, 2, 1, 0, 0, 0)),
            ('2001-03-04', datetime(2001, 3, 4, 0, 0, 0)),
            ('2002-05-06T01:23:45Z', datetime(2002, 5, 6, 1, 23, 45)),
            ('2013-06-16T22:34:56-07:00', datetime(2013, 6, 17, 5, 34, 56)),
        )
        for dt_str, expected_datetime in cases:
            # exercise -----------------
            dt = CT_CoreProperties._parse_W3CDTF_to_datetime(dt_str)
            # verify -------------------
            assert_that(dt, is_(expected_datetime))

    def test_date_getters_have_none_on_element_not_present(self):
        """CT_CoreProperties date props are None where element not present"""
        # setup ------------------------
        coreProperties = a_coreProperties().element
        # verify -----------------------
        assert_that(coreProperties.created, is_(None))

    def test_date_getter_value_matches_xml(self):
        """CT_CoreProperties date props match parsed XML"""
        # setup ------------------------
        date_str = '1999-01-23T12:34:56Z'
        coreProperties = a_coreProperties().with_date_prop(
            'created', date_str).element
        # verify -----------------------
        expected_date = datetime(1999, 01, 23, 12, 34, 56)
        assert_that(coreProperties.created, is_(expected_date))

    def test_date_getters_have_none_on_not_datetime(self):
        """CT_CoreProperties date props are None for unparseable elm text"""
        # setup ------------------------
        date_str = 'foobar'
        core_props = a_coreProperties().with_date_prop(
            'created', date_str).element
        # verify -----------------------
        assert_that(core_props.created, is_(None))

    def test_date_setters_produce_correct_xml(self):
        """Assignment to CT_CoreProperties date props yields correct XML"""
        # See CT_CorePropertiesBuilder for how this implicitly tests that
        # 'created' and 'modified' add 'xsi:type="dcterms:W3CDTF"' attribute
        # by adding that on .with_date_prop for those properties.
        # setup ------------------------
        cases = ('created', 'last_printed', 'modified')
        value = datetime(2013, 6, 16, 12, 34, 56)
        for propname in cases:
            coreProperties = a_coreProperties().element  # no child elements
            # exercise -----------------
            setattr(coreProperties, propname, value)
            # verify -------------------
            expected_xml = a_coreProperties().with_date_prop(
                propname, '2013-06-16T12:34:56Z').xml
            self.assertEqualLineByLine(expected_xml, coreProperties)

    def test_date_setter_raises_on_not_datetime(self):
        """Raises on assign non-datetime to CT_CoreProperties date prop"""
        # setup ------------------------
        coreProperties = a_coreProperties().element
        # verify -----------------------
        with self.assertRaises(ValueError):
            coreProperties.created = 'foobar'

    def test_revision_is_zero_on_element_not_present(self):
        """CT_CoreProperties.revision is zero when element not present"""
        # setup ------------------------
        coreProperties = a_coreProperties().element
        # verify -----------------------
        assert_that(coreProperties.revision, is_(0))

    def test_revision_value_matches_xml(self):
        """CT_CoreProperties revision matches parsed XML"""
        # setup ------------------------
        builder = a_coreProperties().with_revision('9')
        coreProperties = builder.element
        # verify -----------------------
        reason = ("wrong revision returned for this XML:\n\n%s" %
                  (builder.xml))
        assert_that(coreProperties.revision, is_(9), reason)

    def test_revision_is_zero_on_invalid_element_text(self):
        """CT_CoreProperties.revision is zero if XML value is invalid"""
        # setup ------------------------
        cases = ('foobar', '-666')
        for invalid_text in cases:
            builder = a_coreProperties().with_revision(invalid_text)
            coreProperties = builder.element
            # verify -----------------------
            reason = ("wrong revision returned for this XML:\n\n%s" %
                      (builder.xml))
            assert_that(coreProperties.revision, is_(0), reason)

    def test_assign_to_revision_produces_correct_xml(self):
        """Assignment to CT_CoreProperties.revision yields correct XML"""
        # setup ------------------------
        coreProperties = a_coreProperties().element  # no child elements
        # exercise ---------------------
        coreProperties.revision = 999
        # verify -----------------------
        expected_xml = a_coreProperties().with_revision('999').xml
        self.assertEqualLineByLine(expected_xml, coreProperties)

    def test_assign_to_revision_raises_on_not_positive_int(self):
        """Raises on assign invalid value to CT_CoreProperties.revision"""
        # setup ------------------------
        cases = ('foobar', -666)
        # verify -----------------------
        for invalid_value in cases:
            coreProperties = a_coreProperties().element
            with self.assertRaises(ValueError):
                coreProperties.revision = invalid_value
