# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.coreprops` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.oxml.coreprops import CT_CoreProperties
from pptx.parts.coreprops import CorePropertiesPart


class DescribeCorePropertiesPart(object):
    """Unit-test suite for `pptx.parts.coreprops.CorePropertiesPart` objects."""

    @pytest.mark.parametrize(
        ("prop_name", "expected_value"),
        [
            ("author", "python-pptx"),
            ("category", ""),
            ("comments", ""),
            ("content_status", "DRAFT"),
            ("identifier", "GXS 10.2.1ab"),
            ("keywords", "foo bar baz"),
            ("language", "US-EN"),
            ("last_modified_by", "Steve Canny"),
            ("subject", "Spam"),
            ("title", "Presentation"),
            ("version", "1.2.88"),
        ],
    )
    def it_knows_the_string_property_values(
        self, core_properties: CorePropertiesPart, prop_name: str, expected_value: str
    ):
        assert getattr(core_properties, prop_name) == expected_value

    @pytest.mark.parametrize(
        ("prop_name", "tagname", "value"),
        [
            ("author", "dc:creator", "scanny"),
            ("category", "cp:category", "silly stories"),
            ("comments", "dc:description", "Bar foo to you"),
            ("content_status", "cp:contentStatus", "FINAL"),
            ("identifier", "dc:identifier", "GT 5.2.xab"),
            ("keywords", "cp:keywords", "dog cat moo"),
            ("language", "dc:language", "GB-EN"),
            ("last_modified_by", "cp:lastModifiedBy", "Billy Bob"),
            ("subject", "dc:subject", "Eggs"),
            ("title", "dc:title", "Dissertation"),
            ("version", "cp:version", "81.2.8"),
        ],
    )
    def it_can_change_the_string_property_values(self, prop_name: str, tagname: str, value: str):
        coreProperties = self.coreProperties_xml(None, None)
        core_properties = CorePropertiesPart.load(None, None, None, coreProperties)  # type: ignore

        setattr(core_properties, prop_name, value)

        assert core_properties._element.xml == self.coreProperties_xml(tagname, value)

    @pytest.mark.parametrize(
        ("prop_name", "expected_value"),
        [
            ("created", dt.datetime(2012, 11, 17, 16, 37, 40)),
            ("last_printed", dt.datetime(2014, 6, 4, 4, 28)),
            ("modified", None),
        ],
    )
    def it_knows_the_date_property_values(
        self,
        core_properties: CorePropertiesPart,
        prop_name: str,
        expected_value: dt.datetime | None,
    ):
        actual_datetime = getattr(core_properties, prop_name)
        assert actual_datetime == expected_value

    @pytest.mark.parametrize(
        ("prop_name", "tagname", "value", "str_val", "attrs"),
        [
            (
                "created",
                "dcterms:created",
                dt.datetime(2001, 2, 3, 4, 5),
                "2001-02-03T04:05:00Z",
                ' xsi:type="dcterms:W3CDTF"',
            ),
            (
                "last_printed",
                "cp:lastPrinted",
                dt.datetime(2014, 6, 4, 4),
                "2014-06-04T04:00:00Z",
                "",
            ),
            (
                "modified",
                "dcterms:modified",
                dt.datetime(2005, 4, 3, 2, 1),
                "2005-04-03T02:01:00Z",
                ' xsi:type="dcterms:W3CDTF"',
            ),
        ],
    )
    def it_can_change_the_date_property_values(
        self, prop_name: str, tagname: str, value: dt.datetime, str_val: str, attrs: str
    ):
        coreProperties = self.coreProperties_xml(None, None)
        core_properties = CorePropertiesPart.load(None, None, None, coreProperties)  # type: ignore

        setattr(core_properties, prop_name, value)

        assert core_properties._element.xml == self.coreProperties_xml(tagname, str_val, attrs)

    @pytest.mark.parametrize(
        ("str_val", "expected_value"),
        [("42", 42), (None, 0), ("foobar", 0), ("-17", 0), ("32.7", 0)],
    )
    def it_knows_the_revision_number(self, str_val: str | None, expected_value: int):
        tagname = "" if str_val is None else "cp:revision"
        coreProperties = self.coreProperties_xml(tagname, str_val)
        core_properties = CorePropertiesPart.load(None, None, None, coreProperties)  # type: ignore

        assert core_properties.revision == expected_value

    def it_can_change_the_revision_number(self):
        coreProperties = self.coreProperties_xml(None, None)
        core_properties = CorePropertiesPart.load(None, None, None, coreProperties)  # type: ignore

        core_properties.revision = 42

        assert core_properties._element.xml == self.coreProperties_xml("cp:revision", "42")

    def it_can_construct_a_default_core_props(self):
        core_props = CorePropertiesPart.default(None)  # type: ignore
        # verify -----------------------
        assert isinstance(core_props, CorePropertiesPart)
        assert core_props.content_type is CT.OPC_CORE_PROPERTIES
        assert core_props.partname == "/docProps/core.xml"
        assert isinstance(core_props._element, CT_CoreProperties)
        assert core_props.title == "PowerPoint Presentation"
        assert core_props.last_modified_by == "python-pptx"
        assert core_props.revision == 1
        assert core_props.modified is not None
        # core_props.modified only stores time with seconds resolution, so
        # comparison needs to be a little loose (within two seconds)
        modified_timedelta = (
            dt.datetime.now(dt.timezone.utc).replace(tzinfo=None) - core_props.modified
        )
        max_expected_timedelta = dt.timedelta(seconds=2)
        assert modified_timedelta < max_expected_timedelta

    # -- fixtures ----------------------------------------------------

    def coreProperties_xml(self, tagname: str | None, str_val: str | None, attrs: str = "") -> str:
        tmpl = (
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/'
            'package/2006/metadata/core-properties" xmlns:dc="http://purl.or'
            'g/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype'
            '/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://'
            'www.w3.org/2001/XMLSchema-instance">%s</cp:coreProperties>\n'
        )
        if not tagname:
            child_element = ""
        elif not str_val:
            child_element = "\n  <%s%s/>\n" % (tagname, attrs)  # pragma: no cover
        else:
            child_element = "\n  <%s%s>%s</%s>\n" % (tagname, attrs, str_val, tagname)
        return tmpl % child_element

    @pytest.fixture
    def core_properties(self):
        xml = (
            b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
            b'\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.o'
            b'rg/package/2006/metadata/core-properties" xmlns:dc="http://pur'
            b'l.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcm'
            b'itype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="h'
            b'ttp://www.w3.org/2001/XMLSchema-instance">\n'
            b"  <cp:contentStatus>DRAFT</cp:contentStatus>\n"
            b"  <dc:creator>python-pptx</dc:creator>\n"
            b'  <dcterms:created xsi:type="dcterms:W3CDTF">2012-11-17T11:07:'
            b"40-05:30</dcterms:created>\n"
            b"  <dc:description/>\n"
            b"  <dc:identifier>GXS 10.2.1ab</dc:identifier>\n"
            b"  <dc:language>US-EN</dc:language>\n"
            b"  <cp:lastPrinted>2014-06-04T04:28:00Z</cp:lastPrinted>\n"
            b"  <cp:keywords>foo bar baz</cp:keywords>\n"
            b"  <cp:lastModifiedBy>Steve Canny</cp:lastModifiedBy>\n"
            b"  <cp:revision>4</cp:revision>\n"
            b"  <dc:subject>Spam</dc:subject>\n"
            b"  <dc:title>Presentation</dc:title>\n"
            b"  <cp:version>1.2.88</cp:version>\n"
            b"</cp:coreProperties>\n"
        )
        return CorePropertiesPart.load(None, None, None, xml)  # type: ignore
