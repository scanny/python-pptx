"""Regression test for issue #259 — custom document properties (field codes).

Issue #259 asks for read/write access to the user-defined custom document properties
stored in ``/docProps/custom.xml``. These are the properties PowerPoint surfaces under
*File → Info → Properties → Advanced → Custom*, and they are the properties a Word- or
PowerPoint-level ``{ DOCPROPERTY }`` field code resolves against. They are distinct from
both the Dublin Core properties exposed by :attr:`Presentation.core_properties`
(``/docProps/core.xml``) and the application-level properties in ``/docProps/app.xml``
(handled by the Wave 2 #131 fix).

The feature is delivered by ``feat/issue-259-custom-document-properties`` which adds
:class:`pptx.parts.customprops.CustomPropertiesPart` (mirroring the
:class:`ExtendedPropertiesPart` pattern) and a dict-like ``Presentation.custom_properties``
facade. This regression test demonstrates the five supported value types (str, int,
float, bool, datetime) round-tripping through save + reopen, and exercises the core
``dict``-style operations the reporter asked for.
"""

from __future__ import annotations

import datetime as dt
import io
import zipfile

from pptx import Presentation


class DescribeIssue259CustomProperties(object):
    """End-to-end regression coverage for issue #259."""

    def it_round_trips_each_supported_value_type_through_save_and_reopen(self):
        prs = Presentation()

        prs.custom_properties["Department"] = "Engineering"
        prs.custom_properties["Revision"] = 7
        prs.custom_properties["Score"] = 98.6
        prs.custom_properties["Approved"] = True
        prs.custom_properties["Deadline"] = dt.datetime(2025, 3, 15, 12, 0, 0)

        buf = io.BytesIO()
        prs.save(buf)

        # -- the custom-properties part is now present in the zip --
        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            names = zf.namelist()
        assert "docProps/custom.xml" in names

        # -- reopen and confirm each value round-trips with its python type --
        buf.seek(0)
        reopened = Presentation(buf)
        cp = reopened.custom_properties
        assert cp["Department"] == "Engineering"
        assert isinstance(cp["Department"], str)
        assert cp["Revision"] == 7
        assert isinstance(cp["Revision"], int)
        assert cp["Score"] == 98.6
        assert isinstance(cp["Score"], float)
        assert cp["Approved"] is True
        assert cp["Deadline"] == dt.datetime(2025, 3, 15, 12, 0, 0)

    def it_does_not_create_a_custom_xml_part_unless_properties_are_set(self):
        prs = Presentation()

        # -- merely reading from `custom_properties` must not force the part into
        # -- existence; a freshly-created presentation with no custom props should
        # -- not gain a `/docProps/custom.xml` member on save.
        assert len(prs.custom_properties) == 0
        assert prs.custom_properties.keys() == []

        buf = io.BytesIO()
        prs.save(buf)

        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            assert "docProps/custom.xml" not in zf.namelist()

    def it_supports_delete_and_overwrite(self):
        prs = Presentation()
        prs.custom_properties["A"] = "initial"
        prs.custom_properties["B"] = 1
        # -- overwrite (same type) --
        prs.custom_properties["A"] = "updated"
        # -- overwrite (different type) --
        prs.custom_properties["B"] = "now-a-string"
        # -- delete --
        prs.custom_properties["C"] = 3
        del prs.custom_properties["C"]

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        cp = reopened.custom_properties

        assert cp["A"] == "updated"
        assert cp["B"] == "now-a-string"
        assert "C" not in cp

    def it_supports_dict_style_iteration_and_membership(self):
        prs = Presentation()
        prs.custom_properties.update(
            {"alpha": "a", "beta": "b", "gamma": "g"}
        )

        cp = prs.custom_properties
        assert "alpha" in cp
        assert "omega" not in cp
        assert set(cp.keys()) == {"alpha", "beta", "gamma"}
        assert dict(cp.items()) == {"alpha": "a", "beta": "b", "gamma": "g"}
        assert len(cp) == 3

    def it_stores_and_reloads_a_custom_property_the_user_can_reference_by_name(self):
        # -- this is the scenario driving the original request: a document author
        # -- wants a python-pptx script to plant a known-name property that an
        # -- in-slide `{ DOCPROPERTY MyProp }` field code can resolve against
        # -- after the deck is opened in PowerPoint.
        prs = Presentation()
        prs.custom_properties["MyProp"] = "custom-value-42"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        # -- the value is stored as a vt:lpwstr child under a <property name="MyProp">
        # -- in docProps/custom.xml — this is the exact wire format a PowerPoint
        # -- DOCPROPERTY field code reads from at render time.
        with zipfile.ZipFile(buf) as zf:
            custom_xml = zf.read("docProps/custom.xml").decode("utf-8")
        assert 'name="MyProp"' in custom_xml
        assert "<vt:lpwstr>custom-value-42</vt:lpwstr>" in custom_xml

        reopened = Presentation(buf)
        assert reopened.custom_properties["MyProp"] == "custom-value-42"
