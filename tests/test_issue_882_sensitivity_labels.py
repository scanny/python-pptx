"""Regression test for issue #882 — Microsoft Information Protection sensitivity labels.

Issue #882 asks how to set Microsoft Information Protection (MIP) / Azure Information
Protection sensitivity labels on a PowerPoint presentation so downstream consumers (DLP,
compliance, audit tooling) recognise the deck as labelled.

Sensitivity labels are not a first-class OOXML feature; they ride on top of the
``/docProps/custom.xml`` part as a conventional set of custom properties whose names
follow the ``MSIP_Label_<GUID>_<Field>`` pattern. The five canonical fields PowerPoint
(and the Azure Information Protection / MIP clients) emit per label are:

- ``MSIP_Label_<GUID>_Enabled``   — ``"true"`` / ``"false"`` (string, not bool)
- ``MSIP_Label_<GUID>_SetDate``   — ISO-8601 UTC timestamp string
- ``MSIP_Label_<GUID>_Method``    — ``"Standard"`` or ``"Privileged"`` (string)
- ``MSIP_Label_<GUID>_Name``      — human-readable label name (string)
- ``MSIP_Label_<GUID>_SiteId``    — tenant GUID (string)
- ``MSIP_Label_<GUID>_ContentBits`` — integer bit-field flagging where the label was
                                      applied (header/footer/watermark), stored as a
                                      string in the real wire format.

``Presentation.custom_properties`` (added in 2026.05.0 for issue #259) already provides
dict-like, typed access to ``/docProps/custom.xml``. The entire MSIP feature is
therefore a documentation / recipe deliverable — callers set and read the well-known
keys directly. These tests lock that contract in:

* Authoring every canonical MSIP field for a single label round-trips losslessly.
* The keys serialise to real ``<property>`` entries in ``docProps/custom.xml``.
* Multiple labels (distinct GUIDs) can coexist.
* Unlabelled presentations don't gain a custom-properties part (matching the #259
  contract — read-only access never forces a part into the package).
* Deleting a label strips every ``MSIP_Label_<GUID>_*`` key cleanly.
"""

from __future__ import annotations

import io
import zipfile

from pptx import Presentation


# -- a realistic MSIP GUID used by every scenario below (Azure-style, braces + upper) --
LABEL_GUID = "{defa4170-0d19-0005-0004-bc88714345d2}"
TENANT_SITE_ID = "{72f988bf-86f1-41af-91ab-2d7cd011db47}"


def _keys_for(guid: str) -> dict[str, str | int]:
    """Return the canonical 6-key MSIP bundle for `guid`."""
    prefix = "MSIP_Label_%s" % guid
    return {
        "%s_Enabled" % prefix: "true",
        "%s_SetDate" % prefix: "2025-05-02T09:15:00Z",
        "%s_Method" % prefix: "Standard",
        "%s_Name" % prefix: "Confidential",
        "%s_SiteId" % prefix: TENANT_SITE_ID,
        "%s_ContentBits" % prefix: 0,
    }


class DescribeIssue882SensitivityLabels(object):
    """End-to-end regression coverage for the MIP sensitivity-label recipe."""

    def it_round_trips_a_full_MSIP_label_bundle_through_save_and_reopen(self):
        prs = Presentation()
        bundle = _keys_for(LABEL_GUID)

        for name, value in bundle.items():
            prs.custom_properties[name] = value

        buf = io.BytesIO()
        prs.save(buf)

        # -- the custom-properties part is now materialised --
        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            assert "docProps/custom.xml" in zf.namelist()

        # -- reopen and confirm each MSIP key survived with its python type --
        buf.seek(0)
        reopened = Presentation(buf)
        cp = reopened.custom_properties
        for name, value in bundle.items():
            assert cp[name] == value, "roundtrip mismatch for %s" % name

    def it_serialises_MSIP_keys_as_vt_typed_property_entries(self):
        # -- Downstream DLP / compliance scanners parse ``docProps/custom.xml``
        # -- directly; pin the exact wire format so the recipe stays compatible.
        prs = Presentation()
        for name, value in _keys_for(LABEL_GUID).items():
            prs.custom_properties[name] = value

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        with zipfile.ZipFile(buf) as zf:
            custom_xml = zf.read("docProps/custom.xml").decode("utf-8")

        # -- each string-valued MSIP key surfaces as a <vt:lpwstr> child --
        assert 'name="MSIP_Label_%s_Name"' % LABEL_GUID in custom_xml
        assert "<vt:lpwstr>Confidential</vt:lpwstr>" in custom_xml
        assert "<vt:lpwstr>Standard</vt:lpwstr>" in custom_xml
        assert "<vt:lpwstr>true</vt:lpwstr>" in custom_xml
        assert "<vt:lpwstr>%s</vt:lpwstr>" % TENANT_SITE_ID in custom_xml
        # -- integer-valued MSIP key surfaces as a <vt:i4> child --
        assert "<vt:i4>0</vt:i4>" in custom_xml

    def it_supports_coexisting_MSIP_labels_for_distinct_GUIDs(self):
        # -- It is legal (and not uncommon) for a deck to carry more than one
        # -- MSIP_Label_<GUID>_* family, e.g. during migration between label
        # -- taxonomies. Ensure bundles for two distinct GUIDs coexist.
        other_guid = "{11111111-2222-3333-4444-555555555555}"
        prs = Presentation()

        for name, value in _keys_for(LABEL_GUID).items():
            prs.custom_properties[name] = value
        for name, value in _keys_for(other_guid).items():
            prs.custom_properties[name] = value

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)

        cp = reopened.custom_properties
        assert cp["MSIP_Label_%s_Name" % LABEL_GUID] == "Confidential"
        assert cp["MSIP_Label_%s_Name" % other_guid] == "Confidential"
        # -- 12 MSIP keys total (6 per label) ride on top of whatever was already there --
        msip_keys = [k for k in cp.keys() if k.startswith("MSIP_Label_")]
        assert len(msip_keys) == 12

    def it_does_not_create_a_custom_xml_part_for_an_unlabelled_deck(self):
        # -- Merely reading `custom_properties` must not force the part into
        # -- existence. A presentation that never sets any MSIP key should not
        # -- gain a ``/docProps/custom.xml`` member on save (preserves the
        # -- #259 lazy-materialisation contract).
        prs = Presentation()
        assert len(prs.custom_properties) == 0

        buf = io.BytesIO()
        prs.save(buf)

        buf.seek(0)
        with zipfile.ZipFile(buf) as zf:
            assert "docProps/custom.xml" not in zf.namelist()

    def it_removes_every_MSIP_key_when_the_label_is_cleared(self):
        # -- "Clearing a label" = deleting every MSIP_Label_<GUID>_* key.
        # -- Exercise the documented tear-down pattern end-to-end.
        prs = Presentation()
        bundle = _keys_for(LABEL_GUID)
        for name, value in bundle.items():
            prs.custom_properties[name] = value
        # -- an unrelated custom property must survive --
        prs.custom_properties["Department"] = "Engineering"

        prefix = "MSIP_Label_%s_" % LABEL_GUID
        for name in [k for k in prs.custom_properties.keys() if k.startswith(prefix)]:
            del prs.custom_properties[name]

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)

        cp = reopened.custom_properties
        msip_keys = [k for k in cp.keys() if k.startswith("MSIP_Label_")]
        assert msip_keys == []
        # -- the unrelated property is still there --
        assert cp["Department"] == "Engineering"

    def it_lets_user_code_enumerate_applied_labels_by_guid(self):
        # -- Pin the idiomatic "which labels does this deck carry?" helper
        # -- recipe — strip the leading/trailing fields and dedupe on the
        # -- GUID segment.
        prs = Presentation()
        for name, value in _keys_for(LABEL_GUID).items():
            prs.custom_properties[name] = value

        def applied_label_guids(pres):
            seen = set()
            for key in pres.custom_properties.keys():
                if not key.startswith("MSIP_Label_"):
                    continue
                # -- strip "MSIP_Label_" and "_<Field>" ends --
                tail = key[len("MSIP_Label_") :]
                guid, _, _field = tail.rpartition("_")
                if guid:
                    seen.add(guid)
            return sorted(seen)

        assert applied_label_guids(prs) == [LABEL_GUID]
