# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.opc.flat_opc` module."""

from __future__ import annotations

import io

import pytest
from lxml import etree

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.flat_opc import FlatOpcWriter, _is_xml_content_type
from pptx.opc.package import Part, _Relationships
from pptx.opc.packuri import PackURI

from ..unitutil.mock import FixtureRequest, Mock, instance_mock


_PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"
_NSMAP = {"pkg": _PKG_NS}


class Describe_is_xml_content_type:
    """Unit-test suite for `pptx.opc.flat_opc._is_xml_content_type` helper."""

    @pytest.mark.parametrize(
        "content_type, expected",
        [
            (CT.PML_PRESENTATION_MAIN, True),
            (CT.OPC_RELATIONSHIPS, True),
            (CT.OFC_THEME, True),
            ("application/xml", True),
            ("text/xml", True),
            ("TEXT/XML", True),
            (CT.PNG, False),
            (CT.JPEG, False),
            (CT.X_FONT_TTF, False),
            ("application/octet-stream", False),
        ],
    )
    def it_recognizes_xml_media_types(self, content_type, expected):
        assert _is_xml_content_type(content_type) is expected


class DescribeFlatOpcWriter:
    """Unit-test suite for `pptx.opc.flat_opc.FlatOpcWriter` objects."""

    def it_emits_a_valid_flat_opc_document(self, request: FixtureRequest):
        pkg_rels = self._make_pkg_rels(
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="http://fake/doc" Target="foo.xml"/>'
            b"</Relationships>"
        )
        part_ = self._make_part(
            request,
            partname="/ppt/presentation.xml",
            content_type=CT.PML_PRESENTATION_MAIN,
            blob=b'<?xml version="1.0"?><p:presentation xmlns:p="pml"/>',
            rels_xml=None,
        )

        blob = FlatOpcWriter(pkg_rels, (part_,)).xml_bytes()

        # -- declaration + PI present --
        assert blob.startswith(
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<?mso-application progid="PowerPoint.Show"?>'
        )
        # -- parseable by lxml --
        root = etree.fromstring(blob)
        assert root.tag == "{%s}package" % _PKG_NS
        parts = root.findall("pkg:part", _NSMAP)
        assert len(parts) == 2
        # -- first part is the package-rels stream --
        assert parts[0].get("{%s}name" % _PKG_NS) == "/_rels/.rels"
        assert parts[0].get("{%s}contentType" % _PKG_NS) == CT.OPC_RELATIONSHIPS
        assert parts[0].find("pkg:xmlData", _NSMAP) is not None
        # -- second part is the presentation --
        assert parts[1].get("{%s}name" % _PKG_NS) == "/ppt/presentation.xml"
        assert parts[1].find("pkg:xmlData", _NSMAP) is not None

    def it_encodes_binary_parts_as_base64(self, request: FixtureRequest):
        pkg_rels = self._make_pkg_rels(b'<Relationships xmlns="urn:r"/>')
        binary_blob = b"\x89PNG\r\n\x1a\n" + b"\x00" * 16
        part_ = self._make_part(
            request,
            partname="/ppt/media/image1.png",
            content_type=CT.PNG,
            blob=binary_blob,
            rels_xml=None,
        )

        blob = FlatOpcWriter(pkg_rels, (part_,)).xml_bytes()
        root = etree.fromstring(blob)

        parts = root.findall("pkg:part", _NSMAP)
        image_part = [p for p in parts if p.get("{%s}name" % _PKG_NS).endswith("image1.png")][0]
        bin_data = image_part.find("pkg:binaryData", _NSMAP)
        assert bin_data is not None
        import base64

        assert base64.b64decode(bin_data.text) == binary_blob

    def it_emits_a_rels_part_when_a_part_has_relationships(self, request: FixtureRequest):
        pkg_rels = self._make_pkg_rels(b'<Relationships xmlns="urn:r"/>')
        part_rels_xml = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="http://fake/slide" Target="slide1.xml"/>'
            b"</Relationships>"
        )
        part_ = self._make_part(
            request,
            partname="/ppt/presentation.xml",
            content_type=CT.PML_PRESENTATION_MAIN,
            blob=b'<p:presentation xmlns:p="pml"/>',
            rels_xml=part_rels_xml,
        )

        blob = FlatOpcWriter(pkg_rels, (part_,)).xml_bytes()
        root = etree.fromstring(blob)

        names = [p.get("{%s}name" % _PKG_NS) for p in root.findall("pkg:part", _NSMAP)]
        # -- pkg-rels, part, part-rels, in that order --
        assert names == [
            "/_rels/.rels",
            "/ppt/presentation.xml",
            "/ppt/_rels/presentation.xml.rels",
        ]

    def but_it_omits_the_rels_part_when_a_part_has_no_relationships(
        self, request: FixtureRequest
    ):
        pkg_rels = self._make_pkg_rels(b'<Relationships xmlns="urn:r"/>')
        part_ = self._make_part(
            request,
            partname="/ppt/theme/theme1.xml",
            content_type=CT.OFC_THEME,
            blob=b'<a:theme xmlns:a="dml"/>',
            rels_xml=None,
        )

        blob = FlatOpcWriter(pkg_rels, (part_,)).xml_bytes()
        root = etree.fromstring(blob)

        names = [p.get("{%s}name" % _PKG_NS) for p in root.findall("pkg:part", _NSMAP)]
        assert names == ["/_rels/.rels", "/ppt/theme/theme1.xml"]

    def it_writes_the_flat_opc_xml_to_a_file_path(
        self, request: FixtureRequest, tmp_path
    ):
        pkg_rels = self._make_pkg_rels(b'<Relationships xmlns="urn:r"/>')
        part_ = self._make_part(
            request,
            partname="/ppt/presentation.xml",
            content_type=CT.PML_PRESENTATION_MAIN,
            blob=b'<p:presentation xmlns:p="pml"/>',
            rels_xml=None,
        )
        out_path = str(tmp_path / "out.xml")

        FlatOpcWriter.write(out_path, pkg_rels, (part_,))

        with open(out_path, "rb") as f:
            content = f.read()
        assert content.startswith(b"<?xml")
        assert b"<pkg:package" in content

    def and_it_writes_the_flat_opc_xml_to_a_stream(self, request: FixtureRequest):
        pkg_rels = self._make_pkg_rels(b'<Relationships xmlns="urn:r"/>')
        part_ = self._make_part(
            request,
            partname="/ppt/presentation.xml",
            content_type=CT.PML_PRESENTATION_MAIN,
            blob=b'<p:presentation xmlns:p="pml"/>',
            rels_xml=None,
        )
        buf = io.BytesIO()

        FlatOpcWriter.write(buf, pkg_rels, (part_,))

        buf.seek(0)
        data = buf.read()
        assert data.startswith(b"<?xml")
        assert b"<pkg:package" in data

    # -- fixture helpers ---------------------------------------------------

    def _make_pkg_rels(self, xml_bytes: bytes) -> Mock:
        rels = Mock(spec=_Relationships)
        rels.xml = xml_bytes
        return rels

    def _make_part(
        self,
        request: FixtureRequest,
        partname: str,
        content_type: str,
        blob: bytes,
        rels_xml: bytes | None,
    ) -> Mock:
        part_ = instance_mock(request, Part)
        part_.partname = PackURI(partname)
        part_.content_type = content_type
        part_.blob = blob
        # -- mimic the truthiness / `.rels.xml` contract FlatOpcWriter reads --
        if rels_xml is None:
            part_._rels = {}
            part_.rels = Mock(spec=_Relationships, xml=b"")
        else:
            part_._rels = {"rId1": Mock()}
            part_.rels = Mock(spec=_Relationships, xml=rels_xml)
        return part_


class DescribeFlatOpcRoundTrip:
    """End-to-end test emitting a Flat OPC document from a real |Presentation|."""

    def it_emits_a_parseable_flat_opc_for_a_default_presentation(self):
        """Sanity: default `Presentation()` saves as valid Flat OPC XML."""
        from pptx import Presentation

        prs = Presentation()
        buf = io.BytesIO()

        prs.save_flat_xml(buf)

        buf.seek(0)
        root = etree.parse(buf).getroot()
        assert root.tag == "{%s}package" % _PKG_NS
        # -- at least the package-rels and the presentation part are present --
        parts = root.findall("pkg:part", _NSMAP)
        names = {p.get("{%s}name" % _PKG_NS) for p in parts}
        assert "/_rels/.rels" in names
        assert "/ppt/presentation.xml" in names
        # -- content-type is carried on each pkg:part, not via [Content_Types].xml --
        assert "/[Content_Types].xml" not in names
        # -- at least one slide master is present --
        assert any(n.startswith("/ppt/slideMasters/slideMaster") for n in names)
