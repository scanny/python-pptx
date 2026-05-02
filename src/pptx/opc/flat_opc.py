"""Serialize an OPC package as a single Flat OPC (Office XML) document.

Flat OPC is the "single-file XML" wire format defined in ECMA-376 Part 4
(a.k.a. Microsoft's "XML Presentation" format in PowerPoint 2007+). The
whole package is represented as one XML document rooted at ``<pkg:package>``
with one ``<pkg:part>`` child per package item.

XML parts embed their payload verbatim inside ``<pkg:xmlData>``; binary parts
are base64-encoded inside ``<pkg:binaryData>``. The package and per-part
relationship streams are emitted as ordinary XML parts whose
``pkg:contentType`` is the OPC relationships content type.

No ``[Content_Types].xml`` part is emitted (each ``pkg:part`` carries its
content type as an attribute, which supersedes the content-types stream in
the Flat OPC representation).
"""

from __future__ import annotations

import base64
from typing import IO, TYPE_CHECKING, Iterable, Sequence

from lxml import etree

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PACKAGE_URI, PackURI

if TYPE_CHECKING:
    from pptx.opc.package import Part, _Relationships  # pyright: ignore[reportPrivateUsage]


# -- ECMA-376 Part 4 Flat OPC namespace --
_PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"

# -- Office progid processing-instruction for PowerPoint.  Without this PI
# -- PowerPoint still opens a .xml package but launches Word; adding it
# -- routes a double-click to PowerPoint (matches "Save As XML Presentation"
# -- behaviour).  This is the only reason we emit a PI.
_MSO_APPLICATION_PROGID = "PowerPoint.Show"


class FlatOpcWriter:
    """Emit a Flat OPC XML document for a package.

    Instantiate with the package-level relationships and the sequence of
    parts produced by :meth:`pptx.opc.package.OpcPackage.iter_parts`, then
    call :meth:`write` (or :meth:`xml_bytes` for the serialized bytes).
    """

    def __init__(self, pkg_rels: _Relationships, parts: Sequence[Part]):
        self._pkg_rels = pkg_rels
        self._parts = tuple(parts)

    @classmethod
    def write(
        cls,
        pkg_file: str | IO[bytes],
        pkg_rels: _Relationships,
        parts: Sequence[Part],
    ) -> None:
        """Serialize `pkg_rels` + `parts` to `pkg_file` as a Flat OPC XML document.

        `pkg_file` is either a filesystem path (``str``) or a file-like object
        opened for binary writing.
        """
        blob = cls(pkg_rels, parts).xml_bytes()
        if isinstance(pkg_file, str):
            with open(pkg_file, "wb") as f:
                f.write(blob)
        else:
            pkg_file.write(blob)

    def xml_bytes(self) -> bytes:
        """Return the Flat OPC document as XML bytes with declaration and PI."""
        root = self._build_root()
        # -- ``lxml.etree.tostring`` writes the XML declaration only when an
        # -- encoding is supplied; Flat OPC requires the standalone="yes" form.
        # -- The ``mso-application`` PI is attached to the ElementTree so that
        # -- lxml emits it between the declaration and the root element.
        tree = etree.ElementTree(root)
        root.addprevious(
            etree.ProcessingInstruction(
                "mso-application", 'progid="%s"' % _MSO_APPLICATION_PROGID
            )
        )
        return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)

    # -- internal -----------------------------------------------------------

    def _build_root(self) -> etree._Element:
        """Return the ``<pkg:package>`` root element with all part children."""
        nsmap = {"pkg": _PKG_NS}
        root = etree.Element("{%s}package" % _PKG_NS, nsmap=nsmap)

        for partname, content_type, blob in self._iter_part_entries():
            self._append_part(root, partname, content_type, blob)

        return root

    def _iter_part_entries(self) -> Iterable[tuple[PackURI, str, bytes]]:
        """Yield ``(partname, content_type, blob)`` triples in Flat OPC order.

        The sequence is: package relationships part first (``/_rels/.rels``),
        then each package part in caller-supplied order, with that part's
        relationship stream (if any) immediately following it.
        """
        # -- package rels --
        pkg_rels_blob = self._pkg_rels.xml  # includes XML declaration
        yield (PACKAGE_URI.rels_uri, CT.OPC_RELATIONSHIPS, pkg_rels_blob)

        for part in self._parts:
            yield (part.partname, part.content_type, part.blob)
            # -- per-part rels, only when the part has any relationships --
            if part._rels:  # pyright: ignore[reportPrivateUsage]
                yield (
                    part.partname.rels_uri,
                    CT.OPC_RELATIONSHIPS,
                    part.rels.xml,
                )

    def _append_part(
        self,
        root: etree._Element,
        partname: PackURI,
        content_type: str,
        blob: bytes,
    ) -> None:
        """Append one ``<pkg:part>`` child for `partname` to `root`."""
        part_elm = etree.SubElement(root, "{%s}part" % _PKG_NS)
        part_elm.set("{%s}name" % _PKG_NS, str(partname))
        part_elm.set("{%s}contentType" % _PKG_NS, content_type)

        if _is_xml_content_type(content_type):
            self._append_xml_data(part_elm, blob)
        else:
            self._append_binary_data(part_elm, blob)

    def _append_xml_data(self, part_elm: etree._Element, blob: bytes) -> None:
        """Parse `blob` and nest its root inside a ``<pkg:xmlData>`` child."""
        xml_data = etree.SubElement(part_elm, "{%s}xmlData" % _PKG_NS)
        # -- parse with a fresh parser so we don't inherit lxml's global
        # -- entity-resolution policy.  `remove_blank_text=False` preserves
        # -- any significant whitespace in the embedded XML.
        parser = etree.XMLParser(resolve_entities=False)
        inner = etree.fromstring(blob, parser=parser)
        xml_data.append(inner)

    def _append_binary_data(self, part_elm: etree._Element, blob: bytes) -> None:
        """Base64-encode `blob` into a ``<pkg:binaryData>`` child."""
        bin_data = etree.SubElement(part_elm, "{%s}binaryData" % _PKG_NS)
        bin_data.text = base64.b64encode(blob).decode("ascii")


def _is_xml_content_type(content_type: str) -> bool:
    """Return |True| when `content_type` identifies an XML payload.

    Used to decide whether a part's blob should be embedded inside
    ``<pkg:xmlData>`` (for XML parts) or base64-encoded inside
    ``<pkg:binaryData>`` (for everything else).

    The predicate follows the ECMA-376 Part 4 rule: any media type ending
    in ``+xml`` or equal to ``application/xml``/``text/xml`` is treated as
    XML.  OPC relationships (``...relationships+xml``) therefore naturally
    round-trip as XML parts.
    """
    ct = content_type.lower()
    if ct.endswith("+xml"):
        return True
    return ct in ("application/xml", "text/xml")
