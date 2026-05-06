"""Re-export of :mod:`ooxml_opc.oxml` with pptx-local helpers.

The :class:`CT_Types` / :class:`CT_Relationships` / :class:`CT_Default` /
:class:`CT_Override` / :class:`CT_Relationship` xmlchemy element classes
live in :mod:`ooxml_opc.oxml` and are re-exported here so external
callers (including test code) continue to import them from this module.

pptx-local helpers retained:

* :func:`_normalize_xml_decl_quotes` — rewrites the lxml single-quoted
  XML declaration to double quotes to match Microsoft Office.
* :func:`oxml_to_encoded_bytes`, :func:`oxml_tostring` — thin wrappers
  around :func:`etree.tostring` that apply the declaration-quote
  normalization on the returned bytes.
* :data:`nsmap` — OPC-local prefix → URI mapping.

The shared CT_ classes are also registered in pptx's own element-class
lookup so ``pptx.oxml.parse_xml`` returns them on parse.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ooxml_opc.oxml import (  # noqa: F401 -- re-exports
    CT_Default,
    CT_Override,
    CT_Relationship,
    CT_Relationships,
    CT_Types,
    serialize_part_xml,
)
from ooxml_opc.constants import NAMESPACE as NS

if TYPE_CHECKING:
    from ooxml_opc.oxml import BaseOxmlElement

__all__ = [
    "CT_Default",
    "CT_Override",
    "CT_Relationship",
    "CT_Relationships",
    "CT_Types",
    "nsmap",
    "oxml_to_encoded_bytes",
    "oxml_tostring",
    "serialize_part_xml",
]


#: OPC-local nsmap.
nsmap = {
    "ct": NS.OPC_CONTENT_TYPES,
    "pr": NS.OPC_RELATIONSHIPS,
    "r": NS.OFC_RELATIONSHIPS,
}


def _normalize_xml_decl_quotes(xml_bytes: bytes) -> bytes:
    """Rewrite the leading XML-declaration in `xml_bytes` to use double quotes.

    lxml emits the XML declaration with single quotes; Microsoft Office
    and strict validators expect double quotes. No-op when there is no
    XML declaration.
    """
    if not xml_bytes.startswith(b"<?xml"):
        return xml_bytes
    end = xml_bytes.find(b"?>")
    if end == -1:
        return xml_bytes
    decl = xml_bytes[: end + 2]
    rest = xml_bytes[end + 2 :]
    decl = decl.replace(b"'", b'"')
    return decl + rest


def oxml_to_encoded_bytes(
    element: "BaseOxmlElement",
    encoding: str = "utf-8",
    pretty_print: bool = False,
    standalone: "bool | None" = None,
) -> bytes:
    xml = etree.tostring(
        element, encoding=encoding, pretty_print=pretty_print, standalone=standalone
    )
    return _normalize_xml_decl_quotes(xml)


def oxml_tostring(
    elm: "BaseOxmlElement",
    encoding: "str | None" = None,
    pretty_print: bool = False,
    standalone: "bool | None" = None,
):
    xml = etree.tostring(elm, encoding=encoding, pretty_print=pretty_print, standalone=standalone)
    if isinstance(xml, bytes):
        return _normalize_xml_decl_quotes(xml)
    return xml


# -- Register the shared CT_ classes in pptx's element-class lookup so
# -- ``pptx.oxml.parse_xml`` returns the custom classes rather than the
# -- generic ``_Element``. Deferred to dodge the circular-import path
# -- ``pptx.oxml.coreprops -> pptx.opc.oxml -> pptx.oxml``. --
def _register() -> None:
    from pptx.oxml import register_element_cls

    register_element_cls("ct:Default", CT_Default)
    register_element_cls("ct:Override", CT_Override)
    register_element_cls("ct:Types", CT_Types)
    register_element_cls("pr:Relationship", CT_Relationship)
    register_element_cls("pr:Relationships", CT_Relationships)


try:
    _register()
except ImportError:  # pragma: no cover
    pass
