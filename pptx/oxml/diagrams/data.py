# encoding: utf-8

"""Custom element classes for data Diagram-related XML elements."""

from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)
from pptx.oxml.simpletypes import XsdString


class CT_DataModel(BaseOxmlElement):
    """`dgm:dataModel` custom element class"""

    _tag_seq = (
        "dgm:ptList",
        "dgm:cxnLst",
        "dgm:bg",
        "dgm:whole",
        "dgm:extLst",
    )

    extLst = ZeroOrOne("dgm:extLst")


class CT_DataModelExtLst(BaseOxmlElement):
    """`dgm:extLst` custom element class"""
    ext = ZeroOrMore("a:ext", successors=())


class CT_DataModelExt(BaseOxmlElement):
    """`dsp:dataModelExt` custom element class"""
    relId = RequiredAttribute("relId", XsdString)

