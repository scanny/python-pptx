"""lxml custom element classes for legend-related XML elements."""

from __future__ import annotations

from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_Legend(BaseOxmlElement):
    """
    ``<c:legend>`` custom element class
    """

    _tag_seq = (
        "c:legendPos",
        "c:legendEntry",
        "c:layout",
        "c:overlay",
        "c:spPr",
        "c:txPr",
        "c:extLst",
    )
    legendPos = ZeroOrOne("c:legendPos", successors=_tag_seq[1:])
    legendEntry = ZeroOrMore("c:legendEntry", successors=_tag_seq[2:])
    layout = ZeroOrOne("c:layout", successors=_tag_seq[3:])
    overlay = ZeroOrOne("c:overlay", successors=_tag_seq[4:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[6:])
    del _tag_seq

    @property
    def defRPr(self):
        """
        `./c:txPr/a:p/a:pPr/a:defRPr` great-great-grandchild element, added
        with its ancestors if not present.
        """
        txPr = self.get_or_add_txPr()
        defRPr = txPr.defRPr
        return defRPr

    def get_legendEntry_for_idx(self, idx):
        """Return the ``c:legendEntry`` whose ``c:idx/@val`` equals *idx*.

        Returns |None| when no such child is present.
        """
        matches = self.xpath('c:legendEntry[c:idx[@val="%d"]]' % idx)
        if matches:
            return matches[0]
        return None

    def get_or_add_legendEntry_for_idx(self, idx):
        """Return the ``c:legendEntry`` for *idx*, inserting one in order if absent.

        The new entry is inserted ahead of any existing ``c:legendEntry`` with
        a higher ``c:idx`` value so that the sequence remains numerically
        ordered. It carries only the mandatory ``c:idx`` child; callers
        populate ``c:delete`` (or other children) as needed.
        """
        match = self.get_legendEntry_for_idx(idx)
        if match is not None:
            return match
        return self._insert_legendEntry_in_sequence(idx)

    @property
    def hidden_entry_idxs(self):
        """Tuple of integer ``c:idx/@val`` values for entries marked deleted.

        Only ``c:legendEntry`` elements whose ``c:delete/@val`` is truthy
        contribute; any others (e.g. entries carrying only a ``c:txPr``
        formatting override) are ignored.
        """
        results = self.xpath('c:legendEntry[c:delete[not(@val="0")]]/c:idx/@val')
        return tuple(int(val) for val in results)

    @property
    def horz_offset(self):
        """
        The float value in ./c:layout/c:manualLayout/c:x when
        ./c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that
        XPath expression has no match.
        """
        layout = self.layout
        if layout is None:
            return 0.0
        return layout.horz_offset

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:layout/c:manualLayout/c:x@val to *offset* and
        ./c:layout/c:manualLayout/c:xMode@val to "factor". Remove
        ./c:layout/c:manualLayout if *offset* == 0.
        """
        layout = self.get_or_add_layout()
        layout.horz_offset = offset

    def _insert_legendEntry_in_sequence(self, idx):
        """Create a ``c:legendEntry`` for *idx* and insert it in ascending order."""
        new_legendEntry = CT_LegendEntry.new_legendEntry(idx)
        last = None
        for le in self.xpath("c:legendEntry"):
            if le.idx.val > idx:
                le.addprevious(new_legendEntry)
                return new_legendEntry
            last = le
        if last is not None:
            last.addnext(new_legendEntry)
            return new_legendEntry
        # -- no existing legendEntry; place after c:legendPos if present,
        # -- otherwise at the front --
        legendPos = self.legendPos
        if legendPos is not None:
            legendPos.addnext(new_legendEntry)
        else:
            self.insert(0, new_legendEntry)
        return new_legendEntry

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_LegendEntry(BaseOxmlElement):
    """``<c:legendEntry>`` element — one per-entry override within ``c:legend``.

    Carries a required ``c:idx`` specifying which 0-based legend entry it
    governs, then either a ``c:delete`` flag (hide the entry) or a
    ``c:txPr`` formatting override.
    """

    _tag_seq = ("c:idx", "c:delete", "c:txPr", "c:extLst")
    idx = OneAndOnlyOne("c:idx")
    delete_ = ZeroOrOne("c:delete", successors=_tag_seq[2:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[3:])
    del _tag_seq

    @classmethod
    def new_legendEntry(cls, idx):
        """Return a newly created ``c:legendEntry`` with ``c:idx/@val = idx``."""
        xml = '<c:legendEntry %s>\n  <c:idx val="%d"/>\n</c:legendEntry>' % (
            nsdecls("c"),
            idx,
        )
        return parse_xml(xml)


class CT_LegendPos(BaseOxmlElement):
    """
    ``<c:legendPos>`` element specifying position of legend with respect to
    chart as a member of ST_LegendPos.
    """

    val = OptionalAttribute("val", XL_LEGEND_POSITION, default=XL_LEGEND_POSITION.RIGHT)
