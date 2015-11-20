# encoding: utf-8

"""
lxml custom element classes for table-related XML elements.
"""

from __future__ import absolute_import, division

from .. import parse_xml
from ...enum.text import MSO_VERTICAL_ANCHOR
from ..ns import nsdecls
from ..simpletypes import ST_Coordinate, ST_Coordinate32, XsdBoolean
from ..text import CT_TextBody
from ..xmlchemy import (
    BaseOxmlElement, Choice, OneAndOnlyOne, OptionalAttribute,
    RequiredAttribute, ZeroOrMore, ZeroOrOne, ZeroOrOneChoice
)


class CT_Table(BaseOxmlElement):
    """
    ``<a:tbl>`` custom element class
    """
    tblPr = ZeroOrOne('a:tblPr', successors=('a:tblGrid', 'a:tr'))
    tblGrid = OneAndOnlyOne('a:tblGrid')
    tr = ZeroOrMore('a:tr', successors=())

    def add_tr(self, height):
        """
        Return a reference to a newly created <a:tr> child element having its
        ``h`` attribute set to *height*.
        """
        return self._add_tr(h=height)

    @property
    def bandCol(self):
        return self._get_boolean_property('bandCol')

    @bandCol.setter
    def bandCol(self, value):
        self._set_boolean_property('bandCol', value)

    @property
    def bandRow(self):
        return self._get_boolean_property('bandRow')

    @bandRow.setter
    def bandRow(self, value):
        self._set_boolean_property('bandRow', value)

    @property
    def firstCol(self):
        return self._get_boolean_property('firstCol')

    @firstCol.setter
    def firstCol(self, value):
        self._set_boolean_property('firstCol', value)

    @property
    def firstRow(self):
        return self._get_boolean_property('firstRow')

    @firstRow.setter
    def firstRow(self, value):
        self._set_boolean_property('firstRow', value)

    @property
    def lastCol(self):
        return self._get_boolean_property('lastCol')

    @lastCol.setter
    def lastCol(self, value):
        self._set_boolean_property('lastCol', value)

    @property
    def lastRow(self):
        return self._get_boolean_property('lastRow')

    @lastRow.setter
    def lastRow(self, value):
        self._set_boolean_property('lastRow', value)

    def _get_boolean_property(self, propname):
        """
        Generalized getter for the boolean properties on the ``<a:tblPr>``
        child element. Defaults to False if *propname* attribute is missing
        or ``<a:tblPr>`` element itself is not present.
        """
        tblPr = self.tblPr
        if tblPr is None:
            return False
        propval = getattr(tblPr, propname)
        return {
            True:  True,
            False: False,
            None:  False
        }[propval]

    def _set_boolean_property(self, propname, value):
        """
        Generalized setter for boolean properties on the ``<a:tblPr>`` child
        element, setting *propname* attribute appropriately based on *value*.
        If *value* is True, the attribute is set to "1"; a tblPr child
        element is added if necessary. If *value* is False, the *propname*
        attribute is removed if present, allowing its default value of False
        to be its effective value.
        """
        if value not in (True, False):
            raise ValueError(
                "assigned value must be either True or False, got %s" %
                value
            )
        tblPr = self.get_or_add_tblPr()
        setattr(tblPr, propname, value)

    @classmethod
    def new_tbl(cls, rows, cols, width, height, tableStyleId=None):
        """
        Return a new ``<p:tbl>`` element tree
        """
        # working hypothesis is this is the default table style GUID
        if tableStyleId is None:
            tableStyleId = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

        xml = cls._tbl_tmpl() % (tableStyleId)
        tbl = parse_xml(xml)

        # add specified number of rows and columns
        rowheight = height//rows
        colwidth = width//cols

        for col in range(cols):
            # adjust width of last col to absorb any div error
            if col == cols-1:
                colwidth = width - ((cols-1) * colwidth)
            tbl.tblGrid.add_gridCol(width=colwidth)

        for row in range(rows):
            # adjust height of last row to absorb any div error
            if row == rows-1:
                rowheight = height - ((rows-1) * rowheight)
            tr = tbl.add_tr(height=rowheight)
            for col in range(cols):
                tr.add_tc()

        return tbl

    @classmethod
    def _tbl_tmpl(cls):
        return (
            '<a:tbl %s>\n'
            '  <a:tblPr firstRow="1" bandRow="1">\n'
            '    <a:tableStyleId>%s</a:tableStyleId>\n'
            '  </a:tblPr>\n'
            '  <a:tblGrid/>\n'
            '</a:tbl>' % (nsdecls('a'), '%s')
        )


class CT_TableCell(BaseOxmlElement):
    """
    ``<a:tc>`` custom element class
    """
    txBody = ZeroOrOne('a:txBody', successors=('a:tcPr', 'a:extLst',))
    tcPr = ZeroOrOne('a:tcPr', successors=('a:extLst',))

    @property
    def anchor(self):
        """
        String held in ``anchor`` attribute of ``<a:tcPr>`` child element of
        this ``<a:tc>`` element.
        """
        if self.tcPr is None:
            return None
        return self.tcPr.anchor

    @anchor.setter
    def anchor(self, anchor_enum_idx):
        """
        Set value of anchor attribute on ``<a:tcPr>`` child element
        """
        if anchor_enum_idx is None and self.tcPr is None:
            return
        tcPr = self.get_or_add_tcPr()
        tcPr.anchor = anchor_enum_idx

    @property
    def marT(self):
        """
        Read/write integer top margin value represented in ``marT`` attribute
        of the ``<a:tcPr>`` child element of this ``<a:tc>`` element. If the
        attribute is not present, the default value ``45720`` (0.05 inches)
        is returned for top and bottom; ``91440`` (0.10 inches) is the
        default for left and right. Assigning |None| to any ``marX``
        property clears that attribute from the element, effectively setting
        it to the default value.
        """
        return self._get_marX('marT', 45720)

    @marT.setter
    def marT(self, value):
        self._set_marX('marT', value)

    @property
    def marR(self):
        """
        Right margin value represented in ``marR`` attribute.
        """
        return self._get_marX('marR', 91440)

    @marR.setter
    def marR(self, value):
        self._set_marX('marR', value)

    @property
    def marB(self):
        """
        Bottom margin value represented in ``marB`` attribute.
        """
        return self._get_marX('marB', 45720)

    @marB.setter
    def marB(self, value):
        self._set_marX('marB', value)

    @property
    def marL(self):
        """
        Left margin value represented in ``marL`` attribute.
        """
        return self._get_marX('marL', 91440)

    @marL.setter
    def marL(self, value):
        self._set_marX('marL', value)

    @classmethod
    def new(cls):
        """
        Return a new ``<a:tc>`` element tree.
        """
        xml = cls._tc_tmpl()
        tc = parse_xml(xml)
        return tc

    def _get_marX(self, attr_name, default):
        """
        Generalized method to get margin values.
        """
        if self.tcPr is None:
            return default
        return int(self.tcPr.get(attr_name, default))

    def _new_txBody(self):
        return CT_TextBody.new_a_txBody()

    def _set_marX(self, marX, value):
        """
        Set value of marX attribute on ``<a:tcPr>`` child element. If *marX*
        is |None|, the marX attribute is removed. *marX* is a string, one of
        ``('marL', 'marR', 'marT', 'marB')``.
        """
        if value is None and self.tcPr is None:
            return
        tcPr = self.get_or_add_tcPr()
        setattr(tcPr, marX, value)

    @classmethod
    def _tc_tmpl(cls):
        return (
            '<a:tc %s>\n'
            '  <a:txBody>\n'
            '    <a:bodyPr/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p/>\n'
            '  </a:txBody>\n'
            '  <a:tcPr/>\n'
            '</a:tc>' % nsdecls('a')
        )


class CT_TableCellProperties(BaseOxmlElement):
    """
    ``<a:tcPr>`` custom element class
    """
    eg_fillProperties = ZeroOrOneChoice((
        Choice('a:noFill'), Choice('a:solidFill'), Choice('a:gradFill'),
        Choice('a:blipFill'), Choice('a:pattFill'), Choice('a:grpFill')),
        successors=('a:headers', 'a:extLst')
    )
    anchor = OptionalAttribute('anchor', MSO_VERTICAL_ANCHOR)
    marL = OptionalAttribute('marL', ST_Coordinate32)
    marR = OptionalAttribute('marR', ST_Coordinate32)
    marT = OptionalAttribute('marT', ST_Coordinate32)
    marB = OptionalAttribute('marB', ST_Coordinate32)


class CT_TableCol(BaseOxmlElement):
    """
    ``<a:gridCol>`` custom element class
    """
    w = RequiredAttribute('w', ST_Coordinate)


class CT_TableGrid(BaseOxmlElement):
    """
    ``<a:tblGrid>`` custom element class
    """
    gridCol = ZeroOrMore('a:gridCol')

    def add_gridCol(self, width):
        """
        Return a reference to a newly created <a:gridCol> child element
        having its ``w`` attribute set to *width*.
        """
        return self._add_gridCol(w=width)


class CT_TableProperties(BaseOxmlElement):
    """
    ``<a:tblPr>`` custom element class
    """
    bandRow = OptionalAttribute('bandRow', XsdBoolean, default=False)
    bandCol = OptionalAttribute('bandCol', XsdBoolean, default=False)
    firstRow = OptionalAttribute('firstRow', XsdBoolean, default=False)
    firstCol = OptionalAttribute('firstCol', XsdBoolean, default=False)
    lastRow = OptionalAttribute('lastRow', XsdBoolean, default=False)
    lastCol = OptionalAttribute('lastCol', XsdBoolean, default=False)


class CT_TableRow(BaseOxmlElement):
    """
    ``<a:tr>`` custom element class
    """
    tc = ZeroOrMore('a:tc', successors=('a:extLst',))
    h = RequiredAttribute('h', ST_Coordinate)

    def add_tc(self):
        """
        Return a reference to a newly added minimal valid ``<a:tc>`` child
        element.
        """
        return self._add_tc()

    def _new_tc(self):
        return CT_TableCell.new()
