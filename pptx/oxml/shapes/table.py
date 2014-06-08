# encoding: utf-8

"""
lxml custom element classes for table-related XML elements.
"""

from __future__ import absolute_import

from .. import parse_xml
from ...enum.text import MSO_ANCHOR
from ..ns import nsdecls, qn
from ..shared import Element, SubElement
from ..simpletypes import ST_Coordinate, XsdBoolean
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, OptionalAttribute, RequiredAttribute,
    ZeroOrMore, ZeroOrOne
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
        If *value* is truthy, the attribute is set to "1"; a tblPr child
        element is added if necessary. If *value* is falsey, the *propname*
        attribute is removed if present, allowing its default value of False
        to be its effective value.
        """
        value = True if value else None
        if value is None and self.tblPr is None:
            return
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
        rowheight = height/rows
        colwidth = width/cols

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
    """``<a:tc>`` custom element class"""
    _tc_tmpl = (
        '<a:tc %s>\n'
        '  <a:txBody>\n'
        '    <a:bodyPr/>\n'
        '    <a:lstStyle/>\n'
        '    <a:p/>\n'
        '  </a:txBody>\n'
        '  <a:tcPr/>\n'
        '</a:tc>' % nsdecls('a')
    )

    def __setattr__(self, attr, value):
        """
        This hack is needed to make setter side of properties work,
        overrides ``__setattr__`` defined in ObjectifiedElement super class
        just enough to route messages intended for custom property setters.
        """
        if attr == 'anchor':
            self._set_anchor(value)
        elif attr in ('marT', 'marR', 'marB', 'marL'):
            self._set_marX(attr, value)
        else:
            super(CT_TableCell, self).__setattr__(attr, value)

    @property
    def anchor(self):
        """
        String held in ``anchor`` attribute of ``<a:tcPr>`` child element of
        this ``<a:tc>`` element.
        """
        if self.tcPr is None:
            return None
        anchor = self.tcPr.get('anchor')
        return MSO_ANCHOR.from_xml(anchor)

    def get_or_add_tcPr(self):
        tcPr = self.tcPr
        if tcPr is None:
            tcPr = Element('a:tcPr')
            idx = 1 if self.txBody else 0
            self.insert(idx, tcPr)
        return tcPr

    def get_or_add_txBody(self):
        """
        Return the <a:rPr> child element of this <a:r> element, newly added
        if not already present.
        """
        if self.txBody is None:
            txBody = Element('a:txBody')
            SubElement(txBody, 'a:bodyPr')
            SubElement(txBody, 'a:p')
            self.insert(0, txBody)
        return self.txBody

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

    @property
    def marR(self):
        """right margin value represented in ``marR`` attribute"""
        return self._get_marX('marR', 91440)

    @property
    def marB(self):
        """bottom margin value represented in ``marB`` attribute"""
        return self._get_marX('marB', 45720)

    @property
    def marL(self):
        """left margin value represented in ``marL`` attribute"""
        return self._get_marX('marL', 91440)

    @staticmethod
    def new_tc():
        """
        Return a new ``<a:tc>`` element tree.
        """
        xml = CT_TableCell._tc_tmpl
        tc = parse_xml(xml)
        return tc

    @property
    def tcPr(self):
        return self.find(qn('a:tcPr'))

    @property
    def txBody(self):
        """
        The <a:txBody> child element, or None if not present.
        """
        return self.find(qn('a:txBody'))

    def _clear_anchor(self):
        """
        Remove anchor attribute from ``<a:tcPr>`` if it exists
        """
        if self.tcPr is None:
            return
        if 'anchor' in self.tcPr.attrib:
            del self.tcPr.attrib['anchor']

    def _get_marX(self, attr_name, default):
        """
        generalized method to get margin values
        """
        if self.tcPr is None:
            return default
        return int(self.tcPr.get(attr_name, default))

    def _set_anchor(self, anchor_enum_idx):
        """
        Set value of anchor attribute on ``<a:tcPr>`` child element
        """
        anchor = MSO_ANCHOR.to_xml(anchor_enum_idx)
        if anchor is None:
            return self._clear_anchor()
        tcPr = self.get_or_add_tcPr()
        tcPr.set('anchor', anchor)

    def _set_marX(self, marX, value):
        """
        Set value of marX attribute on ``<a:tcPr>`` child element. If *marX*
        is |None|, the marX attribute is removed.
        """
        tcPr = self.get_or_add_tcPr()
        if value is None:
            if marX in tcPr.attrib:
                del tcPr.attrib[marX]
        else:
            tcPr.set(marX, str(value))


class CT_TableCellProperties(BaseOxmlElement):
    """
    ``<a:tcPr>`` custom element class
    """
    @property
    def fill_element(self):
        """
        Return the child representing the EG_FillProperties element group
        member in this element, or |None| if no such child is present.
        """
        return self._first_child_found_in(
            'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
            'a:pattFill', 'a:grpFill'
        )

    def get_or_change_to_noFill(self):
        """
        Return the <a:noFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        # return existing one if there is one
        if self.noFill is not None:
            return self.noFill
        # get rid of other fill element type if there is one
        self._remove_if_present(
            'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill',
            'a:grpFill'
        )
        # add noFill element in right sequence
        return self._add_noFill()

    def get_or_change_to_solidFill(self):
        """
        Return the <a:solidFill> child element, replacing any other fill
        element if found, e.g. a <a:gradFill> element.
        """
        # return existing one if there is one
        if self.solidFill is not None:
            return self.solidFill
        # get rid of other fill element type if there is one
        self._remove_if_present(
            'a:noFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'
        )
        # add solidFill element in right sequence
        return self._add_solidFill()

    @property
    def noFill(self):
        """
        The <a:noFill> child element, or None if not present.
        """
        return self.find(qn('a:noFill'))

    @property
    def solidFill(self):
        """
        The <a:solidFill> child element, or None if not present.
        """
        return self.find(qn('a:solidFill'))

    def _add_noFill(self):
        """
        Return a newly added <a:noFill> child element; assume no other fill
        EG_FillProperties element is present.
        """
        noFill = Element('a:noFill')
        successor = self._first_successor_in('a:headers', 'a:extLst')
        if successor is not None:
            successor.addprevious(noFill)
        else:
            self.append(noFill)
        return noFill

    def _add_solidFill(self):
        """
        Return a newly added <a:solidFill> child element.
        """
        solidFill = Element('a:solidFill')
        successor = self._first_successor_in('a:headers', 'a:extLst')
        if successor is not None:
            successor.addprevious(solidFill)
        else:
            self.append(solidFill)
        return solidFill

    def _first_child_found_in(self, *tagnames):
        """
        Return the first child found with tag in *tagnames*, or None if
        not found.
        """
        for tagname in tagnames:
            child = self.find(qn(tagname))
            if child is not None:
                return child
        return None

    def _first_successor_in(self, *successor_tagnames):
        """
        Return the first child with tag in *successor_tagnames*, or None if
        not found.
        """
        for tagname in successor_tagnames:
            successor = self.find(qn(tagname))
            if successor is not None:
                return successor
        return None

    def _remove_if_present(self, *tagnames):
        for tagname in tagnames:
            element = self.find(qn(tagname))
            if element is not None:
                self.remove(element)


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
    bandRow = OptionalAttribute('bandRow', XsdBoolean)
    bandCol = OptionalAttribute('bandCol', XsdBoolean)
    firstRow = OptionalAttribute('firstRow', XsdBoolean)
    firstCol = OptionalAttribute('firstCol', XsdBoolean)
    lastRow = OptionalAttribute('lastRow', XsdBoolean)
    lastCol = OptionalAttribute('lastCol', XsdBoolean)


class CT_TableRow(BaseOxmlElement):
    """
    ``<a:tr>`` custom element class
    """
    tc = ZeroOrMore('a:tc', successors=('a:extLst',))
    h = RequiredAttribute('h', ST_Coordinate)

    def add_tc(self):
        """
        Return a reference to a newly created <a:gridCol> child element
        having its ``w`` attribute set to *width*.
        """
        return self._add_tc()

    def _new_tc(self):
        return CT_TableCell.new_tc()
