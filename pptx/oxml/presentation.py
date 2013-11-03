# encoding: utf-8

"""
Custom element classes for presentation-related XML elements.
"""

from __future__ import absolute_import

from lxml import objectify

from pptx.oxml.core import child, Element
from pptx.oxml.ns import qn


class CT_Presentation(objectify.ObjectifiedElement):
    """
    ``<p:presentation>`` element, root of the Presentation part stored as
    ``/ppt/presentation.xml``.
    """

    def get_or_add_sldIdLst(self):
        """
        Return the <p:sldIdLst> child element, creating one first if
        necessary.
        """
        sldIdLst = child(self, 'p:sldIdLst')
        if sldIdLst is None:
            sldIdLst = self._add_sldIdLst()
        return sldIdLst

    def _add_sldIdLst(self):
        """
        Return a newly created <p:sldIdLst> child element.
        """
        sldIdLst = Element('p:sldIdLst')
        # insert new sldIdLst element in right sequence
        sldSz = child(self, 'p:sldSz')
        if sldSz is not None:
            sldSz.addprevious(sldIdLst)
        else:
            notesSz = child(self, 'p:notesSz')
            notesSz.addprevious(sldIdLst)
        return sldIdLst


class CT_SlideId(objectify.ObjectifiedElement):
    """
    ``<p:sldId>`` element, direct child of <p:sldIdLst> that contains an rId
    reference to a slide in the presentation.
    """
    @property
    def rId(self):
        return self.get(qn('r:id'))


class CT_SlideIdList(objectify.ObjectifiedElement):
    """
    ``<p:sldIdLst>`` element, direct child of <p:presentation> that contains
    a list of the slide parts in the presentation.
    """
    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. 'collection[0]').
        """
        return self.getchildren()[idx]

    def __iter__(self):
        return self.iterchildren()

    def __len__(self):
        return self.countchildren()

    def add_sldId(self, rId):
        """
        Return a reference to a newly created <p:sldId> child element having
        its r:id attribute set to *rId*.
        """
        id = 256 + len(self.getchildren())
        sldId = Element('p:sldId')
        sldId.set('id', str(id))
        sldId.set(qn('r:id'), rId)
        self.append(sldId)
        return sldId
