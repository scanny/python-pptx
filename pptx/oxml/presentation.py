# encoding: utf-8

"""
Custom element classes for presentation-related XML elements.
"""

from __future__ import absolute_import

from .shared import BaseOxmlElement, child, Element, SubElement
from .ns import qn


class CT_Presentation(BaseOxmlElement):
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

    def get_or_add_sldMasterIdLst(self):
        """
        Return the sldMasterIdLst child element, newly added if not present.
        """
        sldMasterIdLst = self.sldMasterIdLst
        if sldMasterIdLst is None:
            sldMasterIdLst = self._add_sldMasterIdLst()
        return sldMasterIdLst

    @property
    def sldMasterIdLst(self):
        """
        The first ``<p:sldMasterIdLst>`` child element, or |None| if not
        present.
        """
        return self.find(qn('p:sldMasterIdLst'))

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

    def _add_sldMasterIdLst(self):
        """
        Return a newly added sldMasterIdLst child element. Assumes one is not
        present.
        """
        sldMasterIdLst = CT_SlideMasterIdList.new()
        self.insert(0, sldMasterIdLst)
        return sldMasterIdLst


class CT_SlideId(BaseOxmlElement):
    """
    ``<p:sldId>`` element, direct child of <p:sldIdLst> that contains an rId
    reference to a slide in the presentation.
    """
    @property
    def rId(self):
        return self.get(qn('r:id'))


class CT_SlideIdList(BaseOxmlElement):
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
        sldId = SubElement(self, 'p:sldId', id=self._next_id)
        sldId.set(qn('r:id'), rId)
        return sldId

    @property
    def _next_id(self):
        return str(256 + len(self))


class CT_SlideMasterIdList(BaseOxmlElement):
    """
    ``<p:sldMasterIdLst>`` element, child of ``<p:presentation>`` containing
    references to the slide masters that belong to the presentation.
    """
    def __len__(self):
        """
        Return the number of ``<p:sldMasterId>`` child elements
        """
        sldMasterId_lst = self.findall(qn('p:sldMasterId'))
        return len(sldMasterId_lst)

    @classmethod
    def new(cls):
        """
        Return a new ``<p:sldMasterIdLst>`` element.
        """
        return Element('p:sldMasterIdLst')

    @property
    def sldMasterId_lst(self):
        """
        Sequence of ``<p:sldMasterId>`` child elements
        """
        return self.findall(qn('p:sldMasterId'))


class CT_SlideMasterIdListEntry(BaseOxmlElement):
    """
    ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
    a reference to a slide master.
    """
    @property
    def rId(self):
        return self.get(qn('r:id'))


class CT_SlideSize(BaseOxmlElement):
    """
    ``<p:sldSz>`` element, direct child of <p:presentation> that contains the
    width and height of slides in the presentation.
    """
    @property
    def cx(self):
        return int(self.get('cx'))
