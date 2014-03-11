# encoding: utf-8

"""
API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.
"""

from __future__ import absolute_import

import os

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import OpcPackage, Part
from pptx.oxml import parse_xml_bytes
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import ImageCollection
from pptx.parts.part import PartCollection
from pptx.parts.slides import SlideCollection
from pptx.util import lazyproperty


class Package(OpcPackage):
    """
    Return an instance of |Package| loaded from *file*, where *file* can be a
    path (a string) or a file-like object. If *file* is a path, it can be
    either a path to a PowerPoint `.pptx` file or a path to a directory
    containing an expanded presentation file, as would result from unzipping
    a `.pptx` file. If *file* is |None|, the default presentation template is
    loaded.
    """

    # path of the default presentation, used when no path specified
    _default_pptx_path = os.path.join(
        os.path.split(__file__)[0], 'templates', 'default.pptx'
    )

    def after_unmarshal(self):
        """
        Called by loading code after all parts and relationships have been
        loaded, to afford the opportunity for any required post-processing.
        """
        # gather image parts into _images
        self._images.load(self.parts)

    @lazyproperty
    def core_properties(self):
        """
        Instance of |CoreProperties| holding the read/write Dublin Core
        document properties for this presentation. Creates a default core
        properties part if one is not present (not common).
        """
        try:
            return self.part_related_by(RT.CORE_PROPERTIES)
        except KeyError:
            core_props = CoreProperties.default()
            self.relate_to(core_props, RT.CORE_PROPERTIES)
            return core_props

    @classmethod
    def open(cls, pkg_file=None):
        """
        Return |Package| instance loaded with contents of .pptx package at
        *pkg_file*, or the default presentation package if *pkg_file* is
        missing or |None|.
        """
        if pkg_file is None:
            pkg_file = cls._default_pptx_path
        return super(Package, cls).open(pkg_file)

    @property
    def presentation(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        return self.main_document

    @lazyproperty
    def _images(self):
        """
        Collection containing a reference to each of the image parts in this
        package.
        """
        return ImageCollection()


class Presentation(Part):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def __init__(self, partname, content_type, presentation_elm, package):
        super(Presentation, self).__init__(
            partname, content_type, element=presentation_elm, package=package
        )

    @classmethod
    def load(cls, partname, content_type, blob, package):
        presentation_elm = parse_xml_bytes(blob)
        presentation = cls(partname, content_type, presentation_elm, package)
        return presentation

    @property
    def sldMasterIdLst(self):
        """
        The ``<p:sldMasterIdLst>`` child element specifying the slide masters
        of this presentation in the XML.
        """
        return self._element.get_or_add_sldMasterIdLst()

    @lazyproperty
    def slide_masters(self):
        """
        Sequence of |SlideMaster| objects belonging to this presentation
        """
        return _SlideMasters(self)

    @lazyproperty
    def slidemasters(self):
        """
        Sequence of |SlideMaster| instances belonging to this presentation.
        """
        slidemasters = PartCollection()
        sm_rels = [
            r for r in self.rels.values() if r.reltype == RT.SLIDE_MASTER
        ]
        for sm_rel in sm_rels:
            slide_master = sm_rel.target_part
            slidemasters.add_part(slide_master)
        return slidemasters

    @lazyproperty
    def slides(self):
        """
        |SlideCollection| object containing the slides in this presentation.
        """
        sldIdLst = self._element.get_or_add_sldIdLst()
        slides = SlideCollection(sldIdLst, self)
        slides.rename_slides()  # start from known state
        return slides


class _SlideMasters(object):
    """
    Collection of |SlideMaster| instances belonging to a presentation. Has
    list access semantics, supporting indexed access, len(), and iteration.
    """
    def __init__(self, presentation):
        super(_SlideMasters, self).__init__()
        self._presentation = presentation

    def __getitem__(self, idx):
        """
        Provide indexed access, (e.g. ``slide_masters[2]``).
        """
        sldMasterId_lst = self._sldMasterIdLst.sldMasterId_lst
        if idx >= len(sldMasterId_lst):
            raise IndexError('slide master index out of range')
        rId = sldMasterId_lst[idx].rId
        return self._presentation.related_parts[rId]

    def __iter__(self):
        """
        Generate a reference to each of the |SlideMaster| instances in the
        collection, in sequence.
        """
        for rId in self._iter_rIds():
            yield self._presentation.related_parts[rId]

    def __len__(self):
        """
        Support len() built-in function (e.g. 'len(slide_masters) == 4').
        """
        return len(self._sldMasterIdLst)

    def _iter_rIds(self):
        """
        Generate the rId for each slide master in the collection, in
        sequence.
        """
        sldMasterId_lst = self._sldMasterIdLst.sldMasterId_lst
        for sldMasterId in sldMasterId_lst:
            yield sldMasterId.rId

    @property
    def _sldMasterIdLst(self):
        """
        The ``<p:sldMasterIdLst>`` element specifying the slide masters in
        this collection. This element is a child of the ``<p:presentation>``
        element, the root element of a presentation part.
        """
        return self._presentation.sldMasterIdLst
