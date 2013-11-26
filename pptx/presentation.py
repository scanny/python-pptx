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
from pptx.oxml.core import serialize_part_xml
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
            partname, content_type, package=package
        )
        self._element = presentation_elm

    @property
    def blob(self):
        return serialize_part_xml(self._element)

    @classmethod
    def load(cls, partname, content_type, blob, package):
        presentation_elm = parse_xml_bytes(blob)
        presentation = cls(partname, content_type, presentation_elm, package)
        return presentation

    @lazyproperty
    def slidemasters(self):
        """
        Sequence of |SlideMaster| instances belonging to this presentation.
        """
        slidemasters = PartCollection()
        sm_rels = [r for r in self.rels if r.reltype == RT.SLIDE_MASTER]
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
        slides = SlideCollection(sldIdLst, self.rels, self)
        return slides
