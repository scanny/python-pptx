# encoding: utf-8

"""
API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.
"""

from __future__ import absolute_import

import os

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PACKAGE_URI
from pptx.opc.package import (
    Part, PartFactory, RelationshipCollection, Unmarshaller
)
from pptx.opc.pkgreader import PackageReader
from pptx.opc.pkgwriter import PackageWriter
from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import serialize_part_xml
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import ImageCollection
from pptx.parts.part import PartCollection
from pptx.parts.slides import SlideCollection
from pptx.util import lazyproperty


class Package(object):
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

    def __init__(self):
        super(Package, self).__init__()
        self._rels = RelationshipCollection(PACKAGE_URI.baseURI)

    def after_unmarshal(self):
        """
        Called by loading code after all parts and relationships have been
        loaded, to afford the opportunity for any required post-processing.
        """
        # gather image parts into _images
        self._images.load(self._parts)

    @lazyproperty
    def core_properties(self):
        """
        Instance of |CoreProperties| holding the read/write Dublin Core
        document properties for this presentation.
        """
        try:
            return self._rels.part_with_reltype(RT.CORE_PROPERTIES)
        except KeyError:
            core_props = CoreProperties.default()
            self._rels.get_or_add(RT.CORE_PROPERTIES, core_props)
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
        pkg_reader = PackageReader.from_file(pkg_file)
        pkg = cls()
        Unmarshaller.unmarshal(pkg_reader, pkg, PartFactory)
        return pkg

    @lazyproperty
    def presentation(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        return self._rels.part_with_reltype(RT.OFFICE_DOCUMENT)

    def save(self, pkg_file):
        """
        Save this package to *pkg_file*, which can be either a path to a file
        (a string) or a file-like object.
        """
        # self._notify_before_marshal()
        for part in self._parts:
            part.before_marshal()
        PackageWriter.write(pkg_file, self._rels, self._parts)

    def _add_relationship(self, reltype, target, rId, is_external=False):
        """
        Return newly added |_Relationship| instance of *reltype* between this
        package and part *target* with key *rId*. Target mode is set to
        ``RTM.EXTERNAL`` if *is_external* is |True|.
        """
        return self._rels.add_relationship(reltype, target, rId, is_external)

    @lazyproperty
    def _images(self):
        """
        Collection containing a reference to each of the image parts in this
        package.
        """
        return ImageCollection()

    @property
    def _parts(self):
        """
        Return a list containing a reference to each of the parts in this
        package.
        """
        return [part for part in Package._walkparts(self._rels)]

    @staticmethod
    def _walkparts(rels, parts=None):
        """
        Recursive function, walk relationships to iterate over all parts in
        this package. Leave out *parts* parameter in call to visit all parts.
        """
        # initial call can leave out parts parameter as a signal to initialize
        if parts is None:
            parts = []
        for rel in rels:
            part = rel.target_part
            # only visit each part once (graph is cyclic)
            if part in parts:
                continue
            parts.append(part)
            yield part
            for part in Package._walkparts(part._rels, parts):
                yield part


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
        sm_rels = [r for r in self._rels if r.reltype == RT.SLIDE_MASTER]
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
        slides = SlideCollection(sldIdLst, self._rels, self)
        return slides
