# encoding: utf-8

"""
API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.
"""

from __future__ import absolute_import

import os
import weakref

import pptx.opc.packaging

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PACKAGE_URI
from pptx.opc.pkgreader import PackageReader
from pptx.opc.rels import RelationshipCollection
from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import serialize_part_xml
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import Image, ImageCollection
from pptx.parts.part import BasePart, PartCollection
from pptx.parts.slides import SlideCollection


def _PartFactory(partname, content_type, blob):
    part_cls = _PartClass(content_type)
    return part_cls.load(partname, content_type, blob)


def _PartClass(content_type):
    """
    Part class selector. Returns the Part subclass appropriate for the given
    reltype, content_type pair.
    """
    PRS_MAIN_CONTENT_TYPES = (
        CT.PML_PRESENTATION_MAIN, CT.PML_TEMPLATE_MAIN,
        CT.PML_SLIDESHOW_MAIN
    )

    from pptx.parts.slides import Slide, SlideLayout, SlideMaster

    if content_type == CT.OPC_CORE_PROPERTIES:
        return CoreProperties
    if content_type in PRS_MAIN_CONTENT_TYPES:
        return Presentation
    if content_type == CT.PML_SLIDE:
        return Slide
    if content_type == CT.PML_SLIDE_LAYOUT:
        return SlideLayout
    if content_type == CT.PML_SLIDE_MASTER:
        return SlideMaster
    if content_type.startswith('image/'):
        return Image
    return BasePart


class Package(object):
    """
    Return an instance of |Package| loaded from *file*, where *file* can be a
    path (a string) or a file-like object. If *file* is a path, it can be
    either a path to a PowerPoint `.pptx` file or a path to a directory
    containing an expanded presentation file, as would result from unzipping
    a `.pptx` file. If *file* is |None|, the default presentation template is
    loaded.
    """
    # track instances as weakrefs so .containing() can be computed
    _instances = []

    # path of the default presentation, used when no path specified
    _default_pptx_path = os.path.join(
        os.path.split(__file__)[0], 'templates', 'default.pptx'
    )

    def __init__(self):
        super(Package, self).__init__()
        self._instances.append(weakref.ref(self))  # track instances in cls var
        self._rels = RelationshipCollection(PACKAGE_URI.baseURI)

    @classmethod
    def containing(cls, part):
        """Return package instance that contains *part*"""
        for pkg in cls.instances():
            if part in pkg._parts:
                return pkg
        raise KeyError("No package contains part %r" % part)

    @property
    def core_properties(self):
        """
        Instance of |CoreProperties| holding the read/write Dublin Core
        document properties for this presentation.
        """
        if not hasattr(self, '_core_properties'):
            try:
                rel = self._rels.get_rel_of_type(RT.CORE_PROPERTIES)
                self._core_properties = rel.target_part
            except KeyError:
                core_props = CoreProperties._default()
                rId = self._rels.next_rId
                rel = self._rels.add_relationship(
                    RT.CORE_PROPERTIES, core_props, rId
                )
                self._core_properties = core_props
        return self._core_properties

    @classmethod
    def instances(cls):
        """Return tuple of Package instances that have been created"""
        # clean garbage collected pkgs out of _instances
        cls._instances[:] = [wkref for wkref in cls._instances
                             if wkref() is not None]
        # return instance references in a tuple
        pkgs = [wkref() for wkref in cls._instances]
        return tuple(pkgs)

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
        pkg._load(pkg_reader)
        return pkg

    @property
    def presentation(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        if not hasattr(self, '_presentation'):
            rel = self._rels.get_rel_of_type(RT.OFFICE_DOCUMENT)
            self._presentation = rel.target_part
        return self._presentation

    def save(self, file):
        """
        Save this package to *file*, where *file* can be either a path to a
        file (a string) or a file-like object.
        """
        pkgng_pkg = pptx.opc.packaging.Package().marshal(self)
        pkgng_pkg.save(file)

    def _add_relationship(self, reltype, target, rId, is_external=False):
        """
        Return newly added |_Relationship| instance of *reltype* between this
        package and part *target* with key *rId*. Target mode is set to
        ``RTM.EXTERNAL`` if *is_external* is |True|.
        """
        return self._rels.add_relationship(reltype, target, rId, is_external)

    @property
    def _images(self):
        """
        Collection containing a reference to each of the image parts in this
        package.
        """
        if not hasattr(self, '_image_parts'):
            self._image_parts = ImageCollection()
        return self._image_parts

    def _load(self, pkg_reader):
        """
        Load all the model-side parts and relationships from the on-disk
        package by loading package-level relationship parts and propagating
        the load down the relationship graph.
        """
        part_factory = _PartFactory
        pkg = self

        parts = self._unmarshal_parts(pkg_reader, part_factory)
        self._unmarshal_relationships(pkg_reader, pkg, parts)
        for part in parts.values():
            part.after_unmarshal()

        # gather references to image parts into _images
        def is_image_part(part):
            return (
                isinstance(part, Image) and
                part.partname.startswith('/ppt/media/')
            )
        for part in self._parts:
            if is_image_part(part):
                self._images._loadpart(part)

    @property
    def _parts(self):
        """
        Return a list containing a reference to each of the parts in this
        package.
        """
        return [part for part in Package._walkparts(self._rels)]

    @staticmethod
    def _unmarshal_parts(pkg_reader, part_factory):
        """
        Return a dictionary of |Part| instances unmarshalled from
        *pkg_reader*, keyed by partname. Side-effect is that each part in
        *pkg_reader* is constructed using *part_factory*.
        """
        parts = {}
        for partname, content_type, blob in pkg_reader.iter_sparts():
            parts[partname] = part_factory(partname, content_type, blob)
        return parts

    @staticmethod
    def _unmarshal_relationships(pkg_reader, pkg, parts):
        """
        Add a relationship to the source object corresponding to each of the
        relationships in *pkg_reader* with its target_part set to the actual
        target part in *parts*.
        """
        for source_uri, srel in pkg_reader.iter_srels():
            source = pkg if source_uri == '/' else parts[source_uri]
            target = (srel.target_ref if srel.is_external
                      else parts[srel.target_partname])
            source._add_relationship(srel.reltype, target, srel.rId,
                                     srel.is_external)

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
            for part in Package._walkparts(part._relationships, parts):
                yield part


class Presentation(BasePart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    def __init__(self, partname, content_type, presentation_elm):
        super(Presentation, self).__init__(partname, content_type)
        self._element = presentation_elm

    def after_unmarshal(self):
        # selectively unmarshal relationships for now
        for rel in self._relationships:
            if rel.reltype == RT.SLIDE_MASTER:
                self.slidemasters._loadpart(rel.target_part)

    @property
    def blob(self):
        return serialize_part_xml(self._element)

    @classmethod
    def load(cls, partname, content_type, blob):
        presentation_elm = parse_xml_bytes(blob)
        presentation = cls(partname, content_type, presentation_elm)
        return presentation

    @property
    def slidemasters(self):
        """
        Sequence of |SlideMaster| instances belonging to this presentation.
        """
        if not hasattr(self, '_slidemasters'):
            self._slidemasters = PartCollection()
        return self._slidemasters

    @property
    def slides(self):
        """
        |SlideCollection| object containing the slides in this presentation.
        """
        if not hasattr(self, '_slides'):
            sldIdLst = self._element.get_or_add_sldIdLst()
            rels = self._relationships
            self._slides = SlideCollection(sldIdLst, rels, self)
        return self._slides
