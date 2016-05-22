# encoding: utf-8

"""
API classes for dealing with presentations and other objects one typically
encounters as an end-user of the PowerPoint user interface.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .opc.constants import RELATIONSHIP_TYPE as RT
from .opc.package import OpcPackage
from .opc.packuri import PackURI
from .parts.coreprops import CorePropertiesPart
from .parts.image import Image, ImagePart
from .util import lazyproperty


class Package(OpcPackage):
    """
    Return an instance of |Package| loaded from *file*, where *file* can be a
    path (a string) or a file-like object. If *file* is a path, it can be
    either a path to a PowerPoint `.pptx` file or a path to a directory
    containing an expanded presentation file, as would result from unzipping
    a `.pptx` file. If *file* is |None|, the default presentation template is
    loaded.
    """
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
            core_props = CorePropertiesPart.default()
            self.relate_to(core_props, RT.CORE_PROPERTIES)
            return core_props

    def get_or_add_image_part(self, image_file):
        """
        Return an |ImagePart| object containing the image in *image_file*. If
        the image part already exists in this package, it is reused,
        otherwise a new one is created.
        """
        return self._image_parts.get_or_add_image_part(image_file)

    def next_image_partname(self, ext):
        """
        Return a |PackURI| instance representing the next available image
        partname, by sequence number. *ext* is used as the extention on the
        returned partname.
        """
        def first_available_image_idx():
            image_idxs = sorted([
                part.partname.idx for part in self.iter_parts()
                if part.partname.startswith('/ppt/media/image')
            ])
            for i, image_idx in enumerate(image_idxs):
                idx = i + 1
                if idx < image_idx:
                    return idx
            return len(image_idxs)+1

        idx = first_available_image_idx()
        return PackURI('/ppt/media/image%d.%s' % (idx, ext))

    @property
    def presentation_part(self):
        """
        Reference to the |Presentation| instance contained in this package.
        """
        return self.main_document_part

    @lazyproperty
    def _image_parts(self):
        """
        |_ImageParts| object providing access to the image parts in this
        package.
        """
        return _ImageParts(self)


class _ImageParts(object):
    """
    Provides access to the image parts in a package.
    """
    def __init__(self, package):
        super(_ImageParts, self).__init__()
        self._package = package

    def __iter__(self):
        """
        Generate a reference to each |ImagePart| object in the package.
        """
        image_parts = []
        for rel in self._package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype != RT.IMAGE:
                continue
            image_part = rel.target_part
            if image_part in image_parts:
                continue
            image_parts.append(image_part)
            yield image_part

    def get_or_add_image_part(self, image_file):
        """
        Return an |ImagePart| object containing the image in *image_file*,
        which is either a path to an image file or a file-like object
        containing an image. If an image part containing this same image
        already exists, that instance is returned, otherwise a new image part
        is created.
        """
        image = Image.from_file(image_file)
        image_part = self._find_by_sha1(image.sha1)
        if image_part is None:
            image_part = ImagePart.new(self._package, image)
        return image_part

    def _find_by_sha1(self, sha1):
        """
        Return an |ImagePart| object belonging to this package or |None| if
        no matching image part is found. The image part is identified by the
        SHA1 hash digest of the image binary it contains.
        """
        for image_part in self:
            if image_part.sha1 == sha1:
                return image_part
        return None
