# encoding: utf-8

"""
Part objects, including BasePart.
"""

from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.util import Collection


class BasePart(Part):
    """
    Base class for parts such as Slide and Presentation. Provides common code
    and serves as default class for parts having no custom part class
    defined.
    """
    def __init__(self, partname, content_type, blob=None):
        super(BasePart, self).__init__(partname, content_type, blob)

    @property
    def blob(self):
        """
        Return binary value of this part. Intended to be overridden by
        subclasses. Default is to return load blob.
        """
        return self._blob

    @property
    def content_type(self):
        """
        REMOVE ME, SHOULD BE ABLE TO GET FROM opc.package.Part superclass
        """
        return self._content_type_

    @classmethod
    def load(cls, partname, content_type, blob):
        return cls(partname, content_type, blob)

    @property
    def partname(self):
        """
        |PackURI| instance holding partname of this part, e.g.
        '/ppt/slides/slide1.xml'
        """
        return self._partname

    @partname.setter
    def partname(self, partname):
        self._partname = partname

    @property
    def _content_type(self):
        """
        Content type of this part, e.g.
        'application/vnd.openxmlformats-officedocument.theme+xml'.
        """
        return self._content_type_

    @_content_type.setter
    def _content_type(self, content_type):
        self._content_type_ = content_type


class PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(PartCollection, self).__init__()

    def add_part(self, part):
        """
        Insert a new part into the collection such that list remains sorted
        in logical partname order (e.g. slide10.xml comes after slide9.xml).
        """
        new_partidx = part.partname.idx
        for idx, seq_part in enumerate(self._values):
            partidx = PackURI(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)
