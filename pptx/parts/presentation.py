# encoding: utf-8

"""
Presentation part, the main part in a .pptx package.
"""

from __future__ import absolute_import

from ..opc.package import XmlPart
from ..presentation import Presentation
from ..util import lazyproperty


class PresentationPart(XmlPart):
    """
    Top level class in object model, represents the contents of the /ppt
    directory of a .pptx file.
    """
    @property
    def core_properties(self):
        """
        A |CoreProperties| object providing read/write access to the core
        properties of this presentation.
        """
        return self.package.core_properties

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    def save(self, path_or_stream):
        """
        Save this presentation package to *path_or_stream*, which can be
        either a path to a filesystem location (a string) or a file-like
        object.
        """
        self.package.save(path_or_stream)
