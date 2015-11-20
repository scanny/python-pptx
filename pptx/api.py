# encoding: utf-8

"""
Directly exposed API classes, Presentation for now. Provides some syntactic
sugar for interacting with the pptx.presentation.Package graph and also
provides some insulation so not so many classes in the other modules need to
be named as internal (leading underscore).
"""

from __future__ import absolute_import, print_function, unicode_literals

from warnings import warn

from pptx.package import Package


class Presentation(object):
    """
    Return a |Presentation| instance loaded from *file_*, where *file_* can
    be either a path to a ``.pptx`` file (a string) or a file-like object.
    If *file_* is missing or ``None``, load the built-in default presentation
    template.
    """
    def __init__(self, pkg_file=None):
        super(Presentation, self).__init__()
        self._package = Package.open(pkg_file)
        self._presentation = self._package.presentation

    @property
    def core_properties(self):
        """
        Instance of |CoreProperties| holding the read/write Dublin Core
        document properties for this presentation.
        """
        return self._package.core_properties

    @property
    def slide_layouts(self):
        """
        Sequence of |SlideLayout| instances belonging to the first
        |SlideMaster| of this presentation.
        """
        return self._presentation.slide_masters[0].slide_layouts

    @property
    def slidelayouts(self):
        """
        Deprecated. Use ``.slide_layouts`` property instead.
        """
        msg = (
            'Presentation.slidelayouts property is deprecated. Use .slide_la'
            'youts instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_layouts

    @property
    def slide_master(self):
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self._presentation.slide_masters[0]

    @property
    def slidemaster(self):
        """
        Deprecated. Use ``.slide_master`` property instead.
        """
        msg = (
            'Presentation.slidemaster property is deprecated. Use .slide_m'
            'aster instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_masters

    @property
    def slide_masters(self):
        """
        List of |SlideMaster| objects belonging to this presentation.
        """
        return self._presentation.slide_masters

    @property
    def slidemasters(self):
        """
        Deprecated. Use ``.slide_masters`` property instead.
        """
        msg = (
            'Presentation.slidemasters property is deprecated. Use .slide_'
            'masters instead.'
        )
        warn(msg, UserWarning, stacklevel=2)
        return self.slide_masters

    @property
    def slide_height(self):
        """
        Height of slides in this presentation, in English Metric Units (EMU)
        """
        return self._presentation.slide_height

    @slide_height.setter
    def slide_height(self, height):
        self._presentation.slide_height = height

    @property
    def slide_width(self):
        """
        Width of slides in this presentation, in English Metric Units (EMU)
        """
        return self._presentation.slide_width

    @slide_width.setter
    def slide_width(self, width):
        self._presentation.slide_width = width

    @property
    def slides(self):
        """
        |_Slides| object containing the slides in this presentation.
        """
        return self._presentation.slides

    def save(self, file):
        """
        Save this presentation to *file*, where *file* can be either a path to
        a file (a string) or a file-like object.
        """
        return self._package.save(file)
