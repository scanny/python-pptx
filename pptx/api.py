# encoding: utf-8

"""
Directly exposed API classes, Presentation for now. Provides some syntactic
sugar for interacting with the pptx.presentation.Package graph and also
provides some insulation so not so many classes in the other modules need to
be named as internal (leading underscore).
"""

from pptx.presentation import Package


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
    def slidelayouts(self):
        """
        Tuple containing the |SlideLayout| instances belonging to the
        first |SlideMaster| of this presentation.
        """
        return tuple(self._presentation.slidemasters[0].slidelayouts)

    @property
    def slidemaster(self):
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self._presentation.slidemasters[0]

    @property
    def slidemasters(self):
        """
        List of |SlideMaster| objects belonging to this presentation.
        """
        return self._presentation.slidemasters

    @property
    def slides(self):
        """
        |SlideCollection| object containing the slides in this
        presentation.
        """
        return self._presentation.slides

    def save(self, file):
        """
        Save this presentation to *file*, where *file* can be either a path to
        a file (a string) or a file-like object.
        """
        return self._package.save(file)
