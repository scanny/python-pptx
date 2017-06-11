# encoding: utf-8

"""Shapes based on the `p:pic` element, including Picture and Movie."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from .base import BaseShape
from ..dml.line import LineFormat
from ..enum.shapes import MSO_SHAPE_TYPE, PP_MEDIA_TYPE
from ..shared import ParentedElementProxy
from ..util import lazyproperty


class _BasePicture(BaseShape):
    """Base class for shapes based on a `p:pic` element."""

    def __init__(self, pic, parent):
        super(_BasePicture, self).__init__(pic, parent)
        self._pic = pic

    @property
    def crop_bottom(self):
        """
        A |float| representing the relative portion cropped from the bottom
        of this shape where 1.0 represents 100%. For example, 25% is
        represented by 0.25. Negative values are valid as are values greater
        than 1.0.
        """
        return self._element.srcRect_b

    @property
    def crop_left(self):
        """
        A |float| representing the relative portion cropped from the left
        side of this shape where 1.0 represents 100%.
        """
        return self._element.srcRect_l

    @property
    def crop_right(self):
        """
        A |float| representing the relative portion cropped from the right
        side of this shape where 1.0 represents 100%.
        """
        return self._element.srcRect_r

    @property
    def crop_top(self):
        """
        A |float| representing the relative portion cropped from the top of
        this shape where 1.0 represents 100%.
        """
        return self._element.srcRect_t

    def get_or_add_ln(self):
        """
        Return the `a:ln` element containing the line format properties XML
        for this `p:pic`-based shape.
        """
        return self._pic.get_or_add_ln()

    @lazyproperty
    def line(self):
        """
        An instance of |LineFormat|, providing access to the properties of
        the outline bordering this shape, such as its color and width.
        """
        return LineFormat(self)

    @property
    def ln(self):
        """
        The ``<a:ln>`` element containing the line format properties such as
        line color and width. |None| if no ``<a:ln>`` element is present.
        """
        return self._pic.ln


class Movie(_BasePicture):
    """A movie shape, one that places a video on a slide.

    Like |Picture|, a movie shape is based on the `p:pic` element. A movie is
    composed of a video and a *poster frame*, the placeholder image that
    represents the video before it is played.
    """

    @lazyproperty
    def media_format(self):
        """The |_MediaFormat| object for this movie.

        The |_MediaFormat| object provides access to formatting properties
        for the movie.
        """
        return _MediaFormat(self._element, self)

    @property
    def media_type(self):
        """Member of :ref:`PpMediaType` describing this shape.

        The return value is unconditionally `PP_MEDIA_TYPE.MOVIE` in this
        case.
        """
        return PP_MEDIA_TYPE.MOVIE

    @property
    def poster_frame(self):
        """Return |Image| object containing poster frame for this movie.

        Returns |None| if this movie has no poster frame (uncommon).
        """
        slide_part, rId = self.part, self._element.blip_rId
        if rId is None:
            return None
        return slide_part.get_image(rId)

    @property
    def shape_type(self):
        """Return member of :ref:`MsoShapeType` describing this shape.

        The return value is unconditionally ``MSO_SHAPE_TYPE.MEDIA`` in this
        case.
        """
        return MSO_SHAPE_TYPE.MEDIA


class Picture(_BasePicture):
    """A picture shape, one that places an image on a slide.

    Based on the `p:pic` element.
    """

    @property
    def image(self):
        """
        An |Image| object providing access to the properties and bytes of the
        image in this picture shape.
        """
        slide_part, rId = self.part, self._element.blip_rId
        if rId is None:
            raise ValueError('no embedded image')
        return slide_part.get_image(rId)

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, unconditionally
        ``MSO_SHAPE_TYPE.PICTURE`` in this case.
        """
        return MSO_SHAPE_TYPE.PICTURE


class _MediaFormat(ParentedElementProxy):
    """Provides access to formatting properties for a Media object.

    Media format properties are things like start point, volume, and
    compression type.
    """
