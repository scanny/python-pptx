.. _PpMediaType:

``PP_MEDIA_TYPE``
=================

Indicates the OLE media type.

Example::

    from pptx.enum.shapes import PP_MEDIA_TYPE

    movie = slide.shapes[0]
    assert movie.media_type == PP_MEDIA_TYPE.MOVIE

----

MOVIE
    Video media such as MP4.

OTHER
    Other media types

SOUND
    Audio media such as MP3.

MIXED
    Return value only; indicates multiple media types.
