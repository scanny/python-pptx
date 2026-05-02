
Working with media (video and audio)
====================================

A slide can embed video or audio media using :meth:`SlideShapes.add_movie`.
Although its name suggests video-only use, ``add_movie`` also supports audio
clips (MP3, WAV, M4A, WMA, MIDI, OGG, AIFF, ...) by specifying an
``audio/*`` MIME type.

Adding a video clip
-------------------

.. code-block:: python

    from pptx.util import Emu

    left, top = Emu(914400), Emu(914400)   # 1 in, 1 in
    width, height = Emu(3657600), Emu(2743200)  # 4 in x 3 in
    movie = slide.shapes.add_movie(
        "clip.mp4", left, top, width, height,
        poster_frame_image="thumbnail.png",
        mime_type="video/mp4",
    )

A poster frame is optional; when omitted, python-pptx uses a small
loudspeaker placeholder image.

Adding an audio clip
--------------------

To insert audio, pass an ``audio/*`` MIME type. The returned shape is still
a |Movie| object (reflecting the ``add_movie`` naming), and PowerPoint
renders it with the familiar speaker icon.

.. code-block:: python

    from pptx.util import Emu

    icon_size = Emu(914400)  # 1 in x 1 in
    audio = slide.shapes.add_movie(
        "narration.mp3", Emu(914400), Emu(914400), icon_size, icon_size,
        mime_type="audio/mpeg",
    )

Supported audio MIME types include:

===================   ================
MIME type             Typical extension
===================   ================
``audio/mpeg``        ``.mp3``
``audio/mp3``         ``.mp3``
``audio/mp4``         ``.m4a``
``audio/wav``         ``.wav``
``audio/x-wav``       ``.wav``
``audio/x-ms-wma``    ``.wma``
``audio/midi``        ``.mid``
``audio/aiff``        ``.aiff``
``audio/ogg``         ``.ogg``
``audio/unknown``     (fallback)
===================   ================

Under the hood, when the supplied ``mime_type`` begins with ``audio/``,
python-pptx emits an ``<a:audioFile>`` element inside the shape's
``<p:nvPr>`` rather than the default ``<a:videoFile>``. This matches what
PowerPoint itself writes for audio clips and allows decks that already
contain audio to be opened, modified, and saved without errors.

Limitations
-----------

* The size must be specified explicitly; no auto-scaling is performed.
* The MIME type is not auto-detected from the file contents — always pass it
  explicitly for audio, so the correct ``<a:audioFile>`` element is emitted.
* Advanced audio/video controls such as auto-play, loop, mute, or bookmark
  timings are not yet exposed by the library.
