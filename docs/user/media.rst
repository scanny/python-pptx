
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

Auto-playing a movie
--------------------

By default, PowerPoint waits for a user click before a movie (or audio
clip) starts playing. Pass ``autoplay=True`` to have the clip begin
automatically when the slide is shown (equivalent to
PowerPoint's "Start: Automatically" ribbon option):

.. code-block:: python

    movie = slide.shapes.add_movie(
        "clip.mp4", left, top, width, height,
        poster_frame_image="thumbnail.png",
        mime_type="video/mp4",
        autoplay=True,
    )

Internally this sets :attr:`Movie.start_condition` to ``"withPrevious"``.
For finer-grained control — e.g. starting *after* the previous effect
finishes, or with a delay — assign to :attr:`Movie.start_condition` and
:attr:`Movie.start_time` directly on the returned movie:

.. code-block:: python

    movie = slide.shapes.add_movie("clip.mp4", ..., autoplay=True)
    movie.start_condition = "afterPrevious"
    movie.start_time = 2.5  # seconds

Recipe: per-slide audio narration
---------------------------------

Audio narration — a voice-over that plays automatically when the slide is
shown, is hidden from the audience, and advances the slide when the clip
ends — is a composition of three public APIs the library already exposes:

* :meth:`SlideShapes.add_movie` with an ``audio/*`` MIME type and
  ``autoplay=True`` to embed the clip and have it start automatically.
* :attr:`BaseShape.is_hidden` to mark the speaker-icon shape hidden so it
  doesn't draw over the slide content during presentation.
* :attr:`Transition.advance_after_time` to auto-advance the slide once the
  narration has finished (supply the clip duration in milliseconds).

.. code-block:: python

    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # embed the narration and start it with the slide ------------------
    icon = Emu(228600)  # 0.25 in; small because it will be hidden
    narration = slide.shapes.add_movie(
        "narration.mp3",
        left=Emu(0), top=Emu(0), width=icon, height=icon,
        mime_type="audio/mpeg",
        autoplay=True,
    )

    # hide the speaker icon from the audience --------------------------
    narration.is_hidden = True

    # optional: auto-advance after the clip's duration (in milliseconds)
    slide.transition.advance_after_time = 5000  # 5.0 seconds

Deriving ``advance_after_time`` from the audio file itself requires a
duration parser (``mutagen``, ``wave``, ``tinytag``, ``ffprobe``, ...).
python-pptx intentionally does not add such a runtime dependency — supply
the duration yourself, in milliseconds, from whichever tool your pipeline
already uses.

To place the (still-hidden) shape off-slide instead of in the corner, set
``left`` and ``top`` to negative values — PowerPoint preserves the offset
on save and never renders a hidden shape anyway, so the positioning is a
matter of author preference.

Adding a URL-linked (online) video
----------------------------------

.. versionadded:: 2026.05.0

PowerPoint supports embedding a video that lives on an external URL —
YouTube, Vimeo, or any HTTPS host that serves a playable video — via its
"Insert Online Video" command. The video bytes are **not** copied into
the ``.pptx``; instead the shape carries an external relationship to the
URL, which PowerPoint's media player loads at presentation time.

:meth:`SlideShapes.add_movie_link` emits the same shape. A poster-frame
image is required because there is no media part from which to derive a
default loudspeaker graphic:

.. code-block:: python

    from pptx.util import Inches

    movie = slide.shapes.add_movie_link(
        "https://www.youtube.com/embed/dQw4w9WgXcQ",
        "poster.png",
        Inches(1), Inches(1), Inches(4), Inches(3),
    )

Choosing the URL form
~~~~~~~~~~~~~~~~~~~~~

PowerPoint's online-video player requires a URL it can actually load
inside a slide. For YouTube, that means the ``/embed/`` form, **not** the
ordinary ``watch?v=…`` address. The URL you'd paste into a browser's
address bar won't play in-slide — it only works on youtube.com itself.

=================================================================  ===========
URL pattern                                                        Plays in
                                                                   PowerPoint?
=================================================================  ===========
``https://www.youtube.com/embed/VIDEO_ID``                         yes
``https://www.youtube.com/watch?v=VIDEO_ID``                       no
``https://youtu.be/VIDEO_ID``                                      no
``https://player.vimeo.com/video/VIDEO_ID``                        yes
``https://example.com/path/to/clip.mp4`` (direct HTTPS, CORS-OK)   yes
=================================================================  ===========

Any ``watch?v=VIDEO_ID`` or ``youtu.be/VIDEO_ID`` URL you already have is
trivial to convert — keep the ``VIDEO_ID`` and substitute it into the
``/embed/`` form. ``add_movie_link`` does **not** validate the URL; the
caller is responsible for supplying a playable form.

Trade-offs vs :meth:`~SlideShapes.add_movie`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* **Package size**: URL-linked video keeps the ``.pptx`` small — only the
  poster PNG lives in the package. Embedded video (``add_movie``) inlines
  the full media bytes and can easily exceed attachment limits.
* **Offline playback**: URL-linked video requires an internet connection
  at presentation time. Embedded video always plays offline.
* **Updates**: the URL-linked shape tracks whatever the host serves, so
  the author can update the video without re-saving the deck. Embedded
  video is pinned to the bytes at save time.
* **Host policy**: YouTube, Vimeo and friends can withdraw, age-gate, or
  change the video independently of your deck — if long-term archival of
  the exact content matters, embed.

``poster_frame_image`` is required: unlike ``add_movie`` there is no
media part to fall back to for a default speaker-icon poster. Supply any
path or binary file-like object accepted by
:meth:`~SlideShapes.add_picture`.

Extracting an embedded media clip
---------------------------------

A |Movie| shape exposes the underlying media bytes directly, mirroring the
``Picture.image.blob`` / ``.ext`` / ``.content_type`` surface on
|Picture|. This is useful when you need to extract the embedded audio or
video from an existing presentation — for transcription, re-encoding, or
republishing — without reaching into the package internals.

.. code-block:: python

    from pptx import Presentation

    prs = Presentation("deck-with-movie.pptx")
    movie = prs.slides[0].shapes[0]  # assuming shape 0 is a Movie

    # write the embedded clip out to disk using its native extension
    with open(f"clip.{movie.ext}", "wb") as f:
        f.write(movie.blob)

    print(movie.content_type)  # e.g. 'video/mp4'

All three accessors return |None| when the shape has no associated media
part (a malformed file), so callers can test with a simple ``if
movie.blob is not None`` guard.

Limitations
-----------

* The size must be specified explicitly; no auto-scaling is performed.
* The MIME type is not auto-detected from the file contents — always pass it
  explicitly for audio, so the correct ``<a:audioFile>`` element is emitted.
* Advanced audio/video controls such as loop, mute, or bookmark timings
  are not yet exposed by the library.
