Use cases
=========

The use case that drove me to begin work on this library has been to automate
the building of slides that are tedious to compose by hand. As an example,
consider the task of composing a slide with an array of 10 headshot images of
folks in a particular department, with the person's name and title next to
their picture. After doing this a dozen times and struggling to get all the
alignment and sizes to the point where my attention to detail is satisfied,
well, my coding fingers got quite itchy.

However I believe a broader application will be server-side document
generation on non-Windows server platforms, Linux primarily I expect. In my
organization, I have found an apparently insatiable demand for PowerPoint
documents as a means of communication. Once one rises beyond the level of
project manager it seems the willingness to interpret text longer than a
bullet point atrophies quite rapidly and PowerPoint becomes an everyday
medium. I've imagined it might be pretty cool to be able to generate a
"presentation-ready" deck for a salesperson that includes a particular subset
of the product catalog they could generate with a few clicks to use in a sales
presentation, for example. As you come up with applications I'd love to hear
about them.


.. _rendering-to-pdf-video-or-image-formats:

Rendering to video, PDF, or image formats
-----------------------------------------

A recurring request — whether phrased as "save the deck as PDF" (`issue #584
<https://github.com/scanny/python-pptx/issues/584>`_), "export to MP4"
(`issue #1049 <https://github.com/scanny/python-pptx/issues/1049>`_), or
"save a slide as an image" (`issue #963
<https://github.com/scanny/python-pptx/issues/963>`_) — is for python-pptx
to produce a rendered output from a deck it has generated.
**python-pptx does not render slides.**
The library reads and writes the ``.pptx``
`Open XML <https://learn.microsoft.com/openspecs/office_standards/ms-offcrypto/>`_
package — it manipulates the XML, images, and other parts that make up the
file, but it has no layout engine, no font rasterizer, and no media encoder.
Producing a PDF (or MP4, PNG, HTML, etc.) from a deck requires a separate
*rendering* application — typically PowerPoint itself or a compatible clone
such as LibreOffice Impress — and rendering is explicitly out of scope for
this library. PDF export in particular is the single most-requested rendered
format; the same integration points apply whether the target is PDF, MP4,
or an image sequence.

If you need to produce rendered output from a deck that python-pptx has
generated, the usual pattern is to build the ``.pptx`` with python-pptx and
then hand it to a rendering tool in a second step. A few common integration
points:

* **LibreOffice (headless) --- recommended on Linux / CI.** The ``soffice``
  binary that ships with LibreOffice can convert to PDF, PNG, and a handful of
  other formats without a GUI, which makes it a natural fit for server-side
  pipelines:

  .. code-block:: shell

      # PDF (one file per deck)
      libreoffice --headless --convert-to pdf deck.pptx --outdir out/

      # PNG (first slide only — see below for per-slide image export)
      libreoffice --headless --convert-to png deck.pptx --outdir out/

  ``--convert-to png`` only writes the **first** slide — that's a LibreOffice
  limitation, not something python-pptx can work around. The standard
  recipe for saving *every* slide as an image (the `issue #963
  <https://github.com/scanny/python-pptx/issues/963>`_ use case) is a
  two-step PDF-then-rasterize pipeline using ``pdftoppm`` (from
  ``poppler-utils``) or ImageMagick's ``convert``:

  .. code-block:: shell

      # 1. deck.pptx -> deck.pdf (one PDF page per slide)
      libreoffice --headless --convert-to pdf deck.pptx --outdir out/

      # 2a. deck.pdf -> slide-1.png, slide-2.png, ... at 150 DPI
      pdftoppm -png -r 150 out/deck.pdf out/slide

      # 2b. or, with ImageMagick (slower, but supports JPG / TIFF / WebP too)
      convert -density 150 out/deck.pdf out/slide-%d.png

  For ``convert``, ``-density`` controls the rasterization DPI and
  ``%d`` expands to the (zero-based) page index. ``pdftoppm`` is usually
  the better choice on Linux / CI — it's faster, has no ``policy.xml``
  PDF restrictions to work around, and produces deterministic filenames.

  LibreOffice does **not** export directly to MP4. The usual route to video
  is to export the deck to a sequence of PNGs (via PDF + ``pdftoppm``, or by
  looping over slides) and then to stitch those frames into a video with
  ``ffmpeg``. Be aware that LibreOffice's renderer is not pixel-identical to
  PowerPoint's — fonts, shape effects, transitions, and embedded media in
  particular may differ. For anything where fidelity matters, validate the
  output against the same deck opened in PowerPoint.

* **PowerPoint itself (Windows / macOS).** If you have a licensed PowerPoint
  install and a Windows host available, PowerPoint can be driven via COM
  automation (``pywin32``) or AppleScript to save as PDF or export an MP4
  (``File -> Export -> Create a Video``) with full fidelity. This is the
  highest-fidelity option because it's the same renderer that end-users see.

* **Aspose.Slides for Python via .NET.** `Aspose.Slides
  <https://products.aspose.com/slides/python-net/>`_ is a commercial library
  with native PPTX rendering, including direct export to PDF, PNG, SVG,
  HTML, and MP4 without requiring PowerPoint or LibreOffice to be installed.
  It is a separate product (not affiliated with python-pptx) and is licensed
  per their terms; mentioned here because it is the most frequently suggested
  drop-in for the "render on a Linux server without LibreOffice" use case.

* **python-pptx-interface.** `python-pptx-interface
  <https://pypi.org/project/python-pptx-interface/>`_ is a third-party wrapper
  that drives PowerPoint on Windows via COM on top of python-pptx, exposing
  convenience methods for saving as PDF and similar operations. It requires
  PowerPoint to be installed on the host and therefore does not help on Linux.

In short: use python-pptx to *produce* the ``.pptx`` and then pick a renderer
based on your platform, fidelity requirements, and licensing constraints.
Pull requests that add a rendering backend to python-pptx itself are out of
scope for the project.


.. _animated-gif-rendering:

Animated GIFs
-------------

A closely related rendering-dependency question is *"why does my animated
GIF only show its first frame after I add it with*
:meth:`~.SlideShapes.add_picture`\ *?"* (see `issue #501
<https://github.com/scanny/python-pptx/issues/501>`_). The answer is the
same: python-pptx writes the ``.pptx`` package but does not render it, and
whether or not a GIF animates is entirely up to the application that later
opens the deck.

The ``<p:pic>`` XML python-pptx emits for a GIF is the same as the XML
PowerPoint itself writes when you drop a ``.gif`` on a slide by hand — a
plain ``<a:blip r:embed="…"/>`` inside a ``<p:blipFill>``. There is **no**
``p:timing`` / ``p:seq`` / ``p:cTn`` entry required to make a GIF loop;
GIF cycling is a *renderer* behaviour keyed off the embedded image's MIME
type, not an OOXML authoring concern. Specifically:

* **PowerPoint for Windows (slideshow mode).** Animates the GIF. The same
  deck in *edit* view intentionally shows only the first frame so the
  editor isn't distracting.
* **PowerPoint for Mac / PowerPoint for the web.** Historically inconsistent;
  some versions animate in slideshow, others do not. When it matters,
  test on the specific target version.
* **LibreOffice Impress.** Does not animate embedded GIFs in slideshow or
  when exporting to PDF/PNG — every frame but the first is discarded by the
  renderer. This is a LibreOffice limitation, not something python-pptx can
  work around in the ``.pptx``.
* **Headless ``libreoffice --convert-to pdf`` / ``--convert-to png``.**
  Same as LibreOffice Impress: first frame only. If you need an animated
  end result, export via PowerPoint's ``Create a Video`` (MP4) path, or
  pre-render the animation yourself (for example via ``ffmpeg`` and then
  insert the resulting MP4 using :meth:`~.SlideShapes.add_movie`).
* **Aspose.Slides / python-pptx-interface / other third-party renderers.**
  Varies by product and version — consult each tool's documentation.

If you need the animation to play on a target renderer that does not
support animated GIFs, the workaround is to convert the GIF to an MP4
(``ffmpeg -i input.gif output.mp4``) and insert it with
:meth:`~.SlideShapes.add_movie` instead. Video playback has proper
``<p:video>`` timing plumbing and is rendered by every major PowerPoint
client and LibreOffice.


.. _saving-pictures-as-svg:

Saving a picture as SVG
-----------------------

A related request (`issue #885
<https://github.com/scanny/python-pptx/issues/885>`_) is *"can python-pptx
save a picture from a slide as an SVG file?"* There are two distinct
variants of this question and they have different answers.

**Extracting an embedded SVG.** |pp| stores every embedded image as opaque
bytes in an image part, addressed by content hash. If the picture on the
slide was authored with SVG content (either by PowerPoint itself, which
stores an SVG alongside a PNG raster fallback — see
:ref:`inserting-svg-images` — or by another tool that dropped an SVG into
the package), the raw SVG bytes are already in the ``.pptx`` and can be
written out verbatim. The blob available on :attr:`.Picture.image` is
whatever the picture's primary ``a:blip`` references, so:

.. code-block:: python

    picture = slide.shapes[0]  # a Picture shape
    image = picture.image
    if image.content_type == "image/svg+xml":
        with open(f"extracted.{image.ext}", "wb") as f:
            f.write(image.blob)

Note the caveat: when the deck was written by PowerPoint, the primary blip
references the **PNG fallback**, not the SVG — so ``image.blob`` in that
case returns PNG bytes. The SVG companion is attached via an
``<asvg:svgBlip>`` drawingML extension on the same blip, referenced by a
separate relationship; |pp| does not currently expose that companion part
through the public API. If you need the SVG, read the ``<asvg:svgBlip>``
relationship ID directly from the picture's XML
(``picture._element.blipFill.blip``) and look up the corresponding
``ImagePart`` via the slide part's relationship table. That is an advanced
use, and the shape of the API may change if native SVG support lands on
the write side (tracked on `issue #652`_).

.. _issue #652: https://github.com/scanny/python-pptx/issues/652

**"Rendering" a raster picture as SVG.** Converting an arbitrary picture
on a slide — a JPEG or PNG, say — into an SVG is **not** something
|pp| does. That is a raster-to-vector operation (tracing), which requires
a separate renderer or vectorizer. The library has no layout engine, no
font rasterizer, and no tracer; see
:ref:`rendering-to-pdf-video-or-image-formats` for the broader rationale.
Sensible external tools for this kind of conversion include:

* **Inkscape (CLI).** ``inkscape --export-type=svg input.png
  --export-filename=out.svg`` performs a raster-to-vector trace;
  ``Path -> Trace Bitmap`` is the interactive equivalent. Inkscape also
  accepts ``.pptx`` indirectly via a LibreOffice-produced PDF (see below).
* **LibreOffice (headless).** ``libreoffice --headless --convert-to svg
  deck.pptx --outdir out/`` emits one SVG per slide, preserving the deck's
  vector content where possible. This is slide-level, not picture-level —
  extract the target shape in a separate post-processing step (e.g., with
  ``xmlstarlet`` against the emitted SVG) if you need just one picture.
* **potrace / autotrace.** Classic open-source bitmap-to-vector tracers
  when the input is a clean line drawing or logo rather than a photograph.

In short: |pp| can hand you the bytes of a picture that *is already* an
SVG in the package, but turning a raster picture *into* an SVG is a
rendering / tracing concern and belongs to an external tool.
