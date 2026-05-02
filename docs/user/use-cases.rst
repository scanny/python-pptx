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


Rendering to video, PDF, or image formats
-----------------------------------------

A recurring request (see for example `issue #1049
<https://github.com/scanny/python-pptx/issues/1049>`_) is for python-pptx to
export a deck to MP4, PDF, PNG, or similar "rendered" outputs. **python-pptx
does not render slides.** The library reads and writes the ``.pptx``
`Open XML <https://learn.microsoft.com/openspecs/office_standards/ms-offcrypto/>`_
package — it manipulates the XML, images, and other parts that make up the
file, but it has no layout engine, no font rasterizer, and no media encoder.
Producing an MP4 (or PDF, PNG, HTML, etc.) from a deck requires a separate
*rendering* application — typically PowerPoint itself or a compatible clone
such as LibreOffice Impress — and rendering is explicitly out of scope for
this library.

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

      # PNG (first slide only — loop through pages / use a PDF->PNG step
      # like pdftoppm or ImageMagick if you need every slide as an image)
      libreoffice --headless --convert-to png deck.pptx --outdir out/

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
