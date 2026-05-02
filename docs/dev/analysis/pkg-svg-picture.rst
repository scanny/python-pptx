.. _pkg_svg_picture:

Inserting SVG images
====================

This analysis documents why |pp| detects SVG content at the image-loading
boundary and raises :class:`.UnsupportedImageTypeError` instead of silently
failing (as it did historically) or attempting to embed the SVG directly.

Issue context
-------------

See `issue #652`_ — "Errors when loading a SVG image into a picture
placeholder". Prior to the fix, passing an ``.svg`` file to
:meth:`.SlideShapes.add_picture` or
:meth:`.PicturePlaceholder.insert_picture` surfaced an opaque
``PIL.UnidentifiedImageError`` from Pillow's format sniffer, which did not
point the caller at the underlying limitation or any workaround.

.. _issue #652: https://github.com/scanny/python-pptx/issues/652

How PowerPoint stores SVG
-------------------------

PowerPoint 2016 (and later) render SVG on a slide but always store the image
as a **pair**:

* a PNG "fallback" raster image (referenced as the primary ``<a:blip
  r:embed="...">`` inside the picture shape's ``<p:blipFill>``), and
* the SVG itself, stored as a separate image part and referenced via an
  ``<a:blipFill>/<a:blip>/<a:extLst>/<a:ext
  uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="..."/>``
  extension (``asvg:`` is the namespace
  ``http://schemas.microsoft.com/office/drawing/2016/SVG/main``).

The fallback PNG is used by legacy viewers (PowerPoint pre-2016, PowerPoint
Online in some modes), by PowerPoint's own thumbnail/preview pipeline, and by
Windows Explorer previews. Without a PNG fallback, PowerPoint shows a broken
picture placeholder when the deck is opened.

Why |pp| cannot synthesize the PNG fallback
-------------------------------------------

Generating the PNG fallback requires rasterizing the SVG. Every available
Python-side rasterizer imposes a substantial dependency footprint:

* **cairosvg** — requires libcairo at runtime.
* **svglib + reportlab** — pure-Python but supports only a subset of SVG
  (notably falls back on complex filters, non-trivial text, and some gradient
  forms).
* **Pillow + librsvg** — requires librsvg (a C library), which is not
  universally available.
* **Inkscape CLI / headless browsers** — external processes, fragile in CI
  and on locked-down desktops.

Adding any of these as a mandatory install-time dependency (or even an
extras-require) would enlarge |pp|'s surface area significantly for a feature
that is cleanly workaroundable by the caller. The project has historically
treated image decoding as "Pillow's job" — see
:class:`pptx.parts.image.Image` — and delegating SVG rasterization would
break that single-source-of-truth pattern.

Chosen behavior
---------------

At the image-loader boundary (``pptx.package._ImageParts.get_or_add_image_part``)
the incoming byte-stream is sniffed for an ``<svg`` root element in its first
2 KiB. Two signals are used:

1. A regex match against the first 2 KiB of the blob, matching ``<svg``
   followed by any of ``\s``, ``>``, or ``/`` so that minimal blobs like
   ``<svg/>`` and pretty-printed documents are both detected.
2. A filename-extension hint (``.svg`` on the path or the ``name`` attribute
   of a file-like object), gated on the head at least looking like XML — this
   catches heavily-commented SVGs whose root tag sits past the 2 KiB window,
   without false-positives on binary files with ``.svg`` names.

When SVG is detected, :class:`.UnsupportedImageTypeError` (a
:class:`.PythonPptxError` subclass) is raised with a message that:

* names SVG explicitly,
* explains that PowerPoint's SVG support requires a raster PNG companion,
* lists concrete pre-rasterize tools the caller can reach for, and
* points at :ref:`inserting-svg-images` in the user guide.

The sniff preserves the stream position for file-like inputs so downstream
consumers (Pillow's format detector, the SHA-1 digester) see the full blob
if future revisions choose to embed the SVG anyway.

The pre-existing ``hasattr(image_part, "sha1")`` guard in
``_ImageParts._find_by_sha1`` is preserved so that a deck which already
contains an SVG part loaded from disk (i.e., a file produced by PowerPoint
itself and round-tripped) still opens cleanly — the defensive skip there
handles the read path, and the new sniff handles the write path.

Future work
-----------

A fully native implementation would:

1. Accept an ``svg_file=`` kwarg on ``add_picture``/``insert_picture``
   alongside the primary raster image, so the caller pre-rasterizes and |pp|
   wires up the pair.
2. Emit the ``<asvg:svgBlip>`` extension inside the picture's ``<a:blip>``
   and relate both image parts to the slide.

This is tracked on `issue #652`_ and is deferred until there is demonstrated
demand beyond the current "give me a clearer error" baseline.
