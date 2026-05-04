Editable SVG pictures
=====================

PowerPoint 365 (Office 2016+) supports inserting **SVG** imagery such
that the file opens as a fully editable vector graphic inside
PowerPoint, while still rendering reliably on older clients that do
not know about SVG. It does this by storing *both* the original SVG
*and* a rasterized PNG fallback inside the ``.pptx`` package and
wiring them together with a small Microsoft extension on the
``<a:blip>`` element:

.. code-block:: xml

    <a:blip r:embed="<png-rId>">
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip r:embed="<svg-rId>"/>
        </a:ext>
      </a:extLst>
    </a:blip>

The PNG in ``r:embed`` is what every renderer falls back to; the
``asvg:svgBlip`` inside the extension is what PowerPoint 365+ picks
up to display the editable SVG.

The convenience method
----------------------

python-pptx exposes this as
:meth:`.SlideShapes.add_picture_svg`::

    from io import BytesIO
    from pptx import Presentation
    from pptx.util import Inches

    svg_bytes = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' \
                b'<circle cx="50" cy="50" r="40" fill="#4B8BBE"/></svg>'

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pic = slide.shapes.add_picture_svg(
        BytesIO(svg_bytes),
        Inches(1), Inches(1), Inches(3), Inches(3),
        png_fallback="logo.png",           # optional
    )
    prs.save("svg-picture.pptx")

The call returns a |Picture| proxy exactly like
:meth:`~pptx.shapes.shapetree.SlideShapes.add_picture`.

Why the PNG fallback is required
--------------------------------

PowerPoint's SVG round-trip relies on a companion PNG being present.
Without it, older clients (and the thumbnail preview used by file
managers and mail clients) have nothing to render. python-pptx does
**not** bundle an SVG rasterizer — pulling in a rendering toolkit such
as ``cairosvg`` or ``svglib`` + ``reportlab`` as a runtime dependency
would bloat the library for callers who never use this feature.

Two ways to supply the fallback:

1. **Pre-rasterize yourself** and pass the PNG path or bytes via the
   ``png_fallback`` keyword. This is strongly recommended when the
   presentation will be opened on older clients or viewed as a
   thumbnail — the PNG is what those viewers draw. Any Pillow-
   supported raster (PNG, JPEG, …) works but PNG matches what
   PowerPoint itself emits.

2. **Omit ``png_fallback``** and python-pptx inserts a built-in
   ``1x1`` transparent placeholder PNG. The SVG still renders
   correctly in PowerPoint 365 and later, but older clients fall back
   to the placeholder and show a tiny transparent square. Use this
   path when you know the audience runs PowerPoint 365 / newer Keynote
   / the Office web viewer (all of which pick the SVG).

Editing and round-trip
----------------------

Opening a ``.pptx`` produced this way in PowerPoint 365 and
double-clicking the picture opens the built-in SVG editor — the
original vector content is editable. Saving the file preserves both
the PNG and the SVG parts, so the pair survives round-trip through
PowerPoint itself.

``Picture.image`` on the returned proxy refers to the PNG fallback
(matching the ``r:embed`` on ``a:blip``); the companion SVG bytes are
stored on a separate ``/ppt/media/imageN.svg`` part in the package.

.. versionadded:: 2026.05.0
