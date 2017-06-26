.. :changelog:

Release History
---------------

0.6.6 (2017-06-17)
++++++++++++++++++

- Add `SlideShapes.add_movie()`, allowing video media to be added to a slide.

- fix #190 Accommodate non-conforming part names having '00' index segment.
- fix #273 Accommodate non-conforming part names having no index segment.
- fix #277 ASCII/Unicode error on non-ASCII multi-level category names
- fix #279 BaseShape.id warning appearing on placeholder access.


0.6.5 (2017-03-21)
++++++++++++++++++

- #267 compensate for non-conforming PowerPoint behavior on c:overlay element

- compensate for non-conforming (to spec) PowerPoint behavior related to
  c:dLbl/c:tx that results in "can't save" error when explicit data labels
  are added to bubbles on a bubble chart.


0.6.4 (2017-03-17)
++++++++++++++++++

- add Chart.chart_title and ChartTitle object
- #263 Use Number type to test for numeric category


0.6.3 (2017-02-28)
++++++++++++++++++

- add DataLabel.font
- add Axis.axis_title


0.6.2 (2017-01-03)
++++++++++++++++++

- add support for NotesSlide (slide notes, aka. notes page)
- add support for arbitrary series ordering in XML
- add Plot.categories providing access to hierarchical categories in an
  existing chart.
- add support for date axes on category charts, including writing a dateAx
  element for the category axis when ChartData categories are date or
  datetime.

**BACKWARD INCOMPATIBILITIES:**

Some changes were made to the boilerplate XML used to create new charts. This
was done to more closely adhere to the settings PowerPoint uses when creating
a chart using the UI. This may result in some appearance changes in charts
after upgrading. In particular:

* Chart.has_legend now defaults to True for Line charts.
* Plot.vary_by_categories now defaults to False for Line charts.


0.6.1 (2016-10-09)
++++++++++++++++++

- add Connector shape type


0.6.0 (2016-08-18)
++++++++++++++++++

- add XY chart types
- add Bubble chart types
- add Radar chart types
- add Area chart types
- add Doughnut chart types
- add Series.points and Point
- add Point.data_label
- add DataLabel.text_frame
- add DataLabel.position
- add Axis.major_gridlines
- add ChartFormat with .fill and .line
- add Axis.format (fill and line formatting)
- add ValueAxis.crosses and .crosses_at
- add Point.format (fill and line formatting)
- add Slide.slide_id
- add Slides.get() (by slide id)
- add Font.language_id
- support blank (None) data points in created charts
- add Series.marker
- add Point.marker
- add Marker.format, .style, and .size


0.5.8 (2015-11-27)
++++++++++++++++++

- add Shape.click_action (hyperlink on shape)
- fix: #128 Chart cat and ser names not escaped
- fix: #153 shapes.title raises on no title shape
- fix: #170 remove seek(0) from Image.from_file()


0.5.7 (2015-01-17)
++++++++++++++++++

- add PicturePlaceholder with .insert_picture() method
- add TablePlaceholder with .insert_table() method
- add ChartPlaceholder with .insert_chart() method
- add Picture.image property, returning Image object
- add Picture.crop_left, .crop_top, .crop_right, and .crop_bottom
- add Shape.placeholder_format and PlaceholderFormat object

**BACKWARD INCOMPATIBILITIES:**

Shape.shape_type is now unconditionally `MSO_SHAPE_TYPE.PLACEHOLDER` for all
placeholder shapes. Previously, some placeholder shapes reported
`MSO_SHAPE_TYPE.AUTO_SHAPE`, `MSO_SHAPE_TYPE.CHART`,
`MSO_SHAPE_TYPE.PICTURE`, or `MSO_SHAPE_TYPE.TABLE` for that property.


0.5.6 (2014-12-06)
++++++++++++++++++

- fix #138 - UnicodeDecodeError in setup.py on Windows 7 Python 3.4


0.5.5 (2014-11-17)
++++++++++++++++++

- feature #51 - add Python 3 support


0.5.4 (2014-11-15)
++++++++++++++++++

- feature #43 - image native size in shapes.add_picture() is now calculated
  based on DPI attribute in image file, if present, defaulting to 72 dpi.
- feature #113 - Add Paragraph.space_before, Paragraph.space_after, and
  Paragraph.line_spacing


0.5.3 (2014-11-09)
++++++++++++++++++

- add experimental feature TextFrame.fit_text()


0.5.2 (2014-10-26)
++++++++++++++++++

- fix #127 - Shape.text_frame fails on shape having no txBody


0.5.1 (2014-09-22)
++++++++++++++++++

- feature #120 - add Shape.rotation
- feature #97 - add Font.underline
- issue #117 - add BMP image support
- issue #95 - add BaseShape.name setter
- issue #107 - all .text properties should return unicode, not str
- feature #106 - add .text getters to Shape, TextFrame, and Paragraph

- Rename Shape.textframe to Shape.text_frame.
  **Shape.textframe property (by that name) is deprecated.**


0.5.0 (2014-09-13)
++++++++++++++++++

- Add support for creating and manipulating bar, column, line, and pie charts
- Major refactoring of XML layer (oxml)
- Rationalized graphical object shape access
  **Note backward incompatibilities below**

**BACKWARD INCOMPATIBILITIES:**

A table is no longer treated as a shape. Rather it is a graphical object
contained in a GraphicFrame shape, as are Chart and SmartArt objects.

Example::

    table = shapes.add_table(...)

    # becomes

    graphic_frame = shapes.add_table(...)
    table = graphic_frame.table

    # or

    table = shapes.add_table(...).table

As the enclosing shape, the id, name, shape type, position, and size are
attributes of the enclosing GraphicFrame object.

The contents of a GraphicFrame shape can be identified using three available
properties on a shape: has_table, has_chart, and has_smart_art. The enclosed
graphical object is obtained using the properties GraphicFrame.table and
GraphicFrame.chart. SmartArt is not yet supported. Accessing one of these
properties on a GraphicFrame not containing the corresponding object raises
an exception.


0.4.2 (2014-04-29)
++++++++++++++++++

- fix: issue #88 -- raises on supported image file having uppercase extension
- fix: issue #89 -- raises on add_slide() where non-contiguous existing ids


0.4.1 (2014-04-29)
++++++++++++++++++

- Rename Presentation.slidemasters to Presentation.slide_masters.
  Presentation.slidemasters property is deprecated.
- Rename Presentation.slidelayouts to Presentation.slide_layouts.
  Presentation.slidelayouts property is deprecated.
- Rename SlideMaster.slidelayouts to SlideMaster.slide_layouts.
  SlideMaster.slidelayouts property is deprecated.
- Rename SlideLayout.slidemaster to SlideLayout.slide_master.
  SlideLayout.slidemaster property is deprecated.
- Rename Slide.slidelayout to Slide.slide_layout. Slide.slidelayout property
  is deprecated.
- Add SlideMaster.shapes to access shapes on slide master.
- Add SlideMaster.placeholders to access placeholder shapes on slide master.
- Add _MasterPlaceholder class.
- Add _LayoutPlaceholder class with position and size inheritable from master
  placeholder.
- Add _SlidePlaceholder class with position and size inheritable from layout
  placeholder.
- Add Table.left, top, width, and height read/write properties.
- Add rudimentary GroupShape with left, top, width, and height properties.
- Add rudimentary Connector with left, top, width, and height properties.
- Add TextFrame.auto_size property.
- Add Presentation.slide_width and .slide_height read/write properties.
- Add LineFormat class providing access to read and change line color and
  width.
- Add AutoShape.line
- Add Picture.line

- Rationalize enumerations. **Note backward incompatibilities below**

**BACKWARD INCOMPATIBILITIES:**

The following enumerations were moved/renamed during the rationalization of
enumerations:

- ``pptx.enum.MSO_COLOR_TYPE`` --> ``pptx.enum.dml.MSO_COLOR_TYPE``
- ``pptx.enum.MSO_FILL`` --> ``pptx.enum.dml.MSO_FILL``
- ``pptx.enum.MSO_THEME_COLOR`` --> ``pptx.enum.dml.MSO_THEME_COLOR``
- ``pptx.constants.MSO.ANCHOR_*`` --> ``pptx.enum.text.MSO_ANCHOR.*``
- ``pptx.constants.MSO_SHAPE`` --> ``pptx.enum.shapes.MSO_SHAPE``
- ``pptx.constants.PP.ALIGN_*`` --> ``pptx.enum.text.PP_ALIGN.*``
- ``pptx.constants.MSO.{SHAPE_TYPES}`` -->
  ``pptx.enum.shapes.MSO_SHAPE_TYPE.*``

Documentation for all enumerations is available in the Enumerations section
of the User Guide.


0.3.2 (2014-02-07)
++++++++++++++++++

- Hotfix: issue #80 generated presentations fail to load in Keynote and other
  Apple applications


0.3.1 (2014-01-10)
++++++++++++++++++

- Hotfix: failed to load certain presentations containing images with
  uppercase extension


0.3.0 (2013-12-12)
++++++++++++++++++

- Add read/write font color property supporting RGB, theme color, and inherit
  color types
- Add font typeface and italic support
- Add text frame margins and word-wrap
- Add support for external relationships, e.g. linked spreadsheet
- Add hyperlink support for text run in shape and table cell
- Add fill color and brightness for shape and table cell, fill can also be set
  to transparent (no fill)
- Add read/write position and size properties to shape and picture
- Replace PIL dependency with Pillow
- Restructure modules to better suit size of library


0.2.6 (2013-06-22)
++++++++++++++++++

- Add read/write access to core document properties
- Hotfix to accomodate connector shapes in _AutoShapeType
- Hotfix to allow customXml parts to load when present


0.2.5 (2013-06-11)
++++++++++++++++++

- Add paragraph alignment property (left, right, centered, etc.)
- Add vertical alignment within table cell (top, middle, bottom)
- Add table cell margin properties
- Add table boolean properties: first column (row header), first row (column
  headings), last row (for e.g. totals row), last column (for e.g. row
  totals), horizontal banding, and vertical banding.
- Add support for auto shape adjustment values, e.g. change radius of corner
  rounding on rounded rectangle, position of callout arrow, etc.


0.2.4 (2013-05-16)
++++++++++++++++++

- Add support for auto shapes (e.g. polygons, flowchart symbols, etc.)


0.2.3 (2013-05-05)
++++++++++++++++++

- Add support for table shapes
- Add indentation support to textbox shapes, enabling multi-level bullets on
  bullet slides.


0.2.2 (2013-03-25)
++++++++++++++++++

- Add support for opening and saving a presentation from/to a file-like
  object.
- Refactor XML handling to use lxml objectify


0.2.1 (2013-02-25)
++++++++++++++++++

- Add support for Python 2.6
- Add images from a stream (e.g. StringIO) in addition to a path, allowing
  images retrieved from a database or network resource to be inserted without
  saving first.
- Expand text methods to accept unicode and UTF-8 encoded 8-bit strings.
- Fix potential install bug triggered by importing ``__version__`` from
  package ``__init__.py`` file.


0.2.0 (2013-02-10)
++++++++++++++++++

First non-alpha release with basic capabilities:

- open presentation/template or use built-in default template
- add slide
- set placeholder text (e.g. bullet slides)
- add picture
- add text box
