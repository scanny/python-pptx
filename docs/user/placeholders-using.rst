.. _placeholders-using:

Working with placeholders
=========================

Placeholders can make adding content a lot easier. If you've ever added a new
textbox to a slide from scratch and noticed how many adjustments it took to
get it the way you wanted you understand why. The placeholder is in the right
position with the right font size, paragraph alignment, bullet style, etc.,
etc. Basically you can just click and type in some text and you've got
a slide.

A placeholder can be also be used to place a rich-content object on a slide.
A picture, table, or chart can each be inserted into a placeholder and so
take on the position and size of the placeholder, as well as certain
of its formatting attributes.


Access a placeholder
--------------------

Every placeholder is also a shape, and so can be accessed using the
:attr:`~.Slide.shapes` property of a slide. However, when looking for
a particular placeholder, the :attr:`~.Slide.placeholders` property can make
things easier.

The most reliable way to access a known placeholder is by its
:attr:`~.PlaceholderFormat.idx` value. The :attr:`idx` value of a placeholder
is the integer key of the slide layout placeholder it inherits properties
from. As such, it remains stable throughout the life of the slide and will be
the same for any slide created using that layout.

It's usually easy enough to take a look at the placeholders on a slide and
pick out the one you want::

    >>> prs = Presentation()
    >>> slide = prs.slides.add_slide(prs.slide_layouts[8])
    >>> for shape in slide.placeholders:
    ...     print('%d %s' % (shape.placeholder_format.idx, shape.name))
    ...
    0  Title 1
    1  Picture Placeholder 2
    2  Text Placeholder 3

... then, having the known index in hand, to access it directly::

    >>> slide.placeholders[1]
    <pptx.parts.slide.PicturePlaceholder object at 0x10d094590>
    >>> slide.placeholders[2].name
    'Text Placeholder 3'

.. note:: Item access on the placeholders collection is like that of
   a dictionary rather than a list. While the key used above is an integer,
   the lookup is on `idx` values, not position in a sequence. If the provided
   value does not match the `idx` value of one of the placeholders,
   |KeyError| will be raised. `idx` values are not necessarily contiguous.

In general, the :attr:`idx` value of a placeholder from a built-in slide
layout (one provided with PowerPoint) will be between 0 and 5. The title
placeholder will always have :attr:`idx` 0 if present and any other
placeholders will follow in sequence, top to bottom and left to right.
A placeholder added to a slide layout by a user in PowerPoint will receive an
:attr:`idx` value starting at 10.


Identify and Characterize a placeholder
---------------------------------------

A placeholder behaves differently that other shapes in some ways. In
particular, the value of its :attr:`~.BaseShape.shape_type` attribute is
unconditionally ``MSO_SHAPE_TYPE.PLACEHOLDER`` regardless of what type of
placeholder it is or what type of content it contains::

    >>> prs = Presentation()
    >>> slide = prs.slides.add_slide(prs.slide_layouts[8])
    >>> for shape in slide.shapes:
    ...     print('%s' % shape.shape_type)
    ...
    PLACEHOLDER (14)
    PLACEHOLDER (14)
    PLACEHOLDER (14)

To find out more, it's necessary to inspect the contents of the placeholder's
:attr:`~.BaseShape.placeholder_format` attribute. All shapes have this
attribute, but accessing it on a non-placeholder shape raises |ValueError|.
The :attr:`~.BaseShape.is_placeholder` attribute can be used to determine
whether a shape is a placeholder::

    >>> for shape in slide.shapes:
    ...     if shape.is_placeholder:
    ...         phf = shape.placeholder_format
    ...         print('%d, %s' % (phf.idx, phf.type))
    ...
    0, TITLE (1)
    1, PICTURE (18)
    2, BODY (2)

Another way a placeholder acts differently is that it inherits its position
and size from its layout placeholder. This inheritance is overridden if the
position and size of a placeholder are changed.


Insert content into a placeholder
---------------------------------

Certain placeholder types have specialized methods for inserting content. In
the current release, the `picture`, `table`, and `chart` placeholders have
content insertion methods. Text can be inserted into `title` and `body`
placeholders in the same way text is inserted into an auto shape.

:meth:`.PicturePlaceholder.insert_picture`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The picture placeholder has an :meth:`~.PicturePlaceholder.insert_picture`
method::

    >>> prs = Presentation()
    >>> slide = prs.slides.add_slide(prs.slide_layouts[8])
    >>> placeholder = slide.placeholders[1]  # idx key, not position
    >>> placeholder.name
    'Picture Placeholder 2'
    >>> placeholder.placeholder_format.type
    PICTURE (18)
    >>> picture = placeholder.insert_picture('my-image.png')

.. note:: A reference to a picture placeholder becomes invalid after its
   :meth:`~.PicturePlaceholder.insert_picture` method is called. This is
   because the process of inserting a picture replaces the original `p:sp`
   XML element with a new `p:pic` element containing the picture. Any attempt
   to use the original placeholder reference after the call will raise
   |AttributeError|. The new placeholder is the return value of the
   :meth:`insert_picture` call and may also be obtained from the placeholders
   collection using the same `idx` key.

A picture inserted in this way is stretched proportionately and cropped to
fill the entire placeholder. Best results are achieved when the aspect ratio
of the source image and placeholder are the same. If the picture is taller
in aspect than the placeholder, its top and bottom are cropped evenly to fit.
If it is wider, its left and right sides are cropped evenly. Cropping can be
adjusted using the crop properties on the placeholder, such as
:attr:`~.PlaceholderPicture.crop_bottom`.

:meth:`.TablePlaceholder.insert_table`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The table placeholder has an :meth:`~.TablePlaceholder.insert_table` method.
The built-in template has no layout containing a table placeholder, so this
example assumes a starting presentation named
``having-table-placeholder.pptx`` having a table placeholder with idx 10 on
its second slide layout::

    >>> prs = Presentation('having-table-placeholder.pptx')
    >>> slide = prs.slides.add_slide(prs.slide_layouts[1])
    >>> placeholder = slide.placeholders[10]  # idx key, not position
    >>> placeholder.name
    'Table Placeholder 1'
    >>> placeholder.placeholder_format.type
    TABLE (12)
    >>> graphic_frame = placeholder.insert_table(rows=2, cols=2)
    >>> table = graphic_frame.table
    >>> len(table.rows), len(table.columns)
    (2, 2)

A table inserted in this way has the position and width of the original
placeholder. Its height is proportional to the number of rows.

Like all rich-content insertion methods, a reference to a table placeholder
becomes invalid after its :meth:`~.TablePlaceholder.insert_table` method is
called. This is because the process of inserting rich content replaces the
original `p:sp` XML element with a new element, a `p:graphicFrame` in this
case, containing the rich-content object. Any attempt to use the original
placeholder reference after the call will raise |AttributeError|. The new
placeholder is the return value of the :meth:`insert_table` call and may also
be obtained from the placeholders collection using the original `idx` key, 10
in this case.

.. note:: The return value of the :meth:`~.TablePlaceholder.insert_table`
   method is a |PlaceholderGraphicFrame| object, which has all the properties
   and methods of a |GraphicFrame| object along with those specific to
   placeholders. The inserted table is contained in the graphic frame and can
   be obtained using its :attr:`~.PlaceholderGraphicFrame.table` property.

:meth:`.ChartPlaceholder.insert_chart`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The chart placeholder has an :meth:`~.ChartPlaceholder.insert_chart` method.
The presentation template built into |pp| has no layout containing a chart
placeholder, so this example assumes a starting presentation named
``having-chart-placeholder.pptx`` having a chart placeholder with idx 10 on
its second slide layout::

    >>> from pptx.chart.data import ChartData
    >>> from pptx.enum.chart import XL_CHART_TYPE

    >>> prs = Presentation('having-chart-placeholder.pptx')
    >>> slide = prs.slides.add_slide(prs.slide_layouts[1])

    >>> placeholder = slide.placeholders[10]  # idx key, not position
    >>> placeholder.name
    'Chart Placeholder 9'
    >>> placeholder.placeholder_format.type
    CHART (12)

    >>> chart_data = ChartData()
    >>> chart_data.categories = ['Yes', 'No']
    >>> chart_data.add_series('Series 1', (42, 24))

    >>> graphic_frame = placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data)
    >>> chart = graphic_frame.chart
    >>> chart.chart_type
    PIE (5)

A chart inserted in this way has the position and size of the original
placeholder.

Note the return value from :meth:`~.ChartPlaceholder.insert_chart` is
a |PlaceholderGraphicFrame| object, not the chart itself.
A |PlaceholderGraphicFrame| object has all the properties and methods of
a |GraphicFrame| object along with those specific to placeholders. The
inserted chart is contained in the graphic frame and can be obtained using
its :attr:`~.PlaceholderGraphicFrame.chart` property.

Like all rich-content insertion methods, a reference to a chart placeholder
becomes invalid after its :meth:`~.ChartPlaceholder.insert_chart` method is
called. This is because the process of inserting rich content replaces the
original `p:sp` XML element with a new element, a `p:graphicFrame` in this
case, containing the rich-content object. Any attempt to use the original
placeholder reference after the call will raise |AttributeError|. The new
placeholder is the return value of the :meth:`insert_chart` call and may also
be obtained from the placeholders collection using the original `idx` key, 10
in this case.


Setting the slide title
-----------------------

Almost all slide layouts have a title placeholder, which any slide based on
the layout inherits when the layout is applied. Accessing a slide's title is
a common operation and there's a dedicated attribute on the shape tree for
it::

    title_placeholder = slide.shapes.title
    title_placeholder.text = 'Air-speed Velocity of Unladen Swallows'
