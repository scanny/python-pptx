
Working with placeholders
=========================

Placeholders can make adding content a lot easier. If you've ever added a new
textbox to a slide from scratch and noticed how many adjustments it took to
get it the way you wanted you understand why. The placeholder is in the right
position with the right font size, paragraph alignment, bullet style, etc.,
etc. Basically you can just click and type in some text and you've got
a slide.


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
attribute, but accessing it on a non-placeholder shape raises |ValueError|::

    >>> for shape in slide.placeholders:
    ...     phf = shape.placeholder_format
    ...     print('%d, %s' % (phf.idx, phf.type))
    ...
    0, TITLE (1)
    1, PICTURE (18)
    2, BODY (2)


Setting the slide title
-----------------------

Almost all slide layouts have a title placeholder, which any slide based on
the layout inherits when the layout is applied. Accessing a slide's title is
a common operation and there's a dedicated attribute on the shape tree for
it::

    title_placeholder = slide.shapes.title
    title_placeholder.text = 'Air-speed Velocity of Unladen Swallows'


Locating other placehoders
--------------------------

The title placeholder is always in the same spot, that's one reason there can
be a dedicated attribute for it. The other placeholders can move around,
although it's usually not hard to sort out which is which.

A slide's placeholders are in its ``placeholders`` attribute, which supports
indexed access, len(), and iteration.

::

    >>> placeholders = slide.placeholders
    >>> len(placeholders)
    3
    >>> title = shapes.title
    >>> assert title is placeholders[0]
    >>> body = placeholders[1]
    >>> body.text = 'Distinguish carefully between African and European swallows'
    
The title placeholder, if present, is guaranteed to be first in the sequence.
A body content shape, if present, is likely to be second. After that, you'll
probably need to experiment to determine which is which::

    for idx, ph in enumerate(shapes.placeholders):
        ph.text = "placeholders[%d]" % idx


Up next ...
-----------
This code will produce a slide with a single bullet. 
Let's take a closer look at how text is added to shapes.
