
Working with placeholders
=========================

Placeholders make adding content a lot easier. If you've ever added a new
textbox to a slide from scratch and noticed how many adjustments it took to get
it the way you wanted you understand why. The placeholder is in the right
position with the right font size, paragraph alignment, bullet style, etc.,
etc. Basically you can just click and blow in some text and you've got a slide.

The same is true of using a placeholder shape in |pp|. Once you have
a reference to the placeholder shape, pretty much the only thing to do is set
its ``.text`` attribute.


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
