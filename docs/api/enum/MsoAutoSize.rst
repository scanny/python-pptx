.. _MsoAutoSize:

``MSO_AUTO_SIZE``
=================

Determines the type of automatic sizing allowed.

The following names can be used to specify the automatic sizing behavior
used to fit a shape's text within the shape bounding box, for example::

    from pptx.enum.text import MSO_AUTO_SIZE

    shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

The word-wrap setting of the text frame interacts with the auto-size setting
to determine the specific auto-sizing behavior.

Note that ``TextFrame.auto_size`` can also be set to |None|, which removes
the auto size setting altogether. This causes the setting to be inherited,
either from the layout placeholder, in the case of a placeholder shape, or
from the theme.

----

NONE
    No automatic sizing of the shape or text will be done. Text can freely
    extend beyond the horizontal and vertical edges of the shape bounding box.

SHAPE_TO_FIT_TEXT
    The shape height and possibly width are adjusted to fit the text. Note
    this setting interacts with the TextFrame.word_wrap property setting. If
    word wrap is turned on, only the height of the shape will be adjusted;
    soft line breaks will be used to fit the text horizontally.

TEXT_TO_FIT_SHAPE
    The font size is reduced as necessary to fit the text within the shape.

MIXED
    Return value only; indicates a combination of automatic sizing schemes are
    used.
