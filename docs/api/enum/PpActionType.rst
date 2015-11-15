.. _PpActionType:

``PP_ACTION_TYPE``
==================

Specifies the type of a mouse action (click or hover action).

Alias: ``PP_ACTION``

Example::

    from pptx.enum.action import PP_ACTION

    assert shape.click_action.action == PP_ACTION.HYPERLINK

----

END_SHOW
    Slide show ends.

FIRST_SLIDE
    Returns to the first slide.

HYPERLINK
    Hyperlink.

LAST_SLIDE
    Moves to the last slide.

LAST_SLIDE_VIEWED
    Moves to the last slide viewed.

NAMED_SLIDE
    Moves to slide specified by slide number.

NAMED_SLIDE_SHOW
    Runs the slideshow.

NEXT_SLIDE
    Moves to the next slide.

NONE
    No action is performed.

OPEN_FILE
    Opens the specified file.

OLE_VERB
    OLE Verb.

PLAY
    Begins the slideshow.

PREVIOUS_SLIDE
    Moves to the previous slide.

RUN_MACRO
    Runs a macro.

RUN_PROGRAM
    Runs a program.
