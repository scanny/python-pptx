============================
Text protocol and interfaces
============================

The protocols and interfaces related to text content, things like text frames,
paragraphs, runs, bold, italic, that sort of thing.

Working directly with the structure of text in OpenXML is pretty fiddly. There
are several layers of hierarchy (textFrame, paragraph, run, text) in the
low-level interface and dealing with their possible states makes for complex
code that tends to obscure intent.

There will have to be one or more higher-level APIs for dealing with text.
However, work has to start somewhere so I'm inclined to start with the direct,
low-level interface and then build up from there.

This page describes the current low-level interface.


.. _text-related-protocol:

Text-related protocol
=====================

Text handling services are provided the following hierarchy of objects, listed
in descending order of granularity.

* Slide
* Shape
* |TextFrame|
* |_Paragraph|
* |_Run|

|TextFrame|
-----------

A text frame contains all the text for a shape. Not all shapes have a text
frame. For example, a picture shape has no text frame. Each text frame
has a set of properties that apply to the text it contains and a sequence of
one or more paragraphs. A shape either has a text frame or it doesn't. A text
frame cannot be added to or deleted from a shape.

The following code sample illustrates the low-level protocol::

    # discovery
    has_textframe = shape.has_textframe
    
    # binding
    textframe = shape.textframe
    paragraphs = textframe.paragraphs
    runs = paragraph[0].runs
    run = runs[0]
    
    # modifying text
    shape.text = 'New text'
    textframe.text = 'New text'
    paragraph.text = 'New text'
    run.text = 'New text'
    
    # adding items
    p = textframe.add_paragraph()
    r = paragraph.add_run()
    
    # deleting items
    paragraph.clear()

