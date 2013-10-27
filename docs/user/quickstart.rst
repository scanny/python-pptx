Getting Started
===============

A quick way to get started is by trying out some of :ref:`the examples below
<examples>` to get a feel for how to use |pp|.

The :ref:`user API documentation <api>` can help you with the fine details of
calling signatures and behaviors.

.. _examples:


Examples
========

Examples to help with getting started ...

.. highlight:: python
   :linenothreshold: 5

----

Hello World! example
--------------------

.. image:: /_static/img/hello-world.png

|

::

    from pptx import Presentation

    prs = Presentation()
    title_slidelayout = prs.slidelayouts[0]
    slide = prs.slides.add_slide(title_slidelayout)
    title = slide.shapes.title
    subtitle = slide.shapes.placeholders[1]

    title.text = "Hello, World!"
    subtitle.text = "python-pptx was here!"

    prs.save('test.pptx')


----

Bullet slide example
--------------------

.. image:: /_static/img/bullet-slide.png

|

::

    from pptx import Presentation

    prs = Presentation()
    bullet_slidelayout = prs.slidelayouts[1]

    slide = prs.slides.add_slide(bullet_slidelayout)
    shapes = slide.shapes

    title_shape = shapes.title
    body_shape = shapes.placeholders[1]

    title_shape.text = 'Adding a Bullet Slide'

    tf = body_shape.textframe
    tf.text = 'Find the bullet slide layout'

    p = tf.add_paragraph()
    p.text = 'Use _TextFrame.text for first bullet'
    p.level = 1

    p = tf.add_paragraph()
    p.text = 'Use _TextFrame.add_paragraph() for subsequent bullets'
    p.level = 2

    prs.save('test.pptx')

Not all shapes can contain text, but those that do always have at least one
paragraph, even if that paragraph is empty and no text is visible within the
shape. ``_BaseShape.has_textframe`` can be used to determine whether a shape
can contain text. (All shapes subclass ``_BaseShape``.) When
``_BaseShape.has_textframe`` is ``True``,
``_BaseShape.textframe.paragraphs[0]`` returns the first paragraph. The text
of the first paragraph can be set using ``textframe.paragraphs[0].text``. As a
shortcut, the writable properties ``_BaseShape.text`` and ``_TextFrame.text``
are provided to accomplish the same thing. Note that these last two calls
delete all the shape's paragraphs except the first one before setting the text
it contains.

----

``add_textbox()`` example
-------------------------

.. image:: /_static/img/add-textbox.png

|

::

    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    blank_slidelayout = prs.slidelayouts[6]
    slide = prs.slides.add_slide(blank_slidelayout)

    left = top = width = height = Inches(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.textframe

    tf.text = "This is text inside a textbox"

    p = tf.add_paragraph()
    p.text = "This is a second paragraph that's bold"
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "This is a third paragraph that's big"
    p.font.size = Pt(40)

    prs.save('test.pptx')


----

``add_picture()`` example
-------------------------

.. image:: /_static/img/add-picture.png

|

::

    from pptx import Presentation
    from pptx.util import Inches, Px

    img_path = 'monty-truth.png'

    prs = Presentation()
    blank_slidelayout = prs.slidelayouts[6]
    slide = prs.slides.add_slide(blank_slidelayout)

    left = top = Inches(1)
    pic = slide.shapes.add_picture(img_path, left, top)

    left = Inches(5)
    width = Px(280)
    height = int(width*1.427)
    pic = slide.shapes.add_picture(img_path, left, top, width, height)

    prs.save('test.pptx')


----


``add_shape()`` example
-------------------------

.. image:: /_static/img/add-shape.png

|

::

    from pptx import Presentation
    from pptx.constants import MSO
    from pptx.util import Inches

    prs = Presentation()
    title_only_slidelayout = prs.slidelayouts[5]
    slide = prs.slides.add_slide(title_only_slidelayout)
    shapes = slide.shapes

    shapes.title.text = 'Adding an AutoShape'

    left = Inches(0.93)  # 0.93" centers this overall set of shapes
    top = Inches(3.0)
    width = Inches(1.75)
    height = Inches(1.0)

    shape = shapes.add_shape(MSO.SHAPE_PENTAGON, left, top, width, height)
    shape.text = 'Step 1'

    left = left + width - Inches(0.4)
    width = Inches(2.0)  # chevrons need more width for visual balance

    for n in range(2, 6):
        shape = shapes.add_shape(MSO.SHAPE_CHEVRON, left, top, width, height)
        shape.text = 'Step %d' % n
        left = left + width - Inches(0.4)

    prs.save('test.pptx')


Constants representing each of the available auto shapes (like
MSO.SHAPE_ROUNDED_RECT, MSO.SHAPE_CHEVRON, etc.) are listed on the
:doc:`autoshape-types` page.


----


``add_table()`` example
-----------------------

.. image:: /_static/img/add-table.png

|

::

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    title_only_slidelayout = prs.slidelayouts[5]
    slide = prs.slides.add_slide(title_only_slidelayout)
    shapes = slide.shapes

    shapes.title.text = 'Adding a Table'

    rows = cols = 2
    left = top = Inches(2.0)
    width = Inches(6.0)
    height = Inches(0.8)

    tbl = shapes.add_table(rows, cols, left, top, width, height)

    # set column widths
    tbl.columns[0].width = Inches(2.0)
    tbl.columns[1].width = Inches(4.0)

    # write column headings
    tbl.cell(0, 0).text = 'Foo'
    tbl.cell(0, 1).text = 'Bar'

    # write body cells
    tbl.cell(1, 0).text = 'Baz'
    tbl.cell(1, 1).text = 'Qux'

    prs.save('test.pptx')


----

Extract all text from slides in presentation
--------------------------------------------

::

    from pptx import Presentation

    prs = Presentation(path_to_presentation)

    # text_runs will be populated with a list of strings,
    # one for each text run in presentation
    text_runs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_textframe:
                continue
            for paragraph in shape.textframe.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text)

