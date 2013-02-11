========
Examples
========

Examples to help with getting started ...

.. highlight:: python
   :linenothreshold: 5

----

Hello World! example
====================

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
====================

.. image:: /_static/img/bullet-slide.png

|

::

    from pptx import Presentation
    
    prs = Presentation()
    bullet_slidelayout = prs.slidelayouts[1]
    
    slide = prs.slides.add_slide(bullet_slidelayout)
    
    shapes = slide.shapes
    title_placeholder = shapes.title
    body_placeholder = shapes.placeholders[1]
    
    title_shape.text = 'Adding a Bullet Slide'
    tf = body_shape.textframe
    tf.text = 'Find the bullet slide layout'
    tf.add_paragraph().text = 'Use Shape.text for first bullet'
    tf.add_paragraph().text = ('Use TextFrame.add_paragraph() for '
                               'subsequent bullets')
    
    prs.save('test.pptx')


----

``add_textbox()`` example
=========================

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
=========================

.. image:: /_static/img/add-picture.png

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


