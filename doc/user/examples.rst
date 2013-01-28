========
Examples
========

Examples to help with getting started ...

Hello World! example
====================

::

    from pptx import Package
    
    pkg = Package()
    prs = pkg.presentation
    title_slidelayout = prs.slidemasters[0].slidelayouts[0]
    slide = prs.slides.add_slide(title_slidelayout)
    title = slide.shapes.title
    subtitle = slide.shapes.placeholders[1]
    
    title.text = "Hello, World!"
    subtitle.text = "python-pptx was here!"
    
    pkg.save('test.pptx')

**image goes here**


