.. python-pptx documentation master file, created by
   sphinx-quickstart on Thu Nov 29 13:59:35 2012.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to python-pptx's documentation!
=======================================

:mod:`python-pptx` is a pure-Python library for reading and writing PowerPoint
(``.pptx``) files.

The first alpha was released on Jan 26 2013 and the library is under active
development.

Here's what a Hello, World! example looks like:

.. image:: /_static/img/hello-world.png

|

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

The documentation is under development alongside the code, so what you'll find
here is a bit spotty and likely incorrect in some places. Until a few more
features are filled out, the code should be the final reference.

----

Contents:

.. toctree::
   :maxdepth: 2

   user/index
   developer/index



Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

