###########
python-pptx
###########

VERSION: 0.2.5 (sixth beta release)


STATUS (as of June 11 2013)
===========================

Stable beta release. Under active development, with new features added in a new
release roughly once a month.


Vision
======

A robust, full-featured, and well-documented general-purpose library for
manipulating Open XML PowerPoint files.

* **robust** - High reliability driven by a comprehensive test suite.

* **full-featured** - Anything that the file format will allow can be
  accomplished via the API. (Note that visions often take some time to fulfill
  completely :).

* **well-documented** - I don't know about you, but I find it hard to remember
  what I was thinking yesterday if I don't write it down. That's not a problem
  for most of my thinking, but when it comes to how I set up an object
  hierarchy to interact, it can be a big time-waster. So I like it when things
  are nicely laid out in black-and-white. Other folks seem to like that too
  :).

* **general-purpose** - Applicability to all conceivable purposes is valued
  over being especially well-suited to any particular purpose. Particular
  purposes can always be accomplished by building a wrapper library of your
  own. Serving general purposes from a particularized library is not so easy.

* **manipulate** - While this library will perhaps most commonly be used for
  *writing* .pptx files, it will also be suitable for *reading* .pptx files
  and inspecting and manipulating their contents. I could see that coming in
  handy for full-text indexing, removing speaker notes, changing out
  templates, adding dynamically generated slides to static boilerplate, that
  sort of thing.


Documentation
=============

Documentation is hosted on Read The Docs (readthedocs.org) at
https://python-pptx.readthedocs.org/en/latest/. The documentation is now in
reasonably robust shape and is being developed steadily alongside the code.


Reaching out
============

We'd love to hear from you if you like |pp|, want a new feature, find a bug,
need help using it, or just have a word of encouragement.

The **mailing list** for |pp| is python.pptx@librelist.com.

The **issue tracker** is on github at `scanny/python-pptx`_.

Feature requests are best broached initially on the mailing list, they can be
added to the issue tracker once we've clarified the best approach,
particularly the appropriate API signature.

.. _`scanny/python-pptx`:
   https://github.com/scanny/python-pptx


Installation
============

|pp| may be installed with ``pip`` if you have it available::

    pip install python-pptx

It can also be installed using ``easy_install``::

    easy_install python-pptx

If neither ``pip`` nor ``easy_install`` is available, it can be installed
manually by downloading the distribution from PyPI, unpacking the tarball,
and running ``setup.py``::

    tar xvzf python-pptx-0.1.0a1.tar.gz
    cd python-pptx-0.1.0a1
    python setup.py install

|pp| depends on the ``lxml`` package and the Python Imaging Library
(``PIL``). Both ``pip`` and ``easy_install`` will take care of satisfying
those dependencies for you, but if you use this last method you will need to
install those yourself.


Release History
===============

June 11, 2013 - v0.2.5
   * Add paragraph alignment property (left, right, centered, etc.)
   * Add vertical alignment within table cell (top, middle, bottom)
   * Add table cell margin properties
   * Add table boolean properties: first column (row header), first row (column
     headings), last row (for e.g. totals row), last column (for e.g. row
     totals), horizontal banding, and vertical banding.
   * Add support for auto shape adjustment values, e.g. change radius of corner
     rounding on rounded rectangle, position of callout arrow, etc.

May 16, 2013 - v0.2.4
   * Add support for auto shapes (e.g. polygons, flowchart symbols, etc.)

May 5, 2013 - v0.2.3
   * Add support for table shapes
   * Add indentation support to textbox shapes, enabling multi-level bullets on
     bullet slides.

Mar 25, 2013 - v0.2.2
   * Add support for opening and saving a presentation from/to a file-like
     object.
   * Refactor XML handling to use lxml objectify

Feb 25, 2013 - v0.2.1
   * Add support for Python 2.6
   * Add images from a stream (e.g. StringIO) in addition to a path, allowing
     images retrieved from a database or network resource to be inserted
     without saving first.
   * Expand text methods to accept unicode and UTF-8 encoded 8-bit strings.
   * Fix potential install bug triggered by importing ``__version__`` from
     package ``__init__.py`` file.

Feb 10, 2013 - v0.2.0
    First non-alpha release with basic capabilities: open presentation/template
    or use built-in default template, add slide, set placeholder text (e.g.
    bullet slides), add picture, add text box.


License
=======

Licensed under the `MIT license`_. Short version: this code is copyrighted by
me (Steve Canny), I give you permission to do what you want with it except
remove my name from the credits. See the LICENSE file for specific terms.

.. _MIT license:
   http://www.opensource.org/licenses/mit-license.php

.. |pp| replace:: ``python-pptx``
