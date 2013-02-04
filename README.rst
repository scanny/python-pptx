###########
python-pptx
###########

VERSION: 0.1.0a1 (first Alpha)


STATUS (as of Feb 1 2013)
=========================

Initial alpha version with limited capabilities. Under active development, with new features added in a new release roughly every two weeks.


Documentation
=============

Documentation is hosted on Read The Docs (readthedocs.org) at
https://python-pptx.readthedocs.org/en/latest/. The documentation is being
developed steadily alongside the code.


Installation
============

``python-pptx`` depends on the ``lxml`` package and the Python Imaging Library
(``PIL``).

``python-pptx`` may be installed with ``pip`` if you have it
available::

    pip install python-pptx

It can also be installed using ``easy_install``::

    easy_install python-pptx

If neither ``pip`` nor ``easy_install`` is available, it can be installed
manually by downloading the distribution from PyPI, unpacking the tarball,
and running ``setup.py``::

    tar xvzf python-pptx-0.1.0a1.tar.gz
    cd python-pptx-0.1.0a1
    python setup.py install


Vision
======

A robust, full-featured, and well-documented general-purpose library for
manipulating Open XML PowerPoint files.

* **robust** - High reliability driven by a full unit-test suite.

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

* **manipulate** - Initially I expect this library to be primarily for
  purposes of *writing* .pptx files. But since we're talking about vision
  here, I think it's not to much to envision that it could be developed to
  also be able to *read* .pptx files and manipulate their contents. I could
  see that coming in handy for full-text indexing, removing speaker notes,
  changing out templates, that sort of thing.


License
=======

Licensed under the `MIT license`_. Short version: this code is copyrighted by
me (Steve Canny), I give you permission to do what you want with it except
remove my name from the credits. See the LICENSE file for specific terms.

.. _MIT license:
   http://www.opensource.org/licenses/mit-license.php
