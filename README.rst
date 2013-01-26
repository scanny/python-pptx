###########
python-pptx
###########

STATUS (as of Jan 25 2013)
==========================

python-pptx-0.1.0a1

Completed distribution items, primarily setup.py. Uploading first alpha to
PyPI.


STATUS (as of Jan 19 2013)
==========================

Development has reached a point where the library is beginning to become
useful for some practical applications and its time to work toward getting an
alpha deployment out there.

The following items have been added over the past couple weeks:

* Presentation.add_slide()
* the ability to change the text of text placeholders and a basic, low-level
  text manipulation interface
* a basic implementation of add_picture()
* acceptance tests based on *behave*

Next is:

* add_textbox()
* deployment of alpha to PyPI
* deployment of end-user documentation to RTD.com


STATUS (as of Jan 3 2013)
=========================

Starting to get interesting now. The ``pptx.packaging`` module is quite stable
and I've been focusing attention on the ``pptx.presentation`` module for the
last week or so. The ``packaging`` module takes care of getting things into
and out of the .pptx package. The ``presentation`` module is what you interact
with directly when using the library, Presentation.open(), prs.add_slide(),
that sort of thing.

A .pptx file will round-trip from a package into memory and back and open up
in PowerPoint fine. The object model on the in-memory side has the objects
Presentation, SlideMaster, SlideLayout, and Slide at the part level, and
Shape, Placeholder (title at least), TextFrame, Paragraph, and Run at the
element level. So the library now actually works to modify an existing
presentation, at least to change the text in placeholder shapes.

Right now I'm working on SlideCollection.add_slide(), which will allow adding
new slides, that will be a big milestone. After that I'll be working on
Shapes.add_x() for x = things like image, textbox, smart-shape probably, table
before too long I'm sure.

Anyway, was time for another push. Seems like I get to those more like once
every two weeks, so that's a reasonable expectation of the tempo going forward
for now.

Unit test coverage is up around 96% and I'm using Test-driven development, so
I expect coverage to stay close to 100%. The suite is up to 146 tests.

There's also a start on documentation, although no user documentation so far,
will have to wait for the top-level API to get a little further along before I
attend to that.


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
