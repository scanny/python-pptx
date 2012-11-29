###########
python-pptx
###########

STATUS (as of Nov 29 2012)
==========================

Under active development, but just starting out. Expect new pushes once a week
or so for the next several weeks. Hoping to get baseline .pptx writing
capability accomplished by the end of the calendar year. That would be:

* creating new presentations from default or arbitrary template (.potx)
* adding new slides
* adding placeholder text (title, bullets, etc.)
* adding shapes and images to slides

This initial push doesn't do much of a practical nature, I've been focused on
getting the scaffolding code sound so I could build out some real
functionality without getting lost in my own code :)

One or two folks wanted to see what I had so far, so here it is. Expect to see
some actual useful functionality in the next month or so, I've got some time
off coming up I'll be using to add some useful end-user features.


Vision
======

A robust, full-featured, general-purpose library for manipulating Open XML
PowerPoint files.

* *robust* - High reliability driven by a full test suite.

* *full-featured* - Anything that the file format will allow can be
  accomplished via the API. (Note that visions often take some time to fulfill
  completely :).

* *general-purpose* - Applicability to all conceivable purposes is valued over
  being especially well-suited to any particular purpose. Particular purposes
  can always be accomplished by building a wrapper library of your own.
  Serving general purposes from a particularized library is not so easy.

* *manipulate* - Initially I expect this library to be primarily for purposes
  of *writing* .pptx files. But since we're talking about vision here, I think
  it's not to much to envision that it could be developed to also be able to
  *read* .pptx files and manipulate their contents. I could see that coming
  in handy for full-text indexing, removing speaker notes, changing out
  templates, that sort of thing.


License
=======

Licensed under the `MIT license <http://www.opensource.org/licenses/mit-license.php>`_.
Short version: this code is copyrighted by me (Steve Canny), I give you
permission to do what you want with it except remove my name from the credits.
See the LICENSE file for specific terms.