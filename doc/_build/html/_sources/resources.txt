=========
Resources
=========

... just collected bits and pieces here for now ...


ECMA Spec
=========
...


Microsoft PowerPoint API documentation
======================================

* `Microsoft PowerPoint 2013 Object Model Reference <http://msdn.microsoft.com/en-us/library/office/ff743835.aspx>`_

* `MSDN - Presentation Members <http://msdn.microsoft.com/en-us/library/office/ff745984(v=office.14).aspx>`_

* `Different MSDN PowerPoint API reference <http://msdn.microsoft.com/en-us/library/documentformat.openxml.presentation.presentation_members.aspx>`_


Documentation Guides
====================

* DOC: Documentation guide
  `Documenting Python Guide <http://docs.python.org/devguide/documenting.html>`_

* DOC: `Google Python Style Guide <http://google-styleguide.googlecode.com/svn/trunk/pyguide.html>`_

* `Read The Docs (readthedocs.org) <https://docs.readthedocs.org/en/latest/index.html>`_

* `OpenComparison documetation on readthedocs.org <http://opencomparison.readthedocs.org/en/latest/contributing.html>`_


Other Resources
===============

* `Python Magic Methods <http://www.rafekettler.com/magicmethods.html>`_

* `lxml.etree Tutorial <http://lxml.de/tutorial.html>`_

* `lxml API Reference <http://lxml.de/api/index.html>`_

* `The factory pattern in Python with __new__ <http://whilefalse.net/2009/10/21/factory-pattern-python-__new__/>`_


Class Hierarchy
===============

::

   pptx
   |
   +--.packaging
      |
      +--PartCollection
      |  |
      |  +--ImageParts
      |  +--SlideLayoutParts
      |  +--SlideMasterParts
      |  +--SlideParts
      |  +--ThemeParts
      |
      +--Part
         |
         +--CollectionPart
         |  |
         |  +--ImagePart
         |  +--SlideLayoutPart
         |  +--SlideMasterPart
         |  +--ThemePart
         |
         +--PresentationPart
         +--PresPropsPart
         +--SlidePart
         +--TableStylesPart
         +--ViewPropsPart


Also try something like this::

   Exception hierarchy
   -------------------

   The class hierarchy for built-in exceptions is:

   .. literalinclude:: ../../Lib/test/exception_hierarchy.txt


