
python-pptx
===========

Release v\ |version| (:ref:`Installation <install>`)

.. include:: ../README.rst


Feature Support
---------------

|pp| has the following capabilities, with many more on the roadmap:

* Round-trip any Open XML presentation (.pptx file) including all its elements
* Add slides
* Populate text placeholders, for example to create a bullet slide
* Add image to slide at arbitrary position and size
* Add textbox to a slide; manipulate text font size and bold
* Add table to a slide
* Add auto shapes (e.g. polygons, flowchart shapes, etc.) to a slide
* Add and manipulate column, bar, line, and pie charts
* Access and change core document properties such as title and subject

Additional capabilities are actively being developed and added on a release
cadence of roughly once per month. If you find a feature you need that |pp|
doesn't yet have, reach out via the mailing list or issue tracker and we'll see
if we can jump the queue for you to pop it in there :)


User Guide
----------

.. toctree::
   :maxdepth: 1

   user/intro
   user/install
   user/quickstart
   user/presentations
   user/slides
   user/understanding-shapes
   user/autoshapes
   user/placeholders-understanding
   user/placeholders-using
   user/text
   user/charts
   user/table
   user/notes
   user/use-cases
   user/concepts


Community Guide
---------------

.. toctree::
   :maxdepth: 1

   community/faq
   community/support
   community/updates


.. _api:

API Documentation
-----------------

.. toctree::
   :maxdepth: 2

   api/presentation
   api/slides
   api/shapes
   api/placeholders
   api/table
   api/chart-data
   api/chart
   api/text
   api/action
   api/dml
   api/image
   api/exc
   api/util
   api/enum/index


Contributor Guide
-----------------

.. toctree::
   :maxdepth: 1

   dev/runtests
   dev/xmlchemy
   dev/development_practices
   dev/philosophy
   dev/analysis/index
   dev/resources/index
