
python-pptx
===========

Release v\ |version| (:ref:`Installation <install>`)

.. include:: ../README.rst


Philosophy
----------

|pp| aims to broadly support the PowerPoint format (PPTX, PowerPoint 2007 and later),
but its primary commitment is to be *industrial-grade*, that is, suitable for use in a
commercial setting. Maintaining this robustness requires a high engineering standard
which includes a comprehensive two-level (e2e + unit) testing regimen. This discipline
comes at a cost in development effort/time, but we consider reliability to be an
essential requirement.


Feature Support
---------------

|pp| has the following capabilities:

* Round-trip any Open XML presentation (.pptx file) including all its elements
* Add slides
* Populate text placeholders, for example to create a bullet slide
* Add image to slide at arbitrary position and size
* Add textbox to a slide; manipulate text font size and bold
* Add table to a slide
* Add auto shapes (e.g. polygons, flowchart shapes, etc.) to a slide
* Toggle bullet formatting on paragraphs (character, auto-number, none, inherit)
* Add and manipulate column, bar, line, and pie charts
* Discover Office 2016+ extended charts (funnel, treemap, sunburst, waterfall, histogram,
  box-and-whisker, map) on a slide as |GraphicFrame| shapes and preserve them on round-trip
  (detailed read/write not yet supported)
* Access and change core document properties such as title and subject
* Toggle header / footer / slide-number / date placeholder visibility on a slide
  master or layout, and insert auto-refresh slide-number or date fields in a
  text frame
* Read and write legacy PowerPoint review comments on slides
* Organize slides into named sections (``p14:sectionLst``): create, rename,
  delete, and assign slide-to-section membership
* Open and save password-protected (ECMA-376 Agile Encryption) .pptx files
  (requires optional ``msoffcrypto-tool``)
* And many others ...

Even with all |pp| does, the PowerPoint document format is very rich and there are still
features |pp| does not support.


New features/releases
---------------------

New features are generally added via sponsorship. If there's a new feature you need for
your use case, feel free to reach out at the email address on the github.com/scanny
profile page. Many of the most used features such as charts were added this way.


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
   user/media
   user/math-equations
   user/notes
   user/ole-objects
   user/comments
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
   api/animation
   api/comments
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
   dev/security
   dev/analysis/index
   dev/resources/index
