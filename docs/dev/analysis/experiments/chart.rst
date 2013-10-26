##################
Add chart to slide
##################

Topic of inquiry
================

What does PowerPoint add to the package when a basic chart is added to a blank
presentation?


Abstract
========

A presentation containing only a single chart was created using the default
template and the blank slide layout. The resulting package was unzipped and its
contents inspected for differences from an empty presentation based on the same
template. Inspection revealed two new package parts and two additional
relationship elements. The new package parts were ``/ppt/charts/chart1.xml``
and ``/ppt/embeddings/Microsoft_Excel_Sheet1.xlsx``. One new relationship was
from the slide to the chart part it contained. The other was from the chart to
the embedded spreadsheet package. The chart shape was contained in
a ``<p:graphicFrame>`` element in the same way a table shape is contained in
that element, except that the chart element is a stub reference to the separate
chart part file.


Procedure
=========

This procedure was run using PowerPoint for Mac 2011 running on an Intel-based
Mac Pro running Mountain Lion 10.8.3.

1. Create a new presentation based on the *White* theme.

#. Change the layout of the default first slide to the *Blank* layout.

#. From the *Charts* ribbon, insert a *Clustered Column* chart (first
   option in list).

#. Close the Excel spreadsheet that opens containing the default dataset.

#. Save the presentation as ``chart.pptx``.

#. Unpack the presentation. This command works on OS X: ``unzip chart.pptx -d
   chart``. Use ``xmllint`` as required to reformat the XML files, e.g.:
   ``xmllint --format slide1.xml >indented-slide1.xml``.

#. Inspect the package contents for new directories, parts, and relationships.
   Inspect ``slide1.xml`` and ``slides/_rels/slide1.xml.rels`` for contents
   related to the chart.


Observations
============

The following package items not present in the base presentation appear::

    ppt
    |-- charts
    |   |-- _rels
    |   |   `-- chart1.xml.rels
    |   `-- chart1.xml
    `-- embeddings
        `-- Microsoft_Excel_Sheet1.xlsx

The ``/ppt`` directory contains ``charts`` and ``embeddings`` directories,
which do not appear in a basic presentation.

The ``charts`` directory contains ``/ppt/charts/chart1.xml``, a distinct part
containing the chart definition XML. The chart part has a ``.rels`` file
``/ppt/charts/_rels/chart1.xml.rels`` indicating the chart has outbound
relationships.

The ``embeddings`` directory contains ``Microsoft_Excel_Sheet1.xlsx``,
a full-fledged embedded Excel spreadsheet containing the dataset for the chart.
It can be opened with Excel as-is once the package is unzipped.

More specific differences are described in the following sections.


``[Content_Types].xml``
-----------------------

``'xlsx'`` is present in ``[Content_Types].xml`` as a default extension content
type, mapping to the SpreadsheetML content type. Everything else in
``[Content_Types].xml`` appears standard.


``slide1.xml``
--------------

The chart shape is contained in the slide's shape tree as
a ``<p:graphicFrame>`` element, in a structure parallel to that of a table
shape. The chart definition itself appears simply as a relationship link
pointing to the chart part.


``slide1.xml.rels``
-------------------

In addition to the usual relationship to its slide layout, the slide has an
additional relationship pointing to the chart part.


``chart1.xml``
--------------

The chart part XML elements belong to the
``http://schemas.openxmlformats.org/drawingml/2006/chart`` namespace with the
root element ``<c:chartSpace>``. The default chart resolves to 263 lines of
`XML`. Main elements are chart, plotArea, catAx (category axis), valAx (value
axis), barChart, ser (series), cat (category), and val (value). Deepest nesting
appears to be 9 levels.


``chart1.xml.rels``
-------------------

``chart1.xml.rels`` contains a single relationship, to the embedded Excel
package.
