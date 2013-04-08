######################
Manipulate a new chart
######################

Topic of inquiry
================

What actions from the UI are required to produce a fully specified chart from
the generic placeholder chart produced by the ``Insert Chart`` menu command?
Assuming for now that the API calls to produce a fully specified chart will
roughly mirror this procedure, knowing the UI actions should inform the API
requirements.


Abstract
========

SCRAP: A presentation containing only a single chart was created using the default
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

SCRAP:

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

...
