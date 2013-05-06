#####
Chart
#####

:Updated:  2013-04-07
:Author:   Steve Canny
:Status:   **WORKING DRAFT**


Introduction
============

One of the shapes available for placing on a PowerPoint slide is the *chart*.
The chart shape is perhaps the most complex of all the PresentationML shapes.
Contributing to its complexity is the fact that its source data is contained in
an embedded Excel spreadsheet.


Research protocol
=================

* ( ) Review MS API documentation
* ( ) Inspect minimal XML produced by PowerPoint® client
* ( ) Review and document relevant schema elements


General Analysis
================

* Need the minimal reliable test of what shape type is contained in
  a CT_GraphicalObjectFrame element. I'm thinking it's determined by the first
  child element that appears in the ``<a:graphic>`` element.
* I suspect that ``SmartArt`` is a third object type that can be contained in
  the ``CT_GraphicalObjectFrame`` element, in addition to ``Table`` and
  ``Chart``.


MS API Analysis
===============

MS API method to add a chart is::

    Shapes.AddChart(Type, Left, Top, Width, Height)


Properties inherited from ``Shape``
-----------------------------------

* There is a ``HasChart`` property on Shape to indicate the shape "has"
  a chart. I'm thinking ``isinstance(shape, Chart)`` will serve that purpose
  for us. The MS ``Shape`` object also has the properties ``Chart`` and
  ``Table``, such that accessing the chart would look like
  ``chart = sld.shapes[9].chart`` rather than ``chart = sld.shapes[9]``. I'm
  not sure why they designed it that way. Best I can think of is they wanted
  to keep the ``Shape`` API separate from the ``Chart`` and ``Table`` API.
  Strikes me as a question to keep in mind as the design continues to emerge.
  I'd hate to discover late there was a better reason than I'm seeing yet :).
* ``top``, ``left``, ``width``, and ``height`` will be inherited from
  ``Shape``, or perhaps from a ``GraphicFrame`` subclass, since those elements
  are shared by both ``Table`` and ``Chart``.
* There is such a thing as a shape placeholder, and having that capability
  might be useful because it would allow certain aspects of the chart to be
  defined in a .pptx template file, such that an end-user could change the size
  and position of the generated charts by uploading a different template file,
  no code changes necessary. Probably not super simple to execute, but not more
  than a medium-sized feature.
* ``name`` should come for free, inherited from BaseShape. Doesn't show on the
  PowerPoint UI, but can be handy for debugging and perhaps other purposes.


Core ``Chart`` members and properties
-------------------------------------

From the `Chart Members`_ page on MSDN.

* ``Axes()`` -- implies an ``Axis`` class. I suppose there can be as many as
  three, not sure.
* ``SeriesCollection()`` -- implies a ``Series`` class. That needs a double-click
  down as these must be a core element.
* ``SetElement()`` -- no clue yet, but looks important.
* ``SetSourceData()`` -- set the source data range for the chart. I expect this
  specifies the Excel sheet range containing the chart data.
* ``ChartArea`` -- dunno what this is yet
* ``ChartData`` -- ...
* ``ChartStyle`` -- ...
* ``ChartTitle`` -- ...
* ``ChartType`` -- ...
* ``DataTable`` -- ...
* ``Format`` -- returns the ``ChartFormat`` object ...
* ``HasDataTable`` -- the DataTable object is worth exploring, could be there's
  an alternative to embedding an Excel spreadsheet.
* ``Legend`` -- returns the ``Legend`` object for the chart
* ``PlotArea`` -- not sure what this is
* ``Shapes`` -- collection of all shapes on the chart. needs looking into
* ``Title`` -- read/write, title of the chart


Not core, but worth exploring for analysis purposes
---------------------------------------------------

* There are ``ApplyChartTemplate()`` and ``SaveChartTemplate()`` methods, which
  implies the existence of something called a chart template. Not sure what
  those are, these two go on the 'investigate further' list.
* ``ChartGroup()`` -- not sure what this is, hoping it's an advanced feature we
  don't need though :)
* ``Delete()`` -- probably can get away without this initially, since there's no
  need to delete something if you're creating charts from scratch. However, if
  it's pretty easy, probably good to get it in there while the head's wrapped
  around the code. Would be required to manipulate an existing slide. Might
  even be handy if templates are used.


XML produced by PowerPoint® application
=======================================

Inspection Notes
----------------

...


XML produced by PowerPoint® client
----------------------------------

.. highlight:: xml

::

    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <c:date1904 val="0"/>
      <c:lang val="en-US"/>
      <c:roundedCorners val="0"/>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
          <c14:style val="118"/>
        </mc:Choice>
        <mc:Fallback>
          <c:style val="18"/>
        </mc:Fallback>
      </mc:AlternateContent>
      <c:chart>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
          <c:layout/>
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <c:ser>
              <c:idx val="0"/>
              <c:order val="0"/>
              <c:tx>
                <c:strRef>
                  <c:f>Sheet1!$B$1</c:f>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0">
                      <c:v>Series 1</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:invertIfNegative val="0"/>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$5</c:f>
                  <c:strCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>Category 1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>Category 2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>Category 3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>Category 4</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$B$2:$B$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>4.3</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>2.5</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>3.5</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>4.5</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:ser>
              <c:idx val="1"/>
              <c:order val="1"/>
              <c:tx>
                <c:strRef>
                  <c:f>Sheet1!$C$1</c:f>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0">
                      <c:v>Series 2</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:invertIfNegative val="0"/>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$5</c:f>
                  <c:strCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>Category 1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>Category 2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>Category 3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>Category 4</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$C$2:$C$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>2.4</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>4.4</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>1.8</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>2.8</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:ser>
              <c:idx val="2"/>
              <c:order val="2"/>
              <c:tx>
                <c:strRef>
                  <c:f>Sheet1!$D$1</c:f>
                  <c:strCache>
                    <c:ptCount val="1"/>
                    <c:pt idx="0">
                      <c:v>Series 3</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:tx>
              <c:invertIfNegative val="0"/>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$5</c:f>
                  <c:strCache>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>Category 1</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>Category 2</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>Category 3</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>Category 4</c:v>
                    </c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$D$2:$D$5</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="4"/>
                    <c:pt idx="0">
                      <c:v>2.0</c:v>
                    </c:pt>
                    <c:pt idx="1">
                      <c:v>2.0</c:v>
                    </c:pt>
                    <c:pt idx="2">
                      <c:v>3.0</c:v>
                    </c:pt>
                    <c:pt idx="3">
                      <c:v>5.0</c:v>
                    </c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
            <c:dLbls>
              <c:showLegendKey val="0"/>
              <c:showVal val="0"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
            </c:dLbls>
            <c:gapWidth val="150"/>
            <c:axId val="2051737496"/>
            <c:axId val="2051748984"/>
          </c:barChart>
          <c:catAx>
            <c:axId val="2051737496"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="b"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="2051748984"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblAlgn val="ctr"/>
            <c:lblOffset val="100"/>
            <c:noMultiLvlLbl val="0"/>
          </c:catAx>
          <c:valAx>
            <c:axId val="2051748984"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="l"/>
            <c:majorGridlines/>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="2051737496"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="between"/>
          </c:valAx>
        </c:plotArea>
        <c:legend>
          <c:legendPos val="r"/>
          <c:layout/>
          <c:overlay val="0"/>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:showDLblsOverMax val="0"/>
      </c:chart>
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="1800"/>
          </a:pPr>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </c:txPr>
      <c:externalData r:id="rId1">
        <c:autoUpdate val="0"/>
      </c:externalData>
    </c:chartSpace>

Resources
=========

.. _Chart Members:
   http://msdn.microsoft.com/en-us/library/office/ff746468(v=office.14).aspx

