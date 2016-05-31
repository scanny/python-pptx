
Charts - Overview
=================

* category and series data is arranged in columns
  
  + rows = data-points/series+1

What does the MS API look like?

Axis Members
http://msdn.microsoft.com/en-us/library/office/ff745187.aspx

* loading one in to manipulate it is outside current scope. Current scope is
  to add one to a new .pptx file

Notes
-----

<c:layoutTarget> is not valid within c:chart/c:legend/c:layout. Probably only
for laying out the chart proper or whatever.


Object graph
------------

* Chart

  + CategoryAxis(BaseAxis)
  + ValueAxis(BaseAxis)


Chart parts glossary
--------------------

**data point (point)**
   An individual numeric value, represented by a bar, point, column, or pie
   slice.

**data series (series)**
   A group of related data points. For example, the columns of a series will
   all be the same color.

**category axis (X axis)**
   The horizontal axis of a two-dimensional or three-dimensional chart.

**value axis (Y axis)**
   The vertical axis of a two-dimensional or three-dimensional chart.

**depth axis (Z axis)**
   The front-to-back axis of a three-dimensional chart.

**grid lines**
   Horizontal or vertical lines that may be added to an axis to aid
   comparison of a data point to an axis value.

**legend**
   A key that explains which data series each color or pattern represents.

**floor**
   The bottom of a three-dimensional chart.

**walls**
   The background of a chart. Three-dimensional charts have a back wall and
   a side wall, which can be formatted separately.

**data labels**
   Numeric labels on each data point. A data label can represent the actual
   value or a percentage.

**axis title**
   Explanatory text label associated with an axis

**data table**
   A optional tabular display within the *plot area* of the values on which
   the chart is based. Not to be confused with the Excel worksheet holding
   the chart values.

**chart title**
   A label explaining the overall purpose of the chart.

**chart area**
   Overall chart object, containing the chart and all its auxiliary pieces
   such as legends and titles.

**plot area**
   Region of the chart area that contains the actual plots, bounded by but
   not including the axes. May contain more than one plot, each with its own
   distinct set of series. A plot is known as a *chart group* in the MS API.

**axis**
   ... may be either a *category axis* or a *value axis* ... on
   a two-dimensional chart, either the horizontal (*x*) axis or the vertical
   (*y*) axis. A 3-dimensional chart also has a depth (*z*) axis. Pie,
   doughnut, and radar charts have a radial axis.

   How many axes do each of the different chart types have?

**series categories**
   ...

**series values**
   ...


Chart types
-----------

* column

  - 2-D column

    + clustered column
    + stacked column
    + 100% stacked column

  - 3-D column

    + 3-D clustered column
    + 3-D stacked column
    + 3-D 100% stacked column
    + 3-D column

  - cylinder
  - cone
  - pyramid

* line

  + 2-D line
  + 3-D line

* pie

  + 2-D pie
  + 3-D pie

* bar

  + 2-D bar
  + 3-D bar
  + cylinder
  + cone
  + pyramid

* area

* scatter

* other

  + stock (e.g. open-high-low-close)
  + surface
  + doughnut
  + bubble
  + radar


XML specimens
-------------

.. highlight:: xml

Containing ``<p:graphicFrame>`` in shape tree::

    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="7" name="Chart Placeholder 4"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr>
          <p:ph sz="quarter" idx="23"/>
          <p:custDataLst>
            <p:tags r:id="rId1"/>
          </p:custDataLst>
          <p:extLst>
            <p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}">
              <p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                         val="4163776498"/>
            </p:ext>
          </p:extLst>
        </p:nvPr>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="142874" y="1502410"/>
        <a:ext cx="8838000" cy="4320000"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="rId5"/>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>


Pie Chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
        <c14:style val="102"/>
      </mc:Choice>
      <mc:Fallback>
        <c:style val="2"/>
      </mc:Fallback>
    </mc:AlternateContent>
    <c:chart>
      <c:autoTitleDeleted val="1"/>
      <c:plotArea>
        <c:layout/>
        <c:pieChart>
          <c:varyColors val="1"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:f>Sheet1!$A$2</c:f>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>Base</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:dLbls>
              <c:numFmt formatCode="0%" sourceLinked="0"/>
              <c:spPr>
                <a:noFill/>
                <a:ln>
                  <a:noFill/>
                </a:ln>
                <a:effectLst/>
              </c:spPr>
              <c:txPr>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="1000"/>
                  </a:pPr>
                  <a:endParaRPr lang="en-US"/>
                </a:p>
              </c:txPr>
              <c:dLblPos val="outEnd"/>
              <c:showLegendKey val="0"/>
              <c:showVal val="1"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
              <c:showLeaderLines val="1"/>
              <c:extLst xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                        xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
                        xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">
                <c:ext xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"
                       uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"/>
              </c:extLst>
            </c:dLbls>
            <c:cat>
              <c:strRef>
                <c:f>Sheet1!$B$1:$F$1</c:f>
                <c:strCache>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>Très probable</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>Plutôt probable</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>Plutôt improbable</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>Très improbable</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>Je ne sais pas</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$B$2:$F$2</c:f>
                <c:numCache>
                  <c:formatCode>0.00%</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>0.1348</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>0.3238</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>0.1803</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>0.2349</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>0.1262</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:ser>
          <c:dLbls>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/>
            <c:showVal val="1"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
            <c:showLeaderLines val="1"/>
          </c:dLbls>
          <c:firstSliceAng val="0"/>
        </c:pieChart>
      </c:plotArea>
      <c:legend>
        <c:legendPos val="b"/>
        <c:layout/>
        <c:overlay val="0"/>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="1000"/>
            </a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </c:txPr>
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


single series line chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
        <c14:style val="102"/>
      </mc:Choice>
      <mc:Fallback>
        <c:style val="2"/>
      </mc:Fallback>
    </mc:AlternateContent>
    <c:chart>
      <c:autoTitleDeleted val="1"/>
      <c:plotArea>
        <c:layout/>
        <c:lineChart>
          <c:grouping val="standard"/>
          <c:varyColors val="0"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:f>Sheet1!$A$2</c:f>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>Base</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:marker>
              <c:symbol val="none"/>
            </c:marker>
            <c:dLbls>
              <c:numFmt formatCode="General" sourceLinked="0"/>
              <c:spPr>
                <a:noFill/>
                <a:ln>
                  <a:noFill/>
                </a:ln>
                <a:effectLst/>
              </c:spPr>
              <c:txPr>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="1000"/>
                  </a:pPr>
                  <a:endParaRPr lang="en-US"/>
                </a:p>
              </c:txPr>
              <c:showLegendKey val="0"/>
              <c:showVal val="1"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
              <c:showLeaderLines val="0"/>
              <c:extLst xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                        xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
                        xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">
                <c:ext xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"
                       uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}">
                  <c15:layout/>
                  <c15:showLeaderLines val="0"/>
                </c:ext>
              </c:extLst>
            </c:dLbls>
            <c:cat>
              <c:strRef>
                <c:f>Sheet1!$B$1:$F$1</c:f>
                <c:strCache>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>Très probable</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>Plutôt probable</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>Plutôt improbable</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>Très improbable</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>Je ne sais pas</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$B$2:$F$2</c:f>
                <c:numCache>
                  <c:formatCode>0</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>19.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>13.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>10.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>46.0</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>12.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:val>
            <c:smooth val="0"/>
          </c:ser>
          <c:dLbls>
            <c:showLegendKey val="0"/>
            <c:showVal val="1"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:marker val="1"/>
          <c:smooth val="0"/>
          <c:axId val="-2097691448"/>
          <c:axId val="-2097683336"/>
        </c:lineChart>
        <c:catAx>
          <c:axId val="-2097691448"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:numFmt formatCode="General" sourceLinked="0"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1000"/>
              </a:pPr>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </c:txPr>
          <c:crossAx val="-2097683336"/>
          <c:crosses val="autoZero"/>
          <c:auto val="1"/>
          <c:lblAlgn val="ctr"/>
          <c:lblOffset val="100"/>
          <c:noMultiLvlLbl val="0"/>
        </c:catAx>
        <c:valAx>
          <c:axId val="-2097683336"/>
          <c:scaling>
            <c:orientation val="minMax"/>
            <c:max val="100.0"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          <c:majorGridlines/>
          <c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1000" b="1"/>
              </a:pPr>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </c:txPr>
          <c:crossAx val="-2097691448"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="between"/>
        </c:valAx>
      </c:plotArea>
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


simple column chart::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <c:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
    <c:date1904 val="0"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
                 Requires="c14">
        <c14:style val="102"/>
      </mc:Choice>
      <mc:Fallback>
        <c:style val="2"/>
      </mc:Fallback>
    </mc:AlternateContent>
    <c:chart>
      <c:autoTitleDeleted val="1"/>
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
                <c:f>Sheet1!$A$2</c:f>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>Base</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:invertIfNegative val="0"/>
            <c:dLbls>
              <c:numFmt formatCode="General" sourceLinked="0"/>
              <c:spPr>
                <a:noFill/>
                <a:ln>
                  <a:noFill/>
                </a:ln>
                <a:effectLst/>
              </c:spPr>
              <c:txPr>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr sz="1000"/>
                  </a:pPr>
                  <a:endParaRPr lang="en-US"/>
                </a:p>
              </c:txPr>
              <c:dLblPos val="outEnd"/>
              <c:showLegendKey val="0"/>
              <c:showVal val="1"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
              <c:showLeaderLines val="0"/>
              <c:extLst xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"
                        xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                <c:ext xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"
                       uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}">
                  <c15:layout/>
                  <c15:showLeaderLines val="0"/>
                </c:ext>
              </c:extLst>
            </c:dLbls>
            <c:cat>
              <c:strRef>
                <c:f>Sheet1!$B$1:$F$1</c:f>
                <c:strCache>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>Très probable</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>Plutôt probable</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>Plutôt improbable</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>Très improbable</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>Je ne sais pas</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:f>Sheet1!$B$2:$F$2</c:f>
                <c:numCache>
                  <c:formatCode>0</c:formatCode>
                  <c:ptCount val="5"/>
                  <c:pt idx="0">
                    <c:v>19.0</c:v>
                  </c:pt>
                  <c:pt idx="1">
                    <c:v>13.0</c:v>
                  </c:pt>
                  <c:pt idx="2">
                    <c:v>10.0</c:v>
                  </c:pt>
                  <c:pt idx="3">
                    <c:v>46.0</c:v>
                  </c:pt>
                  <c:pt idx="4">
                    <c:v>12.0</c:v>
                  </c:pt>
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:ser>
          <c:dLbls>
            <c:dLblPos val="outEnd"/>
            <c:showLegendKey val="0"/>
            <c:showVal val="1"/>
            <c:showCatName val="0"/>
            <c:showSerName val="0"/>
            <c:showPercent val="0"/>
            <c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:gapWidth val="150"/>
          <c:axId val="-2053894120"/>
          <c:axId val="-2053699928"/>
        </c:barChart>
        <c:catAx>
          <c:axId val="-2053894120"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:numFmt formatCode="General" sourceLinked="0"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1000"/>
              </a:pPr>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </c:txPr>
          <c:crossAx val="-2053699928"/>
          <c:crosses val="autoZero"/>
          <c:auto val="1"/>
          <c:lblAlgn val="ctr"/>
          <c:lblOffset val="100"/>
          <c:noMultiLvlLbl val="0"/>
        </c:catAx>
        <c:valAx>
          <c:axId val="-2053699928"/>
          <c:scaling>
            <c:orientation val="minMax"/>
            <c:max val="100.0"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="l"/>
          <c:majorGridlines/>
          <c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1000" b="1"/>
              </a:pPr>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </c:txPr>
          <c:crossAx val="-2053894120"/>
          <c:crosses val="autoZero"/>
          <c:crossBetween val="between"/>
        </c:valAx>
      </c:plotArea>
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


Related Schema Definitions
--------------------------

.. highlight:: xml

::

  <!-- homonym <c:chart> element in graphicData element -->

  <xsd:element name="chart" type="CT_RelId"/>

  <xsd:complexType name="CT_RelId">
    <xsd:attribute ref="r:id" use="required"/>
  </xsd:complexType>


  <!-- elements in chartX.xml part -->

  <xsd:element name="chartSpace" type="CT_ChartSpace"/>

  <xsd:complexType name="CT_ChartSpace">
    <xsd:sequence>
      <xsd:element name="date1904"       type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lang"           type="CT_TextLanguageID"    minOccurs="0"/>
      <xsd:element name="roundedCorners" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="style"          type="CT_Style"             minOccurs="0"/>
      <xsd:element name="clrMapOvr"      type="a:CT_ColorMapping"    minOccurs="0"/>
      <xsd:element name="pivotSource"    type="CT_PivotSource"       minOccurs="0"/>
      <xsd:element name="protection"     type="CT_Protection"        minOccurs="0"/>
      <xsd:element name="chart"          type="CT_Chart"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="externalData"   type="CT_ExternalData"      minOccurs="0"/>
      <xsd:element name="printSettings"  type="CT_PrintSettings"     minOccurs="0"/>
      <xsd:element name="userShapes"     type="CT_RelId"             minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Chart">
    <xsd:sequence>
      <xsd:element name="title"            type="CT_Title"         minOccurs="0"/>
      <xsd:element name="autoTitleDeleted" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="pivotFmts"        type="CT_PivotFmts"     minOccurs="0"/>
      <xsd:element name="view3D"           type="CT_View3D"        minOccurs="0"/>
      <xsd:element name="floor"            type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="sideWall"         type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="backWall"         type="CT_Surface"       minOccurs="0"/>
      <xsd:element name="plotArea"         type="CT_PlotArea"/>
      <xsd:element name="legend"           type="CT_Legend"        minOccurs="0"/>
      <xsd:element name="plotVisOnly"      type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="dispBlanksAs"     type="CT_DispBlanksAs"  minOccurs="0"/>
      <xsd:element name="showDLblsOverMax" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_PlotArea">
    <xsd:sequence>
      <xsd:element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
      <xsd:choice minOccurs="1" maxOccurs="unbounded">
        <xsd:element name="areaChart"      type="CT_AreaChart"/>
        <xsd:element name="area3DChart"    type="CT_Area3DChart"/>
        <xsd:element name="lineChart"      type="CT_LineChart"/>
        <xsd:element name="line3DChart"    type="CT_Line3DChart"/>
        <xsd:element name="stockChart"     type="CT_StockChart"/>
        <xsd:element name="radarChart"     type="CT_RadarChart"/>
        <xsd:element name="scatterChart"   type="CT_ScatterChart"/>
        <xsd:element name="pieChart"       type="CT_PieChart"/>
        <xsd:element name="pie3DChart"     type="CT_Pie3DChart"/>
        <xsd:element name="doughnutChart"  type="CT_DoughnutChart"/>
        <xsd:element name="barChart"       type="CT_BarChart"/>
        <xsd:element name="bar3DChart"     type="CT_Bar3DChart"/>
        <xsd:element name="ofPieChart"     type="CT_OfPieChart"/>
        <xsd:element name="surfaceChart"   type="CT_SurfaceChart"/>
        <xsd:element name="surface3DChart" type="CT_Surface3DChart"/>
        <xsd:element name="bubbleChart"    type="CT_BubbleChart"/>
      </xsd:choice>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element name="valAx"  type="CT_ValAx"/>
        <xsd:element name="catAx"  type="CT_CatAx"/>
        <xsd:element name="dateAx" type="CT_DateAx"/>
        <xsd:element name="serAx"  type="CT_SerAx"/>
      </xsd:choice>
      <xsd:element name="dTable" type="CT_DTable"            minOccurs="0"/>
      <xsd:element name="spPr"   type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="extLst" type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BarChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="barDir"     type="CT_BarDir"/>
      <xsd:element name="grouping"   type="CT_BarGrouping"   minOccurs="0"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"        type="CT_BarSer"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="gapWidth"   type="CT_GapAmount"     minOccurs="0"/>
      <xsd:element name="overlap"    type="CT_Overlap"       minOccurs="0"/>
      <xsd:element name="serLines"   type="CT_ChartLines"    minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_CatAx">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="axId"           type="CT_UnsignedInt"/>
      <xsd:element name="scaling"        type="CT_Scaling"/>
      <xsd:element name="delete"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="axPos"          type="CT_AxPos"/>
      <xsd:element name="majorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="minorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="title"          type="CT_Title"             minOccurs="0"/>
      <xsd:element name="numFmt"         type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="majorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="minorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="tickLblPos"     type="CT_TickLblPos"        minOccurs="0"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="crossAx"        type="CT_UnsignedInt"/>
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"      type="CT_Crosses"/>
        <xsd:element name="crossesAt"    type="CT_Double"/>
      </xsd:choice>
      <xsd:element name="auto"           type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="lblAlgn"        type="CT_LblAlgn"           minOccurs="0"/>
      <xsd:element name="lblOffset"      type="CT_LblOffset"         minOccurs="0"/>
      <xsd:element name="tickLblSkip"    type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="tickMarkSkip"   type="CT_Skip"              minOccurs="0"/>
      <xsd:element name="noMultiLvlLbl"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ValAx">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="axId"           type="CT_UnsignedInt"/>
      <xsd:element name="scaling"        type="CT_Scaling"/>
      <xsd:element name="delete"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="axPos"          type="CT_AxPos"/>
      <xsd:element name="majorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="minorGridlines" type="CT_ChartLines"        minOccurs="0"/>
      <xsd:element name="title"          type="CT_Title"             minOccurs="0"/>
      <xsd:element name="numFmt"         type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="majorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="minorTickMark"  type="CT_TickMark"          minOccurs="0"/>
      <xsd:element name="tickLblPos"     type="CT_TickLblPos"        minOccurs="0"/>
      <xsd:element name="spPr"           type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"           type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="crossAx"        type="CT_UnsignedInt"/>
      <xsd:choice                                                    minOccurs="0">
        <xsd:element name="crosses"   type="CT_Crosses"/>
        <xsd:element name="crossesAt" type="CT_Double"/>
      </xsd:choice>
      <xsd:element name="crossBetween"   type="CT_CrossBetween"      minOccurs="0"/>
      <xsd:element name="majorUnit"      type="CT_AxisUnit"          minOccurs="0"/>
      <xsd:element name="minorUnit"      type="CT_AxisUnit"          minOccurs="0"/>
      <xsd:element name="dispUnits"      type="CT_DispUnits"         minOccurs="0"/>
      <xsd:element name="extLst"         type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Boolean">
    <xsd:attribute name="val" type="xsd:boolean" use="optional" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Double">
    <xsd:attribute name="val" type="xsd:double" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_Scaling">
    <xsd:sequence>
      <xsd:element name="logBase"     type="CT_LogBase"       minOccurs="0"/>
      <xsd:element name="orientation" type="CT_Orientation"   minOccurs="0"/>
      <xsd:element name="max"         type="CT_Double"        minOccurs="0"/>
      <xsd:element name="min"         type="CT_Double"        minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_NumFmt">
    <xsd:attribute name="formatCode"   type="s:ST_Xstring" use="required"/>
    <xsd:attribute name="sourceLinked" type="xsd:boolean"/>
  </xsd:complexType>

  <xsd:complexType name="CT_TickMark">
    <xsd:attribute name="val" type="ST_TickMark" default="cross"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_TickMark">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="cross"/>
      <xsd:enumeration value="in"/>
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="out"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_DLbls">
    <xsd:sequence>
      <xsd:element name="dLbl" type="CT_DLbl" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:choice>
        <xsd:element name="delete" type="CT_Boolean"/>
        <xsd:group   ref="Group_DLbls"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DLbl">
    <xsd:sequence>
      <xsd:element name="idx" type="CT_UnsignedInt"/>
      <xsd:choice>
        <xsd:element name="delete" type="CT_Boolean"/>
        <xsd:group   ref="Group_DLbl"/>
      </xsd:choice>
      <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:group name="Group_DLbls">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="numFmt"          type="CT_NumFmt"            minOccurs="0"/>
      <xsd:element name="spPr"            type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"            type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="dLblPos"         type="CT_DLblPos"           minOccurs="0"/>
      <xsd:element name="showLegendKey"   type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showVal"         type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showCatName"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showSerName"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showPercent"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="showBubbleSize"  type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="separator"       type="xsd:string"           minOccurs="0"/>
      <xsd:element name="showLeaderLines" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="leaderLines"     type="CT_ChartLines"        minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>

  <xsd:complexType name="CT_LineChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="grouping"   type="CT_Grouping"/>
      <xsd:element name="varyColors" type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"        type="CT_LineSer"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"      type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="dropLines"  type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="hiLowLines" type="CT_ChartLines"    minOccurs="0"/>
      <xsd:element name="upDownBars" type="CT_UpDownBars"    minOccurs="0"/>
      <xsd:element name="marker"     type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="smooth"     type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="axId"       type="CT_UnsignedInt"   minOccurs="2" maxOccurs="2"/>
      <xsd:element name="extLst"     type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LineSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"       type="CT_UnsignedInt"/>
      <xsd:element name="order"     type="CT_UnsignedInt"/>
      <xsd:element name="tx"        type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"      type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="marker"    type="CT_Marker"            minOccurs="0"/>
      <xsd:element name="dPt"       type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"     type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline" type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"   type="CT_ErrBars"           minOccurs="0"/>
      <xsd:element name="cat"       type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"       type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="smooth"    type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_BarSer">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="idx"              type="CT_UnsignedInt"/>
      <xsd:element name="order"            type="CT_UnsignedInt"/>
      <xsd:element name="tx"               type="CT_SerTx"             minOccurs="0"/>
      <xsd:element name="spPr"             type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="invertIfNegative" type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="pictureOptions"   type="CT_PictureOptions"    minOccurs="0"/>
      <xsd:element name="dPt"              type="CT_DPt"               minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"            type="CT_DLbls"             minOccurs="0"/>
      <xsd:element name="trendline"        type="CT_Trendline"         minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="errBars"          type="CT_ErrBars"           minOccurs="0"/>
      <xsd:element name="cat"              type="CT_AxDataSource"      minOccurs="0"/>
      <xsd:element name="val"              type="CT_NumDataSource"     minOccurs="0"/>
      <xsd:element name="shape"            type="CT_Shape"             minOccurs="0"/>
      <xsd:element name="extLst"           type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Style">
    <xsd:attribute name="val" type="ST_Style" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Style">
    <xsd:restriction base="xsd:unsignedByte">
      <xsd:minInclusive value="1"/>
      <xsd:maxInclusive value="48"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_PieChart">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="varyColors"    type="CT_Boolean"       minOccurs="0"/>
      <xsd:element name="ser"           type="CT_PieSer"        minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="dLbls"         type="CT_DLbls"         minOccurs="0"/>
      <xsd:element name="firstSliceAng" type="CT_FirstSliceAng" minOccurs="0"/>
      <xsd:element name="extLst"        type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Legend">
    <xsd:sequence>
      <xsd:element name="legendPos"   type="CT_LegendPos"         minOccurs="0"/>
      <xsd:element name="legendEntry" type="CT_LegendEntry"       minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="layout"      type="CT_Layout"            minOccurs="0"/>
      <xsd:element name="overlay"     type="CT_Boolean"           minOccurs="0"/>
      <xsd:element name="spPr"        type="a:CT_ShapeProperties" minOccurs="0"/>
      <xsd:element name="txPr"        type="a:CT_TextBody"        minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"     minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LegendPos">
    <xsd:attribute name="val" type="ST_LegendPos" default="r"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LegendPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="tr"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="t"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_DLblPos">
    <xsd:attribute name="val" type="ST_DLblPos" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_DLblPos">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="bestFit"/>
      <xsd:enumeration value="b"/>
      <xsd:enumeration value="ctr"/>
      <xsd:enumeration value="inBase"/>
      <xsd:enumeration value="inEnd"/>
      <xsd:enumeration value="l"/>
      <xsd:enumeration value="outEnd"/>
      <xsd:enumeration value="r"/>
      <xsd:enumeration value="t"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_LblOffset">
    <xsd:attribute name="val" type="ST_LblOffset" default="100%"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LblOffset">
    <xsd:union memberTypes="ST_LblOffsetPercent ST_LblOffsetUShort"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LblOffsetUShort">
    <xsd:restriction base="xsd:unsignedShort">
      <xsd:minInclusive value="0"/>
      <xsd:maxInclusive value="1000"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_LblOffsetPercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_Overlap">
    <xsd:attribute name="val" type="ST_Overlap" default="0%"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Overlap">
    <xsd:union memberTypes="ST_OverlapPercent ST_OverlapByte"/>
  </xsd:simpleType>

  <xsd:simpleType name="ST_OverlapPercent">
    <xsd:restriction base="xsd:string">
      <xsd:pattern value="(-?0*(([0-9])|([1-9][0-9])|100))%"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:simpleType name="ST_OverlapByte">
    <xsd:restriction base="xsd:byte">
      <xsd:minInclusive value="-100"/>
      <xsd:maxInclusive value="100"/>
    </xsd:restriction>
  </xsd:simpleType>

  <xsd:complexType name="CT_Layout">
    <xsd:sequence>
      <xsd:element name="manualLayout" type="CT_ManualLayout"  minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_ManualLayout">
    <xsd:sequence>
      <xsd:element name="layoutTarget" type="CT_LayoutTarget"  minOccurs="0"/>
      <xsd:element name="xMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="yMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="wMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="hMode"        type="CT_LayoutMode"    minOccurs="0"/>
      <xsd:element name="x"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="y"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="w"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="h"            type="CT_Double"        minOccurs="0"/>
      <xsd:element name="extLst"       type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_LayoutMode">
    <xsd:attribute name="val" type="ST_LayoutMode" default="factor"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_LayoutMode">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="edge"/>
      <xsd:enumeration value="factor"/>
    </xsd:restriction>
  </xsd:simpleType>
