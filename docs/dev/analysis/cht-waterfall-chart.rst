.. _WaterfallChart:


Waterfall Chart
===============

Here's an overview of a [Waterfall Chart](https://en.wikipedia.org/wiki/Waterfall_chart).

A waterfall chart shows a running total as values are added or subtracted. It's useful for
understanding how an initial value (for example, net income) is affected by a series of positive
and negative values.

.. image:: https://support.content.office.net/en-us/media/fe924f0b-bdff-4cb8-bb45-7680e33715d2.png
   :target: https://support.microsoft.com/en-us/office/create-a-waterfall-chart-8de1ece4-ff21-4d37-acd7-546f5527f185

The columns are color coded so you can quickly tell positive from negative numbers. The initial and
the final value columns often start on the horizontal axis, while the intermediate values are
floating columns. Because of this "look", waterfall charts are also called bridge charts.


PowerPoint UI
-------------

To create a waterfall chart by hand in PowerPoint:

1. Click Insert > Insert Waterfall or Stock chart > Waterfall.

You can also use the All Charts tab in Recommended Charts to create a waterfall chart.


XML semantics
-------------

* I think this is going to get into blank cells if multiple series are
  required. Might be worth checking out how using NA() possibly differs as to
  how it appears in the XML.

* There is a single `c:ser` element. Inside are `c:xVal`, `c:yVal`, and
  `c:bubbleSize`, each containing the set of points for that "series".


XML specimen
------------

.. highlight:: xml

Waterfall charts are not supported by PowerPoint versions prior to Office 2016. This is part of the
`2014 Chart Extension Schema <https://docs.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/e2723b0a-9120-42a5-bd11-c252ccb13c1e>`_
that also includes these charts:

- ``boxWhisker``
- ``clusteredColumn``
- ``funnel``
- ``paretoLine``
- ``regionMap``
- ``sunburst``
- ``treemap``
- ``waterfall``

The chart extension schema is defined in ``[Content_Types].xml``

``[Content_Types].xml``::

  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <!-- other types -->
    <Override PartName="/ppt/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>
  </Types>


It is used in the `slide1.xml.rels <cht-waterfall-chart/>` file like this::

``slide1.xml.rels``::

    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/>

The first line is the relationship for normal charts (e.g. a bar chart). The second has the relationship for a waterfall chart.

XML for default waterfall chart uses an Excel sheet with this data structure::

              Series1
    Category 1    100
    Category 2     20
    Category 3     50
    Category 4    -40
    Category 5    130
    Category 6    -60
    Category 7     70
    Category 8    140

The waterfall chart appears inside an `<mc:AlternateContent>` block to accommodate
non-supporting versions. The supporting alternative is a `<p:graphicFrame>` element. The fallback
alternative is a `<p:pic>` element and represents a static image of the chart.

``ppt/slides/slide1.xml``::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
      <p:spTree>
        <p:nvGrpSpPr>
          <p:cNvPr id="1" name=""/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="0" cy="0"/>
            <a:chOff x="0" y="0"/>
            <a:chExt cx="0" cy="0"/>
          </a:xfrm>
        </p:grpSpPr>
        <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">
            <p:graphicFrame>
              <p:nvGraphicFramePr>
                <p:cNvPr id="6" name="Chart 5">
                  <a:extLst>
                    <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                      <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{095E6047-871E-4A6A-93C5-2288E00B8DBB}"/>
                    </a:ext>
                  </a:extLst>
                </p:cNvPr>
                <p:cNvGraphicFramePr/>
                <p:nvPr/>
              </p:nvGraphicFramePr>
              <p:xfrm>
                <a:off x="2032000" y="719666"/>
                <a:ext cx="8128000" cy="5418667"/>
              </p:xfrm>
              <a:graphic>
                <a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">
                  <cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>
                </a:graphicData>
              </a:graphic>
            </p:graphicFrame>
          </mc:Choice>
          <mc:Fallback>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="6" name="Chart 5">
                  <a:extLst>
                    <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                      <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{095E6047-871E-4A6A-93C5-2288E00B8DBB}"/>
                    </a:ext>
                  </a:extLst>
                </p:cNvPr>
                <p:cNvPicPr>
                  <a:picLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1"/>
                </p:cNvPicPr>
                <p:nvPr/>
              </p:nvPicPr>
              <p:blipFill>
                <a:blip r:embed="rId3"/>
                <a:stretch>
                  <a:fillRect/>
                </a:stretch>
              </p:blipFill>
              <p:spPr>
                <a:xfrm>
                  <a:off x="2032000" y="719666"/>
                  <a:ext cx="8128000" cy="5418667"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
              </p:spPr>
            </p:pic>
          </mc:Fallback>
        </mc:AlternateContent>
      </p:spTree>
      <p:extLst>
        <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
          <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="4164693524"/>
        </p:ext>
      </p:extLst>
    </p:cSld>
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </p:sld>


``ppt/charts/_rels/chartEx1.xml``::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>
  </Relationships>

``ppt/charts/chartEx1.xml``::

  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cx:chartSpace
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
      <cx:chartData>
          <cx:externalData r:id="rId1" cx:autoUpdate="0" />
          <cx:data id="0">
              <cx:strDim type="cat">
                  <cx:f>Sheet1!$A$2:$A$9</cx:f>
                  <cx:lvl ptCount="8">
                      <cx:pt idx="0">Category 1</cx:pt>
                      <cx:pt idx="1">Category 2</cx:pt>
                      <cx:pt idx="2">Category 3</cx:pt>
                      <cx:pt idx="3">Category 4</cx:pt>
                      <cx:pt idx="4">Category 5</cx:pt>
                      <cx:pt idx="5">Category 6</cx:pt>
                      <cx:pt idx="6">Category 7</cx:pt>
                      <cx:pt idx="7">Category 8</cx:pt>
                  </cx:lvl>
              </cx:strDim>
              <cx:numDim type="val">
                  <cx:f>Sheet1!$B$2:$B$9</cx:f>
                  <cx:lvl ptCount="8" formatCode="General">
                      <cx:pt idx="0">100</cx:pt>
                      <cx:pt idx="1">20</cx:pt>
                      <cx:pt idx="2">50</cx:pt>
                      <cx:pt idx="3">-40</cx:pt>
                      <cx:pt idx="4">130</cx:pt>
                      <cx:pt idx="5">-60</cx:pt>
                      <cx:pt idx="6">70</cx:pt>
                      <cx:pt idx="7">140</cx:pt>
                  </cx:lvl>
              </cx:numDim>
          </cx:data>
      </cx:chartData>
      <cx:chart>
          <cx:title pos="t" align="ctr" overlay="0" />
          <cx:plotArea>
              <cx:plotAreaRegion>
                  <cx:series layoutId="waterfall" uniqueId="{FF3ADC76-EE77-455F-A520-1908E1D01E6B}">
                      <cx:tx>
                          <cx:txData>
                              <cx:f>Sheet1!$B$1</cx:f>
                              <cx:v>Series1</cx:v>
                          </cx:txData>
                      </cx:tx>
                      <cx:dataLabels pos="outEnd">
                          <cx:visibility seriesName="0" categoryName="0" value="1" />
                      </cx:dataLabels>
                      <cx:dataId val="0" />
                      <cx:layoutPr>
                          <cx:subtotals>
                              <cx:idx val="0" />
                              <cx:idx val="4" />
                              <cx:idx val="7" />
                          </cx:subtotals>
                      </cx:layoutPr>
                  </cx:series>
              </cx:plotAreaRegion>
              <cx:axis id="0">
                  <cx:catScaling gapWidth="0.5" />
                  <cx:tickLabels />
              </cx:axis>
              <cx:axis id="1">
                  <cx:valScaling />
                  <cx:majorGridlines />
                  <cx:tickLabels />
              </cx:axis>
          </cx:plotArea>
          <cx:legend pos="t" align="ctr" overlay="0" />
      </cx:chart>
  </cx:chartSpace>

``ppt/charts/colors1.xml``::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
    <a:schemeClr val="accent1"/>
    <a:schemeClr val="accent2"/>
    <a:schemeClr val="accent3"/>
    <a:schemeClr val="accent4"/>
    <a:schemeClr val="accent5"/>
    <a:schemeClr val="accent6"/>
    <cs:variation/>
    <cs:variation>
      <a:lumMod val="60000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="80000"/>
      <a:lumOff val="20000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="80000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="60000"/>
      <a:lumOff val="40000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="50000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="70000"/>
      <a:lumOff val="30000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="70000"/>
    </cs:variation>
    <cs:variation>
      <a:lumMod val="50000"/>
      <a:lumOff val="50000"/>
    </cs:variation>
  </cs:colorStyle>

``ppt/charts/style1.xml``::

  <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
  <cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="395">
    <cs:axisTitle>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1197"/>
    </cs:axisTitle>
    <cs:categoryAxis>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="1197"/>
    </cs:categoryAxis>
    <cs:chartArea mods="allowNoFillOverride allowNoLineOverride">
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="bg1"/>
        </a:solidFill>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="1330"/>
    </cs:chartArea>
    <cs:dataLabel>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1197"/>
    </cs:dataLabel>
    <cs:dataLabelCallout>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="dk1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="lt1"/>
        </a:solidFill>
        <a:ln>
          <a:solidFill>
            <a:schemeClr val="dk1">
              <a:lumMod val="25000"/>
              <a:lumOff val="75000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="1197"/>
      <cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">
        <a:spAutoFit/>
      </cs:bodyPr>
    </cs:dataLabelCallout>
    <cs:dataPoint>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:fillRef>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
      </cs:spPr>
    </cs:dataPoint>
    <cs:dataPoint3D>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:fillRef>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
      </cs:spPr>
    </cs:dataPoint3D>
    <cs:dataPointLine>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="28575" cap="rnd">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:dataPointLine>
    <cs:dataPointMarker>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:fillRef>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="lt1"/>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
    </cs:dataPointMarker>
    <cs:dataPointMarkerLayout symbol="circle" size="5"/>
    <cs:dataPointWireframe>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="28575" cap="rnd">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:dataPointWireframe>
    <cs:dataTable>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="1197"/>
    </cs:dataTable>
    <cs:downBar>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="dk1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="dk1">
            <a:lumMod val="65000"/>
            <a:lumOff val="35000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="65000"/>
              <a:lumOff val="35000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
    </cs:downBar>
    <cs:dropLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="35000"/>
              <a:lumOff val="65000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:dropLine>
    <cs:errorBar>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="65000"/>
              <a:lumOff val="35000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:errorBar>
    <cs:floor>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:floor>
    <cs:gridlineMajor>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:gridlineMajor>
    <cs:gridlineMinor>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:gridlineMinor>
    <cs:hiLoLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="75000"/>
              <a:lumOff val="25000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:hiLoLine>
    <cs:leaderLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="35000"/>
              <a:lumOff val="65000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:leaderLine>
    <cs:legend>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1197"/>
    </cs:legend>
    <cs:plotArea mods="allowNoFillOverride allowNoLineOverride">
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:plotArea>
    <cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:plotArea3D>
    <cs:seriesAxis>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="1197"/>
    </cs:seriesAxis>
    <cs:seriesLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat">
          <a:solidFill>
            <a:srgbClr val="D9D9D9"/>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:seriesLine>
    <cs:title>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1862"/>
    </cs:title>
    <cs:trendline>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="19050" cap="rnd">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="sysDash"/>
        </a:ln>
      </cs:spPr>
    </cs:trendline>
    <cs:trendlineLabel>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1197"/>
    </cs:trendlineLabel>
    <cs:upBar>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="dk1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="lt1"/>
        </a:solidFill>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
    </cs:upBar>
    <cs:valueAxis>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1197"/>
    </cs:valueAxis>
    <cs:wall>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:wall>
  </cs:chartStyle>



MS API Protocol
---------------

.. highlight:: vb.net

Create (unconventional) multi-series bubble chart in Excel::

    ActiveSheet.Shapes.AddChart2(395, xlWaterfall).Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet1!$A$2:$A$8"
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.Axes(xlValue).Select
    Selection.MajorTickMark = xlOutside


Related Schema Definitions
--------------------------

* https://docs.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/e2723b0a-9120-42a5-bd11-c252ccb13c1e

References
----------

* https://support.microsoft.com/en-us/office/create-a-waterfall-chart-8de1ece4-ff21-4d37-acd7-546f5527f185
