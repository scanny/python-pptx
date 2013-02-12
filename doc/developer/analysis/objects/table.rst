=====
Table
=====

:Updated:  2013-02-11
:Author:   Steve Canny
:Status:   **WORKING DRAFT**


Discovery protocol
==================

* (/) Review MS API documentation
* (/) Inspect minimal XML produced by PowerPoint® client
* (.) Review and document relevant schema elements


Notes
=====

Produced XML inspection
-----------------------

* a ``tableStyles.xml`` part is fleshed out substantially; looks like it's
  populated from built-in defaults "Medium Style 2 - Accent 1". It appears to
  specify colors indirectly by reference to theme-specified values.

MS API Notes
------------

MS API method to add a table is::

    Shapes.AddTable(NumRows, NumColumns, Left, Top, Width, Height)

* There is a HasTable property on Shape to indicate the shape "has" a table.
  Seems like "is" a table would be more apt, but I'm still looking :)

`Table Members`_ page on MSDN.

Most interesting ``Table`` members:

* ``Cell(row, col)`` method to access individual cells.
* ``Columns`` collection reference, with ``Add`` method (``Delete`` method is
  on ``Column`` object)
* ``Rows`` collection reference
* FirstCol and FirstRow boolean properties that indicate whether to apply
  special formatting from theme or whatever to first column/row.
* LastCol, LastRow, and HorizBanding, all also boolean with similar behaviors
* TableStyle read-only to table style in theme. Table.ApplyStyle() method is
  used to set table style.

Columns collection and Rows collection both have an Add() method

`Column Members`_ page on MSDN.

* Delete()
* Width property

`Cell Members`_ page on MSDN.

* Merge() and Split() methods
* Borders reference to Borders collection of LineFormat objects
* Shape reference to shape object that cell is or has.

`LineFormat Members`_ page on MSDN.

* ForeColor
* Weight


.. _Table Members:
   http://msdn.microsoft.com/en-us/library/office/ff745711(v=office.14).aspx

.. _Column Members:
   http://msdn.microsoft.com/en-us/library/office/ff746286(v=office.14).aspx

.. _Cell Members:
   http://msdn.microsoft.com/en-us/library/office/ff744136(v=office.14).aspx

.. _LineFormat Members:
   http://msdn.microsoft.com/en-us/library/office/ff745240(v=office.14).aspx


XML produced by PowerPoint® client
==================================

.. highlight:: xml

::

    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="2" name="Table 1"/>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="1524000" y="1397000"/>
        <a:ext cx="6096000" cy="741680"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid>
              <a:gridCol w="3048000"/>
              <a:gridCol w="3048000"/>
            </a:tblGrid>
            <a:tr h="370840">
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
            </a:tr>
            <a:tr h="370840">
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
              <a:tc>
                <a:txBody>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </a:txBody>
                <a:tcPr/>
              </a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>

Summary
=======

...


Description
===========


Behaviors
=========


Experiment(s)
-------------


Specifications
==============


Related Specifications
======================


Resources
=========

