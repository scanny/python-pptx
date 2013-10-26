#############
Graphic Frame
#############


``<p:graphicFrame>`` shape elements
===================================

... graphic frame is used to hold table, chart, perhaps other graphical
objects ...


XML produced by PowerPoint® client
----------------------------------

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
          ... table or chart element goes here ...
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>


