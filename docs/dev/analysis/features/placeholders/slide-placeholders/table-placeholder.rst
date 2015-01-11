
Table placeholder
=================

A table placeholder is an unpopulated placeholder into which a table can be
inserted. This comprises two concrete placeholder types, a table placeholder
and a content placeholder. Both take the form of a `p:sp` element in the XML.
Once populated with a table, these placeholders become a *placeholder graphic
frame* and take the form of a `p:graphicFrame` element.


Candidate protocol
------------------

Placeholder access::

  >>> table_placeholder = slide.placeholders[10]  # keyed by idx, not offset
  >>> table_placeholder
  <pptx.shapes.placeholder.TablePlaceholder object at 0x100830510>
  >>> table_placeholder.shape_type
  MSO_SHAPE_TYPE.PLACEHOLDER (14)

TablePlaceholder.insert_table()::

  >>> ph_graphic_frame = table_placeholder.insert_table(rows=2, cols=2)
  >>> ph_graphic_frame
  <pptx.shapes.placeholder.PlaceholderGraphicFrame object at 0x10083087a>
  >>> ph_graphic_frame.shape_type
  MSO_SHAPE_TYPE.PLACEHOLDER (14)
  >>> ph_graphic_frame.has_table
  True
  >>> table = ph_graphic_frame.table
  >>> len(table.rows), len(table.columns)
  (2, 2)


Example XML
-----------

.. highlight:: xml

A table-only layout placholder::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="3" name="Table Placeholder 2"/>
      <p:cNvSpPr>
        <a:spLocks noGrp="1"/>
      </p:cNvSpPr>
      <p:nvPr>
        <p:ph type="tbl" sz="quarter" idx="10"/>
      </p:nvPr>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="2743200" y="2057400"/>
        <a:ext cx="3657600" cy="2743200"/>
      </a:xfrm>
      <a:prstGeom prst="rect">
        <a:avLst/>
      </a:prstGeom>
    </p:spPr>
    <p:txBody>
      <a:bodyPr vert="horz"/>
      <a:lstStyle/>
      <a:p>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </p:txBody>
  </p:sp>

An unpopulated table-only placeholder on a slide::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="2" name="Table Placeholder 1"/>
      <p:cNvSpPr>
        <a:spLocks noGrp="1"/>
      </p:cNvSpPr>
      <p:nvPr>
        <p:ph type="tbl" sz="quarter" idx="10"/>
      </p:nvPr>
    </p:nvSpPr>
    <p:spPr/>
  </p:sp>

A table-only placeholder populated with a table::

  <p:graphicFrame>
    <p:nvGraphicFramePr>
      <p:cNvPr id="3" name="Table Placeholder 2"/>
      <p:cNvGraphicFramePr>
        <a:graphicFrameLocks noGrp="1"/>
      </p:cNvGraphicFramePr>
      <p:nvPr>
        <p:ph type="tbl" sz="quarter" idx="10"/>
        <p:extLst>
          <p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}">
            <p14:modId xmlns:p14="http://../..nt/2010/main" val="933747684"/>
          </p:ext>
        </p:extLst>
      </p:nvPr>
    </p:nvGraphicFramePr>
    <p:xfrm>
      <a:off x="2743200" y="2057400"/>
      <a:ext cx="3657600" cy="3337560"/>
    </p:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl>
          <a:tblPr firstRow="1" bandRow="1">
            <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
          </a:tblPr>
          <a:tblGrid>
            <a:gridCol w="457200"/>
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
          </a:tr>
        </a:tbl>
      </a:graphicData>
    </a:graphic>
  </p:graphicFrame>
