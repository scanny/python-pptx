
Picture placeholder
===================

A picture placeholder is an unpopulated placeholder into which an image can
be inserted. This comprises three concrete placeholder types, a picture
placeholder, a clip art placeholder, and a content placeholder. Each of
these, and indeed any unpopulated placeholder, take the form of a `p:sp`
element in the XML. Once populated with an image, these placeholders become
a *placeholder picture* and take the form of a `p:pic` element.


Candidate protocol
------------------

Placeholder access::

  >>> picture_placeholder = slide.placeholders[1]  # keyed by idx, not offset
  >>> picture_placeholder
  <pptx.shapes.placeholder.PicturePlaceholder object at 0x100830510>
  >>> picture_placeholder.shape_type
  MSO_SHAPE_TYPE.PLACEHOLDER (14)

PicturePlaceholder.insert_picture()::

  >>> placeholder_picture = picture_placeholder.insert_picture('image.png')
  >>> placeholder_picture
  <pptx.shapes.placeholder.PlaceholderPicture object at 0x10083087a>
  >>> placeholder_picture.shape_type
  MSO_SHAPE_TYPE.PLACEHOLDER (14)


Example XML
-----------

.. highlight:: xml

A picture-only layout placholder::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="3" name="Picture Placeholder 2"/>
      <p:cNvSpPr>
        <a:spLocks noGrp="1"/>
      </p:cNvSpPr>
      <p:nvPr>
        <p:ph type="pic" idx="1"/>
      </p:nvPr>
    </p:nvSpPr>
    <p:spPr>
      <a:xfrm>
        <a:off x="1792288" y="612775"/>
        <a:ext cx="5486400" cy="4114800"/>
      </a:xfrm>
    </p:spPr>
    <p:txBody>
      <a:bodyPr/>
      <a:lstStyle>
        <!-- a:lvl{n}pPr for n in 1:9 -->
      </a:lstStyle>
      <a:p>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </p:txBody>
  </p:sp>

An unpopulated picture-only placeholder on a slide::

  <p:sp>
    <p:nvSpPr>
      <p:cNvPr id="7" name="Picture Placeholder 6"/>
      <p:cNvSpPr>
        <a:spLocks noGrp="1"/>
      </p:cNvSpPr>
      <p:nvPr>
        <p:ph type="pic" idx="1"/>
      </p:nvPr>
    </p:nvSpPr>
    <p:spPr/>
  </p:sp>

A picture-only placeholder populated with an image::

  <p:pic>
    <p:nvPicPr>
      <p:cNvPr id="8" name="Picture Placeholder 7"
               descr="aphrodite.brinsmead.jpg"/>
      <p:cNvPicPr>
        <a:picLocks noGrp="1" noChangeAspect="1"/>
      </p:cNvPicPr>
      <p:nvPr>
        <p:ph type="pic" idx="1"/>
      </p:nvPr>
    </p:nvPicPr>
    <p:blipFill>
      <a:blip r:embed="rId2">
        <a:extLst>
          <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
            <a14:useLocalDpi xmlns:a14="http://../drawing/2010/main" val="0"/>
          </a:ext>
        </a:extLst>
      </a:blip>
      <a:srcRect t="20000" b="20000"/>  <!-- 20% crop, top and bottom -->
      <a:stretch>
        <a:fillRect/>
      </a:stretch>
    </p:blipFill>
    <p:spPr/>
  </p:pic>
