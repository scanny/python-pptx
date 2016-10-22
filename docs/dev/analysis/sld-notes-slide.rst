
Notes Slide
===========

A slide may have an associated notes page. The notes it contains are
displayed under the slide in the PowerPoint UI when in edit mode. The slide
notes are also shown in presenter mode, are displayed in Notes Pages view,
and are printed on those notes pages.

Internally, a notes page is referred to as a *notes slide*. This is sensible
because it is actually a specialized instance of a slide. It contains shapes,
many of which are placeholders, and allows inserting of new shapes such as
auto shapes, tables, and charts (although it is probably not common). Much of
the functionality for a notes slide is inherited from existing base classes.

A notes slide is created using the notes master as a template. A presentation
has no notes master when created from a template in PowerPoint. One is
created according to a PowerPoint-internal preset default the first time
it is needed, which is generally when a notes slide is created. It's possible
one can also be created by entering the Notes Master view and almost
certainly is created by editing the master found there (haven't tried it
though). A presentation can have at most one notes master.

On creation, certain placeholders (slide image, notes, slide number) are
copied from the notes master onto the new notes slide (if they have not been
removed from the master). These "cloned" placeholders inherit position, size,
and formatting from their corresponding notes master placeholder. If the
position, size, or formatting of a notes slide placeholder is changed, the
changed property is no long inherited (unchanged properties, however,
continue to be inherited).

The remaining three placeholders that can appear on the notes master (date,
header text, footer text) can optionally be made to appear on an individual
notes slide using the Header and Footer... dialog. This dialog also has
a button allowing all notes slides to be updated with the selected options.
There appears to be some heuristic logic that automatically propagates these
settings to notes for new slides as they are created.

A notes slide is not automatically created for each slide. Rather it is
created the first time it is needed, perhaps most commonly when notes text is
typed into the notes area under the slide in the UI.


Candidate Protocol
------------------

A notes slide is created on first access if one doesn't exist. Consequently,
Slide.has_notes_slide is provided to detect the presence of an existing notes
slide, avoiding creation of unwanted notes slide objects::

    >>> slide.has_notes_slide
    False

Because a notes slide is created on first access, Slide.notes_slide
unconditionally returns a |NotesSlide| object. Once created, the same
instance is returned on each call::

    >>> notes_slide = slide.notes_slide
    >>> notes_slide
    <pptx.notes.NotesSlide object at 0x10698c1e0>
    >>> slide.has_notes_slide
    True

Like any slide (slide, slide layout, slide master, etc.), a notes slide has
shapes and placeholders, as well as other standard properties::

    >>> notes_slide.shapes
    <pptx.shapes.shapetree.NotesSlideShapes object at 0x10698c1e0>
    >>> notes_slide.placeholders
    <pptx.shapes.shapetree.NotesSlidePlaceholders object at 0x10698f622>

The distinctive characteristic of a notes slide is the notes it contains.
These notes are contained in the text frame of the notes placeholder and are
manipulated using the same properties and methods as any shape textframe::

    >>> text_frame = notes_slide.notes_text_frame
    >>> text_frame.text = 'foobar'
    >>> text_frame.text
    'foobar'
    >>> text_frame.add_paragraph('barfoo')
    >>> text_frame.text
    'foobar\nbarfoo'


PowerPoint behavior
-------------------

* Notes page has its own view (View > Notes Page)

  + Notes can be edited from there too
  + It has a hide background graphics option and the other background settings
  + Notes have richer edit capabilities (such as ruler for tabs and indents)
    in the Notes view.

* A notes slide, once created, doesn't disappear when its notes content is
  deleted.

* NotesSlide placeholders inherit properties from the NotesMaster placeholder
  *of the same type*.

* For some reason PowerPoint adds a second theme for the NotesMaster, not
  sure what that's all about. I'll just add the default theme2.xml, then
  folks can avoid that if they encounter problems just by pre-loading
  a notesMaster of their own in their template.


Example XML
-----------

.. highlight:: xml

Empty notes page element::

  <p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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
      </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </p:notes>

Default notes page populated with three base placeholders::

  <p:notes
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      >
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
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="2" name="Slide Image Placeholder 1"/>
            <p:cNvSpPr>
              <a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>
            </p:cNvSpPr>
            <p:nvPr>
              <p:ph type="sldImg"/>
            </p:nvPr>
          </p:nvSpPr>
          <p:spPr/>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="3" name="Notes Placeholder 2"/>
            <p:cNvSpPr>
              <a:spLocks noGrp="1"/>
            </p:cNvSpPr>
            <p:nvPr>
              <p:ph type="body" idx="1"/>
            </p:nvPr>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:r>
                <a:rPr lang="en-US" smtClean="0"/>
                <a:t>Notes</a:t>
              </a:r>
              <a:endParaRPr lang="en-US" dirty="0"/>
            </a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="4" name="Slide Number Placeholder 3"/>
            <p:cNvSpPr>
              <a:spLocks noGrp="1"/>
            </p:cNvSpPr>
            <p:nvPr>
              <p:ph type="sldNum" sz="quarter" idx="10"/>
            </p:nvPr>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:fld id="{64BF21E3-C0F6-2742-868D-4CDFA1962D8A}" type="slidenum">
                <a:rPr lang="en-US" smtClean="0"/>
                <a:t>1</a:t>
              </a:fld>
              <a:endParaRPr lang="en-US"/>
            </a:p>
          </p:txBody>
        </p:sp>
      </p:spTree>
      <p:extLst>
        <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
          <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="347586568"/>
        </p:ext>
      </p:extLst>
    </p:cSld>
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </p:notes>


Related Schema Definitions
--------------------------

.. highlight:: xml

The root element of a notesSlide part is a `p:notes` element::

  <xsd:element name="notes" type="CT_NotesSlide"/>

  <xsd:complexType name="CT_NotesSlide">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="cSld"      type="CT_CommonSlideData"/>
      <xsd:element name="clrMapOvr" type="a:CT_ColorMappingOverride" minOccurs="0"/>
      <xsd:element name="extLst"    type="CT_ExtensionListModify"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="showMasterSp"     type="xsd:boolean" default="true"/>
    <xsd:attribute name="showMasterPhAnim" type="xsd:boolean" default="true"/>
  </xsd:complexType>

  <xsd:complexType name="CT_CommonSlideData">
    <xsd:sequence>
      <xsd:element name="bg"          type="CT_Background"       minOccurs="0"/>
      <xsd:element name="spTree"      type="CT_GroupShape"/>
      <xsd:element name="custDataLst" type="CT_CustomerDataList" minOccurs="0"/>
      <xsd:element name="controls"    type="CT_ControlList"      minOccurs="0"/>
      <xsd:element name="extLst"      type="CT_ExtensionList"    minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" use="optional" default=""/>
  </xsd:complexType>
