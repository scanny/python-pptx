
Presentation sections (``p14:sectionLst``)
==========================================

Overview
--------

PowerPoint's *Sections* feature (introduced in 2010) groups consecutive
slides into named, collapsible sections in the Navigation pane. Sections
are stored under a 2010-extension element ``p14:sectionLst`` nested inside
the presentation's ``p:extLst`` container and scoped by the canonical URI
``{521415D9-36F7-43E2-AB2F-B90AF26B5E84}``.

Each section has a display ``name``, a stable GUID ``id``, and a list of
slides identified by the ``id`` value of their entry in the main
``p:sldIdLst`` (*not* by ``r:id``).


Public API
----------

::

    >>> prs = Presentation(path)
    >>> prs.sections
    <pptx.presentation.Sections object>
    >>> len(prs.sections)
    0

    >>> # add a section with two existing slides assigned to it
    >>> section = prs.sections.add_section(
    ...     "Intro", slides=[prs.slides[0], prs.slides[1]]
    ... )
    >>> section.name
    'Intro'
    >>> section.id
    '{EA541957-60A5-47A7-8DF8-F51D8F183008}'
    >>> section.slides
    (<pptx.slide.Slide ...>, <pptx.slide.Slide ...>)

    >>> section.name = "Overview"
    >>> section.add_slide(prs.slides[2])
    >>> section.remove_slide(prs.slides[0])
    >>> prs.sections.remove(section)


XML specimens
-------------

.. highlight:: xml

Real-world ``presentation.xml`` fragment produced by PowerPoint 2016::

  <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
    <p:sldMasterIdLst>...</p:sldMasterIdLst>
    <p:sldIdLst>
      <p:sldId id="256" r:id="rId2"/>
      <p:sldId id="257" r:id="rId3"/>
      <p:sldId id="258" r:id="rId4"/>
    </p:sldIdLst>
    <!-- ... -->
    <p:extLst>
      <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
        <p14:sectionLst>
          <p14:section name="Intro"
                       id="{9B1ED94F-0448-4B93-BA15-4DB6F29DCB9C}">
            <p14:sldIdLst>
              <p14:sldId id="256"/>
              <p14:sldId id="257"/>
            </p14:sldIdLst>
          </p14:section>
          <p14:section name="Body"
                       id="{7FBA8AEB-2F7A-4E32-A8BC-20DB21AF1C7E}">
            <p14:sldIdLst>
              <p14:sldId id="258"/>
            </p14:sldIdLst>
          </p14:section>
        </p14:sectionLst>
      </p:ext>
    </p:extLst>
  </p:presentation>


Key observations
~~~~~~~~~~~~~~~~

* Sections are addressed by the *slide-id integer* (``p:sldId/@id`` in the
  main ``p:sldIdLst``), **not** by ``r:id``. This avoids coupling
  membership to a particular relationship-allocation order and lets
  PowerPoint re-order the ``p:sldIdLst`` without having to rewrite the
  section list.
* A slide may belong to at most one section. PowerPoint does not enforce
  that *every* slide belongs to a section — an unassigned slide shows up
  under a synthetic "Default Section" in the UI.
* Section names are not unique. Section ``id`` values are GUIDs and
  **are** unique; python-pptx raises when an explicit ``id=`` duplicates
  one already present.
* The ``p14:sldIdLst`` child of a section is required by PowerPoint even
  when empty; python-pptx always writes it.


Schema
------

The extension-list scaffolding is ECMA-376 Part 1 (``pml.xsd``):

.. highlight:: xml

::

  <xsd:complexType name="CT_Extension">
    <xsd:sequence>
      <xsd:any processContents="lax" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="uri" type="xsd:token" use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_ExtensionList">
    <xsd:sequence>
      <xsd:group ref="EG_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>

The section-list content under ``p:ext`` is *not* part of the Part 1
schemas — it's defined by Microsoft's ``[MS-PPTX]`` extension
specification. The relevant types (reconstructed from real PowerPoint
output) are::

  <xsd:element name="sectionLst" type="CT_SectionList"/>

  <xsd:complexType name="CT_SectionList">
    <xsd:sequence>
      <xsd:element name="section" type="CT_Section"
                   minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_Section">
    <xsd:sequence>
      <xsd:element name="sldIdLst" type="CT_SectionSlideIdList"/>
      <xsd:element name="extLst"   type="CT_ExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="name" type="xsd:string" use="required"/>
    <xsd:attribute name="id"   type="ST_Guid"    use="required"/>
  </xsd:complexType>

  <xsd:complexType name="CT_SectionSlideIdList">
    <xsd:sequence>
      <xsd:element name="sldId" type="CT_SectionSlideIdListEntry"
                   minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_SectionSlideIdListEntry">
    <xsd:attribute name="id" type="ST_SlideId" use="required"/>
  </xsd:complexType>


Schema uncertainty
------------------

The section ``id`` attribute is treated by PowerPoint as a GUID in the
canonical curly-braced uppercase form. python-pptx accepts the form
``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`` (hexadecimal digits, any case)
for user-supplied ids and auto-generates new ids in uppercase with
braces.  The MS-PPTX documentation does not formally specify a
``CT_Section/extLst`` content model — PowerPoint has been observed to
write an empty ``<p14:extLst/>`` there for long presentations as a
scaffolding slot. python-pptx preserves any such nested extension
elements on read but does not currently expose them through the
high-level API.


Design notes
------------

* ``Presentation.sections`` is a :class:`Sections` collection object,
  returned lazily (``@lazyproperty``) — the same instance is returned on
  each access. An empty collection is returned when the presentation has
  no ``p14:sectionLst``; creating the first section materialises the
  ``p:extLst/p:ext/p14:sectionLst`` chain on demand.
* :meth:`Sections.remove` prunes empty scaffolding on last-section
  removal (``p14:sectionLst`` → ``p:ext`` → ``p:extLst``) so a
  round-tripped deck that had sections temporarily doesn't carry an
  empty extension block.
* Slide-to-section membership is resolved at access time (``Section.slides``)
  via the owning presentation's ``p:sldIdLst``; stale ``p14:sldId``
  entries whose ``id`` no longer matches any slide are silently skipped.
