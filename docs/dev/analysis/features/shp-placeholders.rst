
Placeholders
============

A *placeholder* is a shape that specifies the position and possibly other
properties of the shape but has no actual content. From a user perspective,
... Placeholder shapes are inherited from a master ... A placeholder retains
its placeholder status when it is populated with content. If that content is
removed (in the PowerPoint application), the placeholder re-assumes its
original role and behaviors.


What form should a placeholder take?

* decorator pattern
* subclass of Shape
* wrapper of Shape
* mixin with Shape

One question is whether it should always be a different type (like from
construction onward) or whether it should only be treatd as a placeholder
when someone wants to perform placeholder-specific operations on it.

One consideration there is that certain behaviors (such as inherited
size/position) may always be different, and not be "special" placeholder-only
operations. Either subclassing would be necessary or tests in all differing
operations would need to be propagated. That's a pretty firm argument that
Placeholder needs to be a subclass of Shape (autoshape). How substituted
shapes are different would need to be clarified.

A `<p:pic>` shape from a placeholder doesn't have its own position and size.

This would indicate that all placeholder types would require their own
subclass, so the base shape operations were right. That might argue for
run-time inheritance, like make the shape, then if it's a placeholder then
decorate it before sending it back.

One option would be to add a mixin when required and modify behavior for when
it's a placeholder (ughh, placeholder as subclass seems better).

The general notion of inheriting properties will apply to quite a few
possible properties, will want a generalized way to access those.

>> Shape itself should certainly return None if it doesn't have one of those
properties.


Implementation notions
----------------------

* [ ] Start with MasterPlaceholders collection as thin proxy, to shape that
      out.

* [ ] maybe the key is to distinguish placeholder inheritance behaviors from
      placeholder element properties

* [ ] Start with MasterPlaceholder

  + [ ] Inherits from Shape
  + [ ] Maybe has a PlaceholderElement mixin ... ? (type, orient, sz, idx)
  + [ ] has_placeholder_properties ... ?  with separate PlaceholderProperties
        object?, perhaps returns None if not a placeholder shape

* [ ] Maybe create BasePlaceholder and subclass to MasterPlaceholder. I suspect
      at least there will be different behavior in that it doesn't try to
      inherit properties from a parent object.

* [ ] Add Placeholders, possibly subclassing to MasterPlaceholders later if
      needed.

* [ ] Maybe distinction between *placeholder shape* and whatever is
      a *placeholder element*. A placeholder shape *has* a placeholder
      element, but it's not the only type of shape that can have one.

* [ ] HOWEVER: A shape with a <p:ph> element exhibits inheritance behaviors
      on shape properties

* [ ] Thin proxy as described elsewhere on this page. Selects only `<p:sp>`
      elements that have a `<p:ph>` descendant. Instantiates `Placeholder`
      instance for each. Is iterable and has lookup, but does not instantiate
      a hard sequence like a tuple or list. Consider memoized. Think that
      through though.

      A hard sequence on first call might be a lot easier than memoized.

      Maybe all that's necessary is a dirty flag to reload on XML load or
      something.

      Maybe an expire cache call or something.

      Nah, who wants to keep the cache updated on deletes and adds.

      Maybe memoized, but that would be all.

* [ ] Add lookup by idx, maybe .get_by_idx(idx) that returns None if not found.
      Consider raising KeyError instead. Decision should hinge on whether can
      predict from layout whether placeholder is present in parent or not.
      I think not. Better go with None.

I think the important distinction is between a substituted placeholder and
a regular one.


Shapes that can have a substituted placeholder
----------------------------------------------

* Table
* Picture
* SmartArt and Chart, maybe all similar versions of GraphicalObject, like Table


blah
----

Two variants of placeholder must be considered. An autoshape placeholder
(which includes empty placeholders) and a substituted placeholder.


What operations must a placeholder support?

* clone slide placeholder from layout placeholder
* add text
* insert table, etc., where supported by placeholder type and substitution
  state
* all? shape query operations? such as get effective size and position?
* all? shape operations, such as set position and size


Access protocol
---------------

Collections

Placeholders
MasterPlaceholders
LayoutPlaceholders
SlidePlaceholders

Placeholder ...

Placeholder is different after substitution ... should it be included in set
anymore after substitution?

All thin proxies on the XML, retaining no state, including not retaining
a reference to any XML elements. Access is gained by querying Slide._element
or equivalent as required on a call-by-call basis. A calculated property may
hide any complexities involved.

::

    >>> master_placeholders = SlideMaster.placeholders
    >>> title_ph = master_placeholders.get_by_idx(0)
    >>> title_ph.left, title_ph.top
    (457200, 274638)
    >>> title_ph.width, title_ph.height
    (8229600, 1143000)


Definitions
-----------

placeholder shape
    A shape on a slide that inherits from a layout placeholder.

layout placeholder
    a shorthand name for the placeholder shape on the slide layout from which
    a particular placeholder on a slide inherits shape properties

master placeholder
    the placeholder shape on the slide master which a layout placeholder
    inherits from, if any.


Identification and linkage
--------------------------

... has id, which uniquely identifies shape on slide. idx value identifies
the layout placeholder it inherits from ...


Inheritance behaviors
---------------------

A placeholder shape on a slide is initially little more than a reference to
its "parent" placeholder shape on the slide layout. If it is a placeholder
shape that can accept text, it contains a `<p:txBody>` element. Position,
size, and even geometry are inherited from the layout placeholder, which may
in turn inherit one or more of those properties from a master placeholder.


Substitution behaviors
----------------------

Content may be placed into a placeholder shape two ways, by *insertion* and
by *substitution*. Insertion is simply placing the text insertion point in
the placeholder and typing or pasting in text. Substitution occurs when an
object such as a table or picture is inserted into a placeholder by clicking
on a placeholder button.

An empty placeholder is always a `<p:sp>` (autoshape) element. When an object
such as a table is inserted into the placehoder by clicking on a placeholder
button, the `<p:sp>` element is replaced with the appropriate new shape
element, a table element in this case. The `<p:ph>` element is retained in
the new shape element and preserves the linkage to the layout placeholder
such that the 'empty' placeholder shape can be restored if the inserted
object is deleted.


Operations
----------

* clone on slide create
* query inherited property values
* substitution


Behavior
--------

* Content of a placeholder shape is retained and displayed, even when the
  slide layout is changed to one without a matching layout placeholder.

* The behavior when placeholders are added to a slide layout (from the slide
  master) may also be worth characterizing.

  + ... show master placeholder ...
  + ... add (arbitrary) placeholder ...


Placeholder types
-----------------

* Title (always inherits from master, although layout may override)
* Vertical Title (also inherits from master)
* Content
* Vertical content
* Text
* Vertical text
* Chart
* Table
* SmartArt
* Media
* Clip Art
* Picture


Sample XML
----------

.. highlight:: xml

Baseline textbox shape::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="2" name="TextBox 1"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="3016188" y="3273093"/>
          <a:ext cx="1133644" cy="369332"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="none" rtlCol="0">
          <a:spAutoFit/>
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" dirty="0" smtClean="0"/>
            <a:t>Some text</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" dirty="0"/>
        </a:p>
      </p:txBody>
    </p:sp>


Content placeholder::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="5" name="Content Placeholder 4"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph idx="1"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:endParaRPr lang="en-US"/>
        </a:p>
      </p:txBody>
    </p:sp>


Notable differences:

* placeholder has `<a:spLocks>` element
* placeholder has `<p:ph>` element
* placeholder has no `<p:spPr>` child elements, this may imply both that:
  
  + all shape properties are initially inherited from the layout placeholder,
    including position, size, and geometry
  + any specific shape property value may be overridden by specifying it on
    the inheriting shape


Matching slide layout placeholder::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Content Placeholder 2"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph idx="1"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr lvl="0"/>
          <a:r>
            <a:rPr lang="en-US" smtClean="0"/>
            <a:t>Click to edit Master text styles</a:t>
          </a:r>
        </a:p>
        <a:p>
          ... and others through lvl="4", five total
        </a:p>
      </p:txBody>
    </p:sp>


Matching slide master placeholder::

    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Text Placeholder 2"/>
        <p:cNvSpPr>
          <a:spLocks noGrp="1"/>
        </p:cNvSpPr>
        <p:nvPr>
          <p:ph type="body" idx="1"/>
        </p:nvPr>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="457200" y="1600200"/>
          <a:ext cx="8229600" cy="4525963"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </p:spPr>
      <p:txBody>
        <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440"
                  bIns="45720" rtlCol="0">
          <a:normAutofit/>
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:pPr lvl="0"/>
          <a:r>
            <a:rPr lang="en-US" smtClean="0"/>
            <a:t>Click to edit Master text styles</a:t>
          </a:r>
        </a:p>
        <a:p>
          ... and others through lvl="4", five total
        </a:p>
      </p:txBody>
    </p:sp>
 

Note:

* master specifies size, position, and geometry
* master specifies text body properties, such as margins (inset) and autofit
