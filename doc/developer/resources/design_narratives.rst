=================
Design Narratives
=================

Narrative explorations into design issues, serving initially as an aid to
reasoning and later as a memorandum of the considerations undertaken during
the design process.


XML wrapper classes
===================

TODO
----
* `(1)` get subclassing working for oxml and oxml_base ...


XSD unmarshalling class community
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

... one approach might be to generate the classes off of a subset of the full
XML Schema, modifying schema element definitions where desired to prune
unused parts of the type graph

... possibly a TypeGraph object that acts as the container, initialize with a
root complexType node ... calls from ComplexType to graph like getTypeByName
and graph can create and return, that keeps type creation code up at the top
level and outside ComplexType ...

... I wonder how far I could get with just a list of (tag, typename,
is_required) tuples. Required elements would be created on construction, and
optional ones would be ...

The rub would come in spTree, and in paragraph (<a:p> where you have an
arbitrary sequence from a set (xsd:choice) of possible elements.

So tuples won't work for long, there needs to be a set of classes that
populate the list, maybe just Element and Choice. These situations might be
rare enough that subclassing the relatively few would work for a start and
then learn what code to generate step by step.


Acceptance Tests
----------------

* optional element is automatically inserted on first access
* optional element is not present on construction of parent
* optional element is not created if not in lxml graph
* optional element added at run time is inserted in correct sequence
* ...


Requirements
------------

Pending requirements
^^^^^^^^^^^^^^^^^^^^

* cascading construction (unmarshaling) based on an existing lxml element
  graph

* generation of entire module file from template so no manual intervention is
  required
* work out a good way to subclass to add specialized functionality to selected
  classes

* repeating elements

  * txBody.p for example

* optional elements ...

  * create on first access
  * ? how respect element order when inserting a new element after
    construction? ... need additional metadata

* choice groups ...


Completed requirements
^^^^^^^^^^^^^^^^^^^^^^

* XML access classes are generated
* child elements can be accessed via dotted notation
* minimal/required sub-tree elements are constructed in cascade on
  construction of root element
* attributes can be accessed via dotted notation; required attributes get no
  special treatment, it's up to the user to make sure they are set with a
  valid value.


Open Issues
-----------


``<xsd:choice>`` handling
^^^^^^^^^^^^^^^^^^^^^^^^^

...




Purposes
--------

1. ... compact and readable access ... using dotted member notation instead of
   ``lxml`` notation in higher level abstraction domains ...

#. Encapsulate and hide

#. Reuse of recurring elements ...


Constructors
------------

... two construction modes:

1. (Roughly) *unmarshal* an existing lxml element

2. Construct a *new* lxml element from scratch



End-user API facade
===================

Sat Jan 26 2013

It might be a design improvement if the end-user API was via a single module
that mediated access to (at least) implementation classes.

... might include:

* Presentation

  * Presentation(path=None) constructor, to use default template
  * .save(path) -- save the presentation as *path*
  * slidemasters attribute (immutable sequence)
  * slides attribute (immutable sequence)

A related possibility is segregating presentation-related parts to a separate
module, perhaps named prs_parts. Any parts that were general to more than one
Open XML document could be placed in a separate module, and parts for WML and
SML could be added in separate modules to fill out the set.


Design Narrative -- Model-side relationships
============================================

**always up-to-date principle**

  Model-side relationships are maintained as new parts are added or existing
  parts are deleted. Relationships for generic parts are maintained from load
  and delivered back for save without change.

The :doc:`../protocols/relationships` page contains documentation for
the :ref:`relationship-related-protocol`.

Interface
---------

=======  ============  =======================================================
attr     client        purpose
=======  ============  =======================================================
rId      presentation  part association during unmarshaling
reltype  presentation  allow relationships to be selected by type
target   presentation  get specifics and content
-------  ------------  -------------------------------------------------------
element  packaging     Package needs this to save pptx
=======  ============  =======================================================

* Unlikely to need .source attribute in interface because only way to get to
  the relationships is by traversing the source.

* All the business of baseURI and target like a relative URI are things
  Relationship can safely hide from clients.



Design Narrative --- Text API
=============================

TextFrame.delete_text()
-----------------------

* A txBody element must have at least one paragraph element, so this method
  would delete all the paragraphs except one (perhaps the first or last one)
  and remove all its text.

* A <a:p> element is not required to contain any child elements, so could just
  empty it of all children or perhaps leave something like this::

    <a:p>
      <a:endParaRPr lang="en-US"/>
    </a:p>


... text is a fairly complicated bit ... deceptively sophisticated one might
fairly say.

* ... will need both simple and sophisticated ways of dealing with text ...

* Use TextFrame2, apparently it's an enhanced version of legacy TextFrame

    TextFrame2 exposes the new text properties introduced in PPT 2007.

* `TextRange Members`_ page on MSDN Office Interop

.. _TextRange Members:
   http://msdn.microsoft.com/en-us/library/microsoft.office.interop
   .powerpoint.textrange_members(v=office.14).aspx


Design Narrative --- Full unmarshaling vs. ElementTree
======================================================

... question of whether a full unmarshaling of part XML using something like
pyXB or generateDS is a sensible design option a bit later on after basic
functionality is completed and perhaps scaling becomes more difficult (if it
does) with just using the lxml.ElementTree objects.


Design Narrative --- blob to element to blob life-cycle
=======================================================

* (?) Detecting is_xml for both loaded and new parts (call .partname?)

* (?) What about added binary parts like Image?

* There's a bit of a smell to this in that redundancy of ordering info is
  added to collections. Operations like reordering adding and deleting will
  need to operate on both the collection and the XML.

* I suppose sub-classes can override _blob() if they need to do something
  special.

* add_part(element) methods will need to take care of adding _element for
  their part.


Hypothesis
----------

blob > element > blob lifecycle can be completely handled in BasePart.

xml elements access self._element. Maybe change _load_blob to __load_blob.

::
    **in _load():**

    if self.is_xml:
        self._element = etree.fromstring(pkgpart.blob)
    else:
        self._load_blob = pkgpart.blob

    **in _blob():**

    if self.is_xml:
        return etree.tostring(self._element, ...)
    else:
        return self._load_blob

----

* If we start with the principle that all operations will be conducted on the
  XML elements and no separate attributes will be stored ...

* We might keep references to parts of the element, but changes to those parts
  are changes to the root reference. So unless we break that, everything
  should work fairly seamlessly.


Slide attributes -- draft list
------------------------------

* overall shape tree transform (not sure what this is exactly)
* shape tree (root group shape)


GroupShape attributes -- draft list
-----------------------------------

* id (slide internal scope I think)
* group_shape_name, top level one might be slide name
* transform (x, y, cx, cy, etc.)
* shapes (sp, groupshape, pic, some others)


Shape attributes -- draft list
------------------------------

* id (slide internal)
* name (assigned)
* locks (like no grouping)
* placeholder (id="0" is title, and id defaults to 0, so title if no id
  specified)
* text


Open issues parking lot
=======================

* Principle: No loaded bits will be removed from the XML. I'm thinking that
  means that unless we keep track of which are loaded and which are new, that
  drives the decision to work with the XML in-place.

* ... there's the issue of whether library will be used to fully unmarshal
  existing documents and manipulate them. The challenge of writing brand-new
  documents is simpler I think.

* There is some irresolution around a possible distinction between part
  classes and element classes, particularly a possible distinction between
  a part class and it's root element. Something to continue to noodle.



Design Narrative -- Using Sphinx for library documentation
==========================================================

Conundrum: How to use the autodoc selectively so a pleasing layout is
produced.

Important things include::

   .. automodule:: <module_name>

   .. autoclass:: <class_name>

   .. autofunction:: <function_name>

The key to using these features is the :members: attribute. If:

You don’t include it at all, only the docstring for the object is brought in:
You just use :members: with no arguments, then all public functions, classes,
and methods are brought it that have docstring. If you explictly list the
members like :members: fn0, class0, _fn1 those explict members are brought.



Design Narrative -- Part blob lifecycle
=======================================

Recorded: 2012-12-24 11:46 PM

* pptx.packaging.Part stores part content as blob

* if pptx.presentation.Part persists the blob and serves it back to
  pkg.marshal, round-trip should work

* presentation parts that unmarshal blob need to provide a blob property that
  marshaling can use to access part content.


TODO:

* (/) refactor pptx.packaging.Part.load to unconditionally save blob
* (/) locate part.write_element and replace with write_blob
* (/) remove element attribute from pptx.packaging.Part

----

* Simplify packaging module by working only with blobs whenever possible

* write_element is handy for items that packaging works on directly, like cti
  and rels items. So no urgent need to get rid of it, just always write parts
  as blobs.

* presentation.Parts need ._blob attribute in their interface so packaging can
  uniformly access contents for marshaling.

   Rationale:

   * _blob is required for binary objects, so at least some parts must have
     that attribute.

   * A need to determine whether to call _blob or element to access part
     contents would complicate marshaling and unmarshalling code.

   * A static part doesn't need to access its blob, it can just carry it until
     it's needed for marshaling.

* principle: packaging.Part always gets and stores blob (lowest common
  denominator).

* Need a blob round-trip between package to model and back


Design Narrative -- Model Load
==============================

Recorded: 2012-12-22 11:01 PM

* __loadwalk()

Requirements
------------

* All parts are constructed exactly once.

* All part relationships are created and populated with target part.

* (?) What to do with package relationships?

* Parts of types with a custom Part-subclass are instances of the custom
  sub-class.

* Custom sub-class instances are triggered to perform unmarshalling once the
  part and its relationships are completely loaded. It might be sensible to
  wait and do this once all parts and relationships are loaded, with a second
  walk or similar implementation.

* Could be that propagating control flow rather than recursive might work
  best, so that local context is kept local to the package or part.

::
    
    def __pkg_level_load(pkgrels):
        # keep track of which parts are already loaded
        part_dict = {}
        
        for pkgrel in pkgrels:
            # unpack working values for part to be loaded
            reltype = pkgrel.reltype
            pkgpart = pkgrel.target_part
            partname = pkgpart.partname
            content_type = pkgpart.content_type
            
            # create target part
            if partname in part_dict:
                part = part_dict[partname]
            else:
                part = Part(reltype, content_type)
                part_dict[partname] = part
                part.load(pkgpart, part_dict)
            
            # create model-side package relationship
            rId = pkgrel.rId
            model_rel = Relationship(rId, reltype, part)
            self.__relationships.append(model_rel)
            
            # unmarshall selectively
            if reltype == RT_OFFICEDOCUMENT:
                self.__presentation = part
            # elif reltype == RT_COREPROPS:
            #     self.__coreprops = part
            # elif reltype == RT_EXTENDEDPROPS:  # /docProps/app.xml
            #     self.__extendedprops = part
            # elif reltype == RT_THUMBNAIL:
            #     self.__thumbnail = part
            
    
    
    
    
    def __loadwalk(pkgrels, part_dict)
        for pkgrel in pkgrels:
            # construct target part
            part = Part(reltype, content_type)
            pass
    
    def __unmarshalwalk(rels, visited_parts):
        pass
    

