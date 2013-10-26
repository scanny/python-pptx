========================
About Part Relationships
========================

Understanding Part Relationships
================================

An Open XML package part can have one or more relationships to other parts in
the package. For example, a slide layout has a relationship to the slide
master it inherits properties from. And if the slide layout includes an image,
such as a company logo, it will also have a relationship to the image part
corresponding to that logo's image file.

Conceptually, a part relationship might be stated::

   Part x is related to part y of type z.

While these relationships are reciprocal in a way of thinking (e.g. part y is
also related to part x), in Open XML they are generally defined in one
direction only. For example, a slide may have a relationship to an image it
contains, but an image never has a stated relationships to a slide. One
exception to this is slide masters and slide layouts. A slide layout has a
relationship to the slide master it inherits from at the same time the slide
master has a relationship to each of the slide layouts that inherit from it.

In the example above, part x can be referred to as the *source part* or
*from-part*, part y as the *target part* or *to-part*, and type z as the
*relationship type*. The relationship type essentially ends up being the part
type of the target part (e.g. image, slide master, etc.).

Inbound and outbound relationships
----------------------------------

Each relationship is *outbound* for its source part and *inbound* for its
target part. The *outbound* relationships of a part are those found in its
*part relationship item*, commonly referred to as its "rels file". Only
outbound relationships are recorded in the package. Inbound relationships are
purely abstract, although they can be inferred from outbound relationships if
an application had a need for them.

Not all parts can have outbound relationships. For example, image and
presentation properties parts never have outbound relationships. All parts,
however, participate in at least one inbound relationship. Any part with no
inbound relationships could be removed from the package without consequence,
and probably should be. For example, if upon saving it was noticed that a
particular image had no inbound relationships, that image would be better not
written to the package. *(Makes me wonder whether loading a package by walking
its relationship graph might be a good idea in some instances.)*


Direct and indirect relationship references
-------------------------------------------

Each relationship is recorded as a relationship element in the rels file
belonging to a specific part. (The package relationship item ``/_rels/.rels``
is the only exception, it belongs to the package, not to any part). Each
relationship entry specifies a part-local relationship id (rId) for the
relationship, the relationship type (essentially the part type of the target,
e.g. slide, image), and a path the target part file. The source part is not
explicitly stated in the relationship entry, it is implicitly the part
the .rels file belongs to.

These can be thought of as *external relationship references* in that the rels
file is a separate package item, external to the part itself. However, in
certain cases, a relationship may be referenced within the XML of the
from-part. These can be thought of as *internal relationship references*.

As an example of where an internal relationship reference is required,
consider a slide containing images of three different products. Each picture
has a corresponding <p:pic> element in the slide's shape tree, each of which
must specify the particular image file it displays, distinguishing it from the
other two image files related to the slide.

The picture elements specify which of the related images to display using the
part-local relationship id (rId) matching the required image, 'rId2' in the
example below::

   <p:blipFill>
     <a:blip r:embed="rId2">
       ...
     </a:blip>
   </p:blipFill>

Which is an indirect reference to ``image1.png`` specified as the target of
'rId2' in the slide's rels file::

   <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
   <Relationships xmlns="http://.../relationships">
     <Relationship Id="rId1" Type="http://.../slideLayout" Target=".../slideLayout2.xml"/>
     <Relationship Id="rId2" Type="http://.../image" Target=".../image1.png"/>
   </Relationships>

This indirection makes sense as a way to limit the coupling of presentation
parts to the mechanics of package composition. For example, when the XML for
the slide part in the example above is being composed, the slide can determine
the reference to the image it's displaying without reference outside its own
natural scope. In contrast, determining the eventual location and filename of
that image in any particular package that was saved would require the slide
code to have visibility into the packaging process, which would prevent
packaging being delegated to a separate, black-box module.

Implicit and explicit relationships
-----------------------------------

There is also a distinction between implicit and explicit relationships which
is described in the spec (ECMA-376-1) in section 9.2. I haven't encountered
those yet in the context of PresentationML (the spec uses an example of
footnotes from WordprocessingML), so I do not discuss them here.


Relationship Mechanics
======================

Relationship life-cycle
-----------------------

A representation of a relationship must operate effectively in two distinct
situations, in-package and in-memory. They must also support lifecycle
transitions from in-package to in-memory and from in-memory back to
in-package.

Abstract model
--------------

Each relationship has the following abstract attributes

.. attribute:: source

   The "from-part" of the relationship.

.. attribute:: id

   Source-local unique identifier for this relationship. Each source part's
   relationship ids should be a sequence of consecutive integers starting from
   1.

.. attribute:: target_type

   The content type of the relationship's to-part.

.. attribute:: target

   A direct reference to the relationship's to-part.



Maintaining relationships by dynamic parts (e.g. Slides)
========================================================

How will dynamic parts (like Slide) interact with its relationship list?

? Should it just add items to the relationship list when it creates new things?

? Does it need some sort of lookup capability in order to delete? Or just have a delete relationship method on RelationshipCollection or something like that.

Need to come up with a plausible set of use cases to think about a design.
Right now the only use case is loading a template into a presentation and
saving a presentation.

* Add an image to a slide.

* Change a slide's slide layout

* comment, notesSlide, tag, image, and slideLayout are the only outbound
  relationship types for a slide, although I expect there are some other
  DrawingML bits I haven't accounted for yet.

On reflection I'm thinking there's not too much urgency on noodling this out
too far, the present construction should work fine for now and be able to be
extended without disrupting other code too much.


SCRAP
=====

Rationale for Relationship to be an association class
-----------------------------------------------------

When loaded into memory, each relationship target must be a reference to an
active part object (or at least a part key that can be resolved to a
reference, but why do this lookup multiple times?). This is both because those
relationships can change and also because the package path, while it can be
calculated at runtime, is not guaranteed to be stable (e.g. a new slide can be
inserted between two existing ones) and is not finally resolved until the
presentation is saved.

General description of package relationships items
--------------------------------------------------

* Relationships items specify the relationships between parts of the package,
  although they are not themselves a part.

* All relationship items are XML documents having a filename with the
  extension '.rels' located in a directory named '_rels' located in the same
  directory as the part.

* The package relationship item has the URI '/_rels/.rels'.

* Part relationship items have the same filename as the part whose
  relationships they describe, with the '.rels' extension appended as a
  suffix. For example, the relationship item for a part named
  ``/ppt/slides/slide1.xml`` would have the URI
  ``/ppt/slides/_rels/slide1.xml.rels``.


