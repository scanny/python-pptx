=============
Odds and Ends
=============

Bits and pieces here without a more specific home


Class Hierarchy
===============

::

   pptx
   |
   +--.packaging
      |
      +--_PartCollection
      |  |
      |  +--ImageParts
      |  +--SlideLayoutParts
      |  +--SlideMasterParts
      |  +--SlideParts
      |  +--ThemeParts
      |
      +--Part
         |
         +--CollectionPart
         |  |
         |  +--ImagePart
         |  +--SlideLayoutPart
         |  +--SlideMasterPart
         |  +--ThemePart
         |
         +--PresentationPart
         +--PresPropsPart
         +--SlidePart
         +--TableStylesPart
         +--ViewPropsPart


Also try something like this::

   Exception hierarchy
   -------------------

   The class hierarchy for built-in exceptions is:

   .. literalinclude:: ../../Lib/test/exception_hierarchy.txt


Feature adding process:
=======================

* Analysis

  * Use case
  * Aspects of the spec and PowerPoint conventions that bear on the feature

* Implementation strategy

* Test Cases

* Documentation

* Implementation

* Commit


Odds and Ends
=============

* DOC: Reserved Terms, e.g. package, part, item, element, etc. These have
  special meaning because of the terminology of Open XML packages and
  ElementTree XML conventions. Singleton, collection, perhaps tuple.

* DOC: Variable naming conventions. Variables containing XML elements are
  named the same as the tag name of the element they hold. For example:
  sldMaster = etree.parse('sldMaster1.xml')

* DOC: Suitable for server-side document generation, database publishing,
  taking tedium out of building intricate slides that have a distinct pattern,
  etc.

* DOC: Content Types mapping is discussed in ECMA-376-2 10.1.2 *Mapping
  Content Types*

* DOC: Explicit vs. Implicit relationships are discussed in section 9.2 of
  ECMA-376-1. Basically relationships to whole parts are explicit and
  relationships to one of the elements within another part (like a footnote)
  are implicit.

* DOC: [Open XML SDK How To's (How do I...)|
  http://207.46.22.237/en-us/library/cc850828.aspx]

* DOC: [Good hints on where formatting comes from in
  template|http://openxmldeveloper.org/discussions/formats/f/15/t/1301.aspx]

* Set work template slide layouts to standard type names for better layout
  re-mapping on theme/template change
  [http://msdn.microsoft.com/en-us/library/office/cc850846.aspx]

* spTree in cSld is a GroupShape (ShapeGroup?) perhaps Shape and ShapeGroup
  are ... ? not sure who should inherit from the other. GroupShapes can
  include other GroupShapes.

