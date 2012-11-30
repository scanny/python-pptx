:class:`Presentation` objects
------------------------

A presentation is a programmer-friendly interface to an Open XML PowerPoint
presentation. ... hides the complexity of the XML and packaging ...
operations correspond roughly to actions on the PowerPoint user interface,
e.g. add slide, set font size, etc.

.. class:: pptx.Presentation(template=None)

   Return an instance of an Open XML PresentationML package. ... knows how
   to parse a Presentation object to form the package.

   *template* is a path to a directory containing the presentation template to
   be used to format the presentation ... if *template* is None, the default
   template is used.

   .. method:: addslide(filepath)

      Write the package zip file to *filepath*. If *filepath* does not end in
      ".pptx", that extension will be added.

   .. method:: write(filepath)

      Write the package zip file to *filepath*. If *filepath* does not end in
      ".pptx", that extension will be added.

* ... presentation instantiates a Package when it needs to save (or later perhaps
  read) a package ...


:class:`SlideCollection` objects
-----------------------

... A collection of all the Slide objects in the specified presentation.

.. class:: pptx.Slides(presentation)

   ... not typically instantiated except internally by :class:`Presentation`.

   ... inherits from :class:`list` (maybe). If so has all those methods except
   for those overridden.

   .. method:: addslide(filepath)

      Creates a new slide, adds it to the :class:`Slides` collection, and
      returns the slide.

   .. attribute:: presentation

      The Presentation instance this slide collection belongs to.


:class:`Slide` objects
-----------------------

... a PowerPoint slide.

.. class:: pptx.Slide([slidelayoutidx=None])

   ... inherits from Part (I think) ... might have multiple inheritance, perhaps Part
   and maybe one or two others

   ... not typically instantiated except internally by :class:`Slides`.

   .. attribute:: shapes

      A collection of all the :class:`Shape` instances on this slide.


:class:`Package` objects
------------------------

A package is ... .

.. class:: pptx.packaging.Package(presentation)

   Return an instance of an Open XML PresentationML package. ... knows how
   to parse a Presentation object to form the package.

   .. method:: write(filepath)

      Write the package zip (.pptx) file to *filepath*.



:class:`PackageItem` objects
------------------------

Package items are the top-level components of a package. In general, they
correspond to a file in the package directory hierarchy, such as
/ppt/presentation.xml.

A package item may be a part (e.g. slide1.xml), a relationships-item (e.g.
slide1.xml.rels), or a Content-types item (e.g. [Content_Types].xml). The
corresponding classes inherit from :class:`PackageItem`.

.. class:: pptx.packaging.PackageItem(package)

   Return a new instance of the :class:`PackageItem` class.

   .. method:: zipwrite(zipfile)

      Write rendered package item file to zip archive using either
      :method:`ZipFile.writestr` (for generated text files such as
      presentation.xml) or :method:`ZipFile.write` (for binary resource files
      such as image1.png)

   .. attribute:: package

      The package instance this package item belongs to.


:class:`Part` objects
---------------------

A part is a particular type of Open XML package item. A part is defined in
ECMA-376, Part 1, Section 8.1 as a component of an Open XML document. Other
items that are not considered a part are also contained in the package, most
prominently including relationship ZIP items, XML files that describe
relationships between parts, such as the slide parts that are included in the
presentation and therefore related to the presentation part.

:class:`Part` inherits from :class:`PackageItem`.

NOTE: Might want to have a Part factory so calling classes don't need to understand the implementation details of the various specific part types.

.. class:: pptx.packaging.Part(package)

   Return a new instance of the :class:`Part` class.

   .. attribute:: relpath

      The path for this part, relative to the package root, e.g.
      /ppt/slides/slide1.xml.

   .. attribute:: type

      The name of the type of this part, e.g. 'app', 'presentation',
      'slideLayout', etc. Corresponds to the standard naming for the part type
      files and relationship schema element.


:class:`RelationshipItem` objects
---------------------------------

:class:`RelationshipItem` inherits from :class:`PackageItem`.

Relationship items specify the relationships between parts of the package,
although they are not themselves a part. All relationship items are XML
documents having a filename with the extension '.rels' located in a directory
named '_rels' located in the same directory as the part. The package
relationship item has the package-relative path '/_rels/.rels'. Part
relationship items have the same filename as the part whose relationships they
describe, with the '.rels' extension appended as a suffix. For example, the
relationship item for a part named 'slide1.xml' would have the
package-relative path '/ppt/slides/_rels/slide1.xml.rels'


:class:`ContentTypesItem` objects
---------------------------------

:class:`ContentTypesItem` inherits from :class:`PackageItem`.

The Content Types package item appears in every package exactly once with the
name '[Content_Types].xml'. Its purpose is to specify the content types (MIME
or MIME-like types) of each of the other files in the package so they can be
processed appropriately.

There need only be one :class:`ContentTypesItem` instance for each package and
this class would not normally need to be either instantiated or directly
accessed by user program code. It it instantiated and called internally by
:class:`Package` for all common purposes, but may be interesting to call
directly for testing or learning purposes.


:class:`PartType` objects
---------------------------------

There are tens of different types of :class:`Part`, including presentation,
slideLayout, slide, handoutMaster, image, etc. :class:`PartType` objects are
used to store the various metadata associated with a part type, including the
filename pattern, directory the part(s) are located in, and what relationships
are possible to and from that part type.

.. class:: pptx.packaging.PartType(keyname)

   Return an instance of the :class:`PartType` class corresponding to the
   *keyname* provided. The *keyname* of a part type is with one or two
   exceptions the root portion of the filename for that part type, for example
   'notesSlide', 'slideLayout', 'slide', and 'image'.

   .. attribute:: keyname

      The name of the type of this part, e.g. 'app', 'presentation',
      'slideLayout', etc. Corresponds to the standard naming for the part type
      files and relationship schema element.

   .. attribute:: dirpath

      Package-relative path to the directory parts of this type are stored in.
      *dirpath* always begin with a slash, representing the "root" of the
      package.

   .. attribute:: filenametemplate

      String containing a template for the root portion of the filename for
      parts of this type. The extension for the filename is determined
      separately, sometimes by the part itself. *filenametemplate* is either
      the name itself e.g. 'presentation.xml' (for parts that appear at most
      once in a package) or a printf-style template with a single decimal
      substitution element for types whose parts can appear more than once in
      a package, e.g. 'slide%d.xml'.


