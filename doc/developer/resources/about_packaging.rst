========================
About Open XML Packaging
========================

Recent Notes
============

* The content type for **XML** parts (only) is strictly determined by the
  relationship type. Binary media parts may have multiple allowable content
  types.

* Each part type has one and only one relationship type.


About Packages
==============

Open XML PowerPoint files are stored in files with the extension ``.pptx``.
These ``.pptx`` files are zip archives containing a separate file for each
main component of the presentation along with other files which contain
metadata about the overall presentation and relationships between its
components.

The overall collection of files is known as a **package**. Each file within
the package represents a **package item**, commonly although sometimes
ambiguously referred to simply as an **item**. Package items that represent
high-level presentation components, such as slides, slide masters, and themes,
are referred to as **parts**. All parts are package items, but not all
package items are parts.

Package items that are not parts are primarily **relationship items** that
express relationships between one part and another. Examples of relationships
include a slide having a relationship to the slide layout it is based on and a
slide master's relationship to an image such as a logo it displays. There is
one special relationship item, the **package relationship item**
(``/_rels/.rels``) which contains relationships between the package and
certain parts rather than relationships between parts.

The only other package item is the **content types item**
(``/[Content_Types].xml``) which contains mappings of the parts to their
content types (roughly MIME types).


Package Loading Strategies
==========================

Strategy 1
----------

The content type item ([Content_Types].xml actually contains references to all
the main parts of the package. So one approach might be:

1. Have passed in a mapping of content types to target classes that each
   know how to load themselves from the xml and a list of relationships.
   These classes are what OpenXML4J calls an *unmarshaller*. These would
   each look something like::
   
      { 'application/vnd...slide+xml'       : Slide
      , 'application/vnd...slideLayout+xml' : SlideLayout
      , ...
      }

2. Have a ContentType class that can load from a content type item and allow 
   lookups by partname. Have it load the content type item from the package.
   Lookups on it first look for an explicit override, but then fall back to
   defaults. Not sure yet whether it should try to be smart about whether a
   package part that looks like one that should have an override will get
   fed to the xml default or not.

3. Walk the package directory tree, and as each file is encountered:

OR

3. Walk the relationships tree, starting from /_rels/.rels

   * look up its content type in the content type manager
   * look up the unmarshaller for that content type
   * dispatch the part to the specified unmarshaller for loading
   * add the resulting part to the package parts collection
   * if it's a rels file, parse it and associate the relationships with the
     appropriate source part.

   If a content type or unmarshaller is not found, throw an exception and
   exit. Skip the content type item (already processed), and for now skip
   the package relationship item. Infer the corresponding part name from
   part rels item names. I think the walk can be configured so rels items
   are encountered only after their part has been processed.

4. Resolve all part relationship targets to in-memory references.


Principles upheld:
^^^^^^^^^^^^^^^^^^

* If there are any stray items in the package (items not referenced in the
  content type part), they are identified and can be dealt with appropriately
  by throwing an exception if it looks like a package part or just writing it
  to the log if it's an extra file. Can set debug during development to throw
  an exception either way just to give a sense of what might typically be
  found in a package or to give notice that a hand-crafted package has an
  internal inconsistency.

* Conversely, if there are any parts referenced in the content type part that
  are not found in the zip archive, that throws an exception too.


Random thoughts
===============

I'm starting to think that the packaging module could be a useful
general-purpose capability that could be applied to .docx and .xlsx files in
addition to .pptx ones.

Also I'm thinking that a generalized loading strategy that walks the zip
archive directory tree and loads files based on combining what it discovers
there with what it can look up in it's spec tables might be an interesting
approach.

I can think of the following possible ways to identify the type of package,
not sure which one is most reliable or definitive:

* Check the file extension. Kind of thinking this should do the trick 99.99%
  of the time.

* Might not hurt to confirm that by finding an expected directory of /ppt, /word, or /xl.

* A little further on that would be to find /ppt/presentation.xml, /word/document.xml, or /xl/workbook.xml.

* Even further would be to find a relationship in the package relationship item to one of those three and not to any of the others.

* Another confirmation might be finding a content type in [Content_Types].xml of:

  * .../presentationml.presentation.main+xml for /ppt/presentation.xml

  * .../spreadsheetml.sheet.main+xml for /xl/workbook.xml (Note that
    macro-enabled workbooks use a different content type
    'application/vnd.ms-excel.sheet.macroEnabled.main+xml', I believe there
    are variants for PresentationML as well, used for templates and slide
    shows.

  * .../wordprocessingml.document.main+xml for /word/document.xml

It's probably worth consulting ECMA-376 Part 2 to see if there are any hard
rules that might help determine what a definitive test would be.
