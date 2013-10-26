####
Font
####

:Updated:  2013-06-22
:Author:   Steve Canny
:Status:   **WORKING DRAFT**


Introduction
============

The ``Font`` object is a descriptively named wrapper for run properties. The
name was chosen to correspond with the `Font object in the MS API`_.

Font properties from MS API
===========================

The following properties from the MS API are candidates for inclusion in |pp|.

Implemented
-----------

Bold
   Determines whether the character format is bold. Read/write.

Size
   Returns or sets the character size, in points. Read/write.


Backlog
-------

Caps
   Gets or sets a value specifying that the text should be capitalized.
   Read/write. Corresponds to the ``cap`` attribute of the |rPr| element.
   May be set to MsoTextCaps enumeration values msoNoCaps, msoSmallCaps, or
   msoAllCaps. Return value may also be msoCapsMixed when more than one run
   is involved.

Color
   Returns a ColorFormat object that represents the color for the specified
   characters. Read-only. Pretty sure this indicates the need for a ``Fill``
   class, one type of which is ``SolidFill``.

Italic
   Determines whether the text is italic. Read/write. Corresponds to the
   ``i`` attribute of the |rPr| element. True, False, or None. When set to
   None, its value is inherited.

Name
   Gets or sets a value specifying the font to use for a selection. Read/write.
   Corresponds to the ``typeface`` attribute of the ``<a:latin>`` child
   element of |rPr|.

StrikeThrough
   Gets or sets a value specifying the text should be rendered in
   a strikethrough appearance. Read/write. Corresponds to the ``strike``
   attribute of the |rPr| element.

Subscript
   Gets or sets a value specifying that the selected text should be displayed
   as subscript. Read/write. Implemented using the ``baseline`` attribute of
   the |rPr| element, setting it to -25000 (-25%).

Superscript
   Gets or sets a value specifying that the selected text should be displayed
   as superscript. Read/write. Implemented using the ``baseline`` attribute
   of the |rPr| element, setting it to 30000 (30%)

UnderlineStyle
   Gets or sets a value specifying the underline style for the selected text.
   Read/write. One of 18 possible values in the MsoTextUnderlineType
   enumeration.


Someday, maybe
--------------

Allcaps
   True if the font is formatted as all capital letters. Read/write.

AutorotateNumbers
   Gets or sets a value that specifies whether the numbers in a numbered list
   should be rotated when the text is rotated. Read/write.

BaselineOffset
   Returns or sets the baseline offset for the specified superscript or
   subscript characters. Read/write. Corresponds to the ``baseline`` attribute
   of the |rPr| element.

Creator
   Gets a value indicating the application the object was created in.
   Read-only.

DoubleStrikeThrough
   True if the specified font is formatted as double strikethrough text.
   Read/write.

Emboss
   Determines whether the character format is embossed. Read/write.

Equalize
   Gets or sets a value specifying whether the text for a selection should be
   spaced equal distances apart. Read/write.

Fill
   Gets the formatting properties for the font of the specified text. Read-only

Glow
   Gets a value indicating whether the font is displayed as a glow effect.
   Read-only.

Highlight
   Gets a value indicating whether the font is displayed as highlighted.
   Read-only.

Kerning
   Gets or sets a value specifying the amount of spacing between text
   characters. Read/write.

Line
   Gets a value specifiying the format of a line. Read-only.

NameAscii
   Returns or sets the font used for ASCII characters (characters with
   character set numbers within the range of 0 to 127). Read/write.

NameComplexScript
   Returns or sets the complex script font name. Used for mixed language text.
   Read/write. Corresponds to the ``<a:cs>`` child element of the |rPr|
   element.

NameFarEast
   Returns or sets the Asian font name. Read/write.

NameOther
   Returns or sets the font used for characters whose character set numbers are
   greater than 127. Read/write.

Reflection
   Gets a value specifying the type of reflection format for the selection of
   text. Read-only.

Shadow
   Gets the value specifying the type of shadow effect for the selection of
   text. Read-only.

Smallcaps
   Gets or sets a value specifying whether small caps should be used with the
   slection of text. Small caps are the same height as the lowercase letters in
   a slection of text. Read/write.

SoftEdgeFormat
   Gets or sets a value specifying the type of soft edge effect used in
   a selection of text. Read/write.

Spacing
   Gets or sets a value specifying the spacing between characters in
   a selection of text. Read/write.

Strike
   Gets or sets a value specifying the strike format used for a selection of
   text, as in stiking the same character multiple times to make it darker.
   MsoTextStrike enumeration. Read/write.

UnderlineColor
   Gets a value specifying the color of the underline for the selected text.
   Read-only.

WordArtformat
   Gets or sets a value specifying the text effect for the selected text.
   Read/write.


Mapping UI to API to XML Schema
===============================

Bold
   UI: select range of text and press Cmd-B or press Bold icon. API: Set
   font.bold to True. Schema: ``b`` attribute of ``<a:rPr>`` element.

BaselineOffset
   UI: Format Text > Font > Offset (%). API: Font.BaselineOffset (float between
   -1 and 1, representing percentage). Schema: ``baseline`` attribute of
   ``<a:rPr>`` element (ST_Percentage).


.. |rPr| replace:: ``<a:rPr>``

.. _`font object in the MS API`:
.. _`MSDN Font2 Object API page`:
   http://msdn.microsoft.com/en-us/library/office/ff863038(v=office.14).aspx
