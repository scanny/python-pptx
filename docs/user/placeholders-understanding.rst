
Understanding placeholders
==========================

Intuitively, a placeholder is a pre-formatted container into which content
can be placed. By providing pre-set formatting to its content, it places many
of the formatting choices in the hands of the template designer while
allowing the end-user to concentrate on the actual content. This speeds the
presentation development process while encouraging visual consistency in
slides created from the same template.

While their typical end-user behaviors are relatively simple, the structures
that support their operation are more complex. This page is for those who
want to better understand the architecture of the placeholder subsystem and
perhaps be less prone to confusion at its sometimes puzzling behavior. If you
don't care why they work and just want to know how to work with them, you may
want to skip forward to the following page :ref:`placeholders-using`.


A placeholder is a shape
------------------------

Placeholders are an orthogonal category of shape, which is to say multiple
shape types can be placeholders. In particular, the auto shape (`p:sp`
element), picture (`p:pic` element), and graphic frame (`p:graphicFrame`)
shape types can be a placeholder. The group shape (`p:grpSp`), connector
(`p:cxnSp`), and content part (`p:contentPart`) shapes cannot be
a placeholder. A graphic frame placeholder can contain a table, a chart, or
SmartArt.


Placeholder types
-----------------

There are 18 types of placeholder.

Title, Center Title, Subtitle, Body
   These placeholders typically appear on a conventional "word chart"
   containing text only, often organized as a title and a series of bullet
   points. All of these placeholders can accept text only.

Content
   This multi-purpose placeholder is the most commonly used for the body of
   a slide. When unpopulated, it displays 6 buttons to allow insertion of
   a table, a chart, SmartArt, a picture, clip art, or a media clip.

Picture, Clip Art
   These both allow insertion of an image. The insert button on a clip art
   placeholder brings up the clip art gallery rather than an image file
   chooser, but otherwise these behave the same.

Chart, Table, Smart Art
   These three allow the respective type of rich graphical content to be
   inserted.

Media Clip
   Allows a video or sound recording to be inserted.

Date, Footer, Slide Number
   These three appear on most slide masters and slide layouts, but do not
   behave as most users would expect. These also commonly appear on the Notes
   Master and Handout Master.

Header
   Only valid on the Notes Master and Handout Master.

Vertical Body, Vertical Object, Vertical Title
   Used with vertically oriented languages such as Japanese.


Unpopulated vs. populated
-------------------------

A placeholder on a slide can be empty or filled. This is most evident with
a picture placeholder. When unpopulated, a placeholder displays customizable
prompt text. A rich content placeholder will also display one or more content
insertion buttons when empty.

A text-only placeholder enters "populated" mode when the first character of
text is entered and returns to "unpopulated" mode when the last character of
text is removed. A rich-content placeholder enters populated mode when
content such as a picture is inserted and returns to unpopulated mode when
that content is deleted. In order to delete a populated placeholder, the
shape must be deleted *twice*. The first delete removes the content and
restores the placeholder to unpopulated mode. An additional delete will
remove the placeholder itself. A deleted placeholder can be restored by
reapplying the layout.


Placeholders inherit
--------------------

A placeholder appearing on a slide is only part of the overall placeholder
mechanism. Placeholder behavior requires three different categories of
placeholder shape; those that exist on a slide master, those on a slide
layout, and those that ultimately appear on a slide in a presentation.

These three categories of placeholder participate in a property inheritance
hierarchy, either as an inheritor, an inheritee, or both. Placeholder shapes
on masters are inheritees only. Conversely placeholder shapes on slides are
inheritors only. Placeholders on slide layouts are both, a possible inheritor
from a slide master placeholder and an inheritee to placeholders on slides
linked to that layout.

A layout inherits from its master differently than a slide inherits from
its layout. A layout placeholder inherits from the master placeholder sharing
the same type. A slide placeholder inherits from the layout placeholder
having the same `idx` value.

In general, all formatting properties are inherited from the "parent"
placeholder. This includes position and size as well as fill, line, and font.
Any directly applied formatting overrides the corresponding inherited value.
Directly applied formatting can be removed be reapplying the layout.


Glossary
--------

placeholder shape
    A shape on a slide that inherits from a layout placeholder.

layout placeholder
    a shorthand name for the placeholder shape on the slide layout from which
    a particular placeholder on a slide inherits shape properties

master placeholder
    the placeholder shape on the slide master which a layout placeholder
    inherits from, if any.
