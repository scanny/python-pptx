
Understanding Shapes
====================

Pretty much anything on a slide is a shape; the only thing I can think of that
can appear on a slide that's not a shape is a slide background. There are
between six and ten different types of shape, depending how you count. I'll
explain some of the general shape concepts you'll need to make sense of how to
work with them and then we'll jump right into working with the specific types.

Technically there are six and only six different types of shapes that can be
placed on a slide:

auto shape
   This is a regular shape, like a rectangle, an ellipse, or a block arrow.
   They come in a large variety of preset shapes, in the neighborhood of 180
   different ones. An auto shape can have a fill and an outline, and can
   contain text. Some auto shapes have adjustments, the little yellow diamonds
   you can drag to adjust how round the corners of a rounded rectangle are for
   example. A text box is also an autoshape, a rectangular one, just by default
   without a fill and without an outline.

picture
   A raster image, like a photograph or clip art is referred to as a *picture*
   in PowerPoint. It's its own kind of shape with different behaviors than an
   autoshape. Note that an auto shape can have a picture fill, in which an
   image "shows through" as the background of the shape instead of a fill color
   or gradient. That's a different thing. But cool.

graphic frame
   This is the technical name for the container that holds a table, a chart,
   a smart art diagram, or media clip. You can't add one of these by itself,
   it just shows up in the file when you add a graphical object. You probably
   won't need to know anything more about these.

group shape
   In PowerPoint, a set of shapes can be *grouped*, allowing them to be
   selected, moved, resized, and even filled as a unit. When you group a set of
   shapes a group shape gets created to contain those member shapes. You can't
   actually see these except by their bounding box when the group is
   selected.

line/connector
   Lines are different from auto shapes because, well, they're linear. Some
   lines can be connected to other shapes and stay connected when the other
   shape is moved. These aren't supported yet either so I don't know much more
   about them. I'd better get to these soon though, they seem like they'd be
   very handy.

content part
   I actually have only the vaguest notion of what these are. It has something
   to do with embedding "foreign" XML like SVG in with the presentation. I'm
   pretty sure PowerPoint itself doesn't do anything with these. My strategy
   is to ignore them. Working good so far.

As for real-life shapes, there are these nine types:

* shape shapes -- auto shapes with fill and an outline
* text boxes -- auto shapes with no fill and no outline
* placeholders -- auto shapes that can appear on a slide layout or master and
  be inherited on slides that use that layout, allowing content to be added
  that takes on the formatting of the placeholder
* line/connector -- as described above
* picture -- as described above
* table -- that row and column thing
* chart -- pie chart, line chart, etc.
* smart art -- not supported yet, although preserved if present
* media clip -- video or audio


Accessing the shapes on a slide
-------------------------------

Each slide has a *shape tree* that holds its shapes. It's called a tree because
it's hierarchical in the general case; a node in the shape tree can be a group
shape which itself can contain shapes and has the same semantics as the shape
tree. For most purposes the shape tree has list semantics. You gain access to
it like so::

    shapes = slide.shapes

We'll see a lot more of the shape tree in the next few sections.


Up next ...
-----------

Okay. That should be enough noodle work to get started. Let's move on to
working with AutoShapes.
