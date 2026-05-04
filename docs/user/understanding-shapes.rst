
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


Shape z-order
-------------

The order shapes appear in a slide's shape tree is also the order they are
drawn; the first shape in the tree is drawn first (and is therefore the
*backmost* shape) and the last shape is drawn last (and is therefore the
*frontmost* shape). This front-to-back stacking is known as the *z-order*.

Newly-added shapes are always appended to the end of the shape tree, so they
appear in front of every previously-added shape. Four methods on every shape
rearrange the z-order within the owning shape tree (the slide's shape tree
for a top-level shape, or the enclosing group shape's tree for a shape inside
a group)::

    shape.bring_to_front()   # -- becomes the frontmost shape --
    shape.send_to_back()     # -- becomes the backmost shape --
    shape.bring_forward()    # -- moves one position forward --
    shape.send_backward()    # -- moves one position backward --

Each method is a no-op when the shape is already at the corresponding
extreme, and none of them moves a shape in or out of a group shape.

A read-only :attr:`~.BaseShape.zorder_index` property reports a shape's
zero-based position within its parent shape-tree (0 is backmost)::

    shape.zorder_index   # -> 2, for example


Effective geometry of shapes inside groups
------------------------------------------

``shape.left``, ``shape.top``, ``shape.width`` and ``shape.height`` always
return the raw values stored in the shape's own ``a:xfrm`` element. For a
top-level shape those are already slide-relative, but for a shape nested
inside a :class:`.GroupShape` the raw values are expressed in the enclosing
group's *child* coordinate system (``a:chOff`` / ``a:chExt``). When the
group has been resized in PowerPoint the group's ``a:ext`` is smaller (or
larger) than its ``a:chExt`` and the child's raw numbers no longer match
the position and size it actually renders at on the slide.

Four read-only companion properties on every shape return the composited,
slide-relative geometry with the enclosing group transform(s) applied::

    shape.effective_left     # -> Length, slide-relative x in EMU
    shape.effective_top      # -> Length, slide-relative y in EMU
    shape.effective_width    # -> Length, rendered width in EMU
    shape.effective_height   # -> Length, rendered height in EMU

For a top-level shape these return the same values as the raw properties.
For a shape inside one or more groups, each group's
``a:chOff``/``a:chExt`` → ``a:off``/``a:ext`` linear transform is applied
in turn up to the slide root. The raw ``left``/``top``/``width``/``height``
properties are unchanged and remain read/write for backwards compatibility.


Ungrouping a group shape
------------------------

A :class:`.GroupShape` can be dissolved in place with its
:meth:`~.GroupShape.ungroup` method. Each direct child is hoisted out of
the enclosing ``p:grpSp`` and onto the slide's top-level shape tree at its
slide-relative effective rectangle, then the now-empty group is removed::

    grp = slide.shapes[2]   # -- a GroupShape containing two autoshapes --
    freed = grp.ungroup()   # -- list of BaseShape freed from the group --

The returned list contains each freed shape in the same z-order it had
inside the group (first returned is backmost within the group). Every
freed shape becomes a top-level sibling on the slide and preserves the
exact slide rectangle at which it rendered before the ungroup.

Nested groups are handled transparently -- when the group being
dissolved is itself nested inside one or more enclosing groups, the
cumulative group-transform cascade (see
:attr:`~.BaseShape.effective_left`) is composited into each child's
coordinates so the children still land at their slide-relative
rectangles. If a freed child is itself a :class:`.GroupShape` (i.e. the
group you ungrouped had a nested sub-group), that sub-group is hoisted
whole: its own ``a:off``/``a:ext`` are rewritten to its slide rectangle
while its internal ``a:chOff``/``a:chExt`` are preserved so its
descendants continue to render in their original slide positions. Call
:meth:`~.GroupShape.ungroup` on the returned sub-group to flatten
further.

After the call returns the dissolved :class:`.GroupShape` refers to an
element that has been removed from the shape tree; do not use the
instance further. See issue #730.


Flat traversal (Selection Pane-equivalent)
------------------------------------------

``slide.shapes`` iterates only the *top-level* shapes on a slide — group
shapes appear as a single entry and their members are hidden. PowerPoint's
Selection Pane, in contrast, lists every shape on the slide including the
children of every group, at every depth. When you need that same flat
listing (for a selection dialog, an audit report, or simply a loop that
touches every shape regardless of nesting) use
:attr:`Slide.shape_tree_flat` — or the underlying
:meth:`SlideShapes.descendants` iterator which also works on layouts,
masters, notes slides, and individual groups::

    for shape in slide.shape_tree_flat:
        print(shape.name, shape.shape_type)

Behaviour notes:

* Shapes are yielded in document (z-order) sequence: a top-level sibling
  preceding a group in the shape tree comes first; group members follow
  in the order they were authored.
* :class:`~.GroupShape` containers are yielded *before* their children,
  matching PowerPoint's Selection Pane display where the group label
  sits above its indented contents. This lets callers see the group
  alongside its members in one pass.
* The iterator walks into nested groups recursively, so a group-inside-
  a-group still produces every leaf shape.

The lookup helpers on :class:`~.SlideShapes` take an optional
``include_descendants=True`` keyword that flips their search from
top-level-only to the same flat traversal::

    # find a shape authored inside a group by its PowerPoint @name
    shape = slide.shapes.get_by_name("Logo", include_descendants=True)

    # every shape named "TODO" anywhere on the slide, including groups
    todos = slide.shapes.find_all_by_name("TODO", include_descendants=True)

Both return the same ordering as :meth:`~.SlideShapes.descendants`
(a matched group appears before any match inside it).

See issue #532.


Removing every shape from a slide
---------------------------------

When you need to wipe a slide's shape tree -- for example when templating
a deck where every slide starts empty but inherits its layout -- use
:meth:`Slide.clear_shapes` (or the underlying
:meth:`SlideShapes.clear`). The default call preserves placeholders so
the slide continues to inherit titles, content regions, and any other
layout-driven content regions; every non-placeholder shape (auto
shapes, pictures, charts, tables, connectors, group shapes and their
contents) is removed::

    slide.clear_shapes()                              # keep placeholders
    slide.clear_shapes(preserve_placeholders=False)   # wipe everything

Per-shape side effects run in the normal way: a removed picture drops
its image relationship, a removed chart drops its embedded chart part,
a removed group is detached together with every shape inside it. Both
methods return ``None`` to match :meth:`list.clear`.

See issue #96.


Accessibility -- shape alt-text and title
-----------------------------------------

Every shape on a slide can carry two accessibility strings that screen
readers and other assistive technology read aloud in place of the shape's
visual content: a short **title** and a longer **alt-text** description.
PowerPoint exposes these in its *Alt Text* pane (``Review > Check
Accessibility > Alt Text``) and stores them as the ``title`` and ``descr``
attributes of the shape's ``cNvPr`` element.

Both attributes are read/write on every shape (AutoShape, Picture, GraphicFrame,
GroupShape, and Connector)::

    shape.alt_text   # -> "bar chart: 2026 quarterly revenue"
    shape.title      # -> "Q revenue chart"

    shape.alt_text = "a detailed description of what the screen reader should say"
    shape.title = "short title"

Both properties default to the empty string when PowerPoint has not written
the attribute. Assigning the empty string clears the attribute so the shape
matches PowerPoint's no-alt-text default.

See issue #508.


Up next ...
-----------

Okay. That should be enough noodle work to get started. Let's move on to
working with AutoShapes.
