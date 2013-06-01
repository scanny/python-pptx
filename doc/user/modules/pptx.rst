###################
:mod:`pptx` Package
###################

Manipulate Open XML PowerPoint® files.

The most commonly used references in ``python-pptx`` can be imported directly
from the package::

    from pptx import Presentation

The components that constitute a presentation, e.g. slides, shapes, etc., are
lodged in a graph of which the |Presentation| object is the root. All existing
presentation components are referenced by traversing the graph and new objects
are added to the graph by calling a method on that object's containing object.
Consequently, the only presentation object that is constructed directly is
|Presentation|.

Example::

   # load a presentation
   prs = Presentation(path_to_pptx_file)

   # get reference to first shape in first slide
   sp = prs.slides[0].shapes[0]
   
   # add a picture shape to slide
   pic = sld.shapes.add_picture(path, x, y, cx, cy)


|Presentation| objects
======================

The |Presentation| class is the only reference that must be imported to work
with presentation files. Typical use interacts with many other classes, but
there is no need to construct them as they are accessed through a property of
their parent object.

.. autoclass:: pptx.Presentation
   :members:
   :member-order: bysource
   :undoc-members:


.. currentmodule:: pptx.presentation


|_SlideCollection| objects
==========================

The |_SlideCollection| object is typically encountered as the
:attr:`~pptx.Presentation.slides` member of |Presentation|. It is not intended
to be constructed directly.

.. autoclass:: pptx.presentation._SlideCollection
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:


|_ShapeCollection| objects
==========================

The |_ShapeCollection| object is typically encountered as the
:attr:`~BaseSlide.shapes` member of |_Slide|. It is not intended to be
constructed directly.

.. autoclass:: pptx.presentation._ShapeCollection
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:


``Shape`` objects
=================

The following properties and methods are common to all shapes.

.. autoclass:: pptx.shapes._BaseShape
   :members:
   :member-order: bysource
   :undoc-members:


The following properties and methods are defined for auto shapes and text boxes.

.. autoclass:: pptx.shapes._Shape
   :members:
   :member-order: bysource
   :undoc-members:


|_AdjustmentCollection| objects
===============================

An |_AdjustmentCollection| object reference is accessed using the
``_Shape.adjustments`` property (read-only).

.. autoclass:: pptx.shapes._AdjustmentCollection
   :members:
   :member-order: bysource
   :undoc-members:


|_Table| objects
================

A |_Table| object is added to a slide using the :meth:`add_table` method on
|_ShapeCollection|.

.. autoclass:: pptx.shapes._Table
   :members:
   :member-order: bysource
   :undoc-members:


|_Column| objects
=================

.. autoclass:: pptx.shapes._Column()
   :members:
   :member-order: bysource
   :undoc-members:


|_Row| objects
==============

.. autoclass:: pptx.shapes._Row()
   :members:
   :member-order: bysource
   :undoc-members:


|_Cell| objects
===============

A |_Cell| object represents a single table cell at a particular row/column
location in the table. |_Cell| objects are not constructed directly. A
reference to a |_Cell| object is obtained using the :meth:`Table.cell` method,
specifying the cell's row/column location.

.. autoclass:: pptx.shapes._Cell
   :members:
   :member-order: bysource
   :undoc-members:


|_TextFrame| objects
====================

.. autoclass:: pptx.shapes._TextFrame()
   :members:
   :member-order: bysource
   :undoc-members:


|_Font| objects
===============

The |_Font| object is encountered as a property of |_Run|, |_Paragraph|, and in
future other presentation text objects.

.. autoclass:: pptx.shapes._Font()
   :members:
   :member-order: bysource
   :undoc-members:


|_Paragraph| objects
====================

.. autoclass:: pptx.shapes._Paragraph()
   :members:
   :member-order: bysource
   :undoc-members:


|_Run| objects
==============

.. autoclass:: pptx.shapes._Run()
   :members:
   :member-order: bysource
   :undoc-members:


:mod:`presentation` Module
==========================

The remaining API classes of the :mod:`presentation` module are described
here.

.. automodule:: pptx.presentation
   :members: _BaseSlide, _SlideMaster, _SlideLayout, _Slide
   :member-order: bysource
   :show-inheritance:
   :undoc-members:

