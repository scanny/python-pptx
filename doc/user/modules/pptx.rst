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
are added to the graph by calling a method on that objects containing object.
Consequently, the only presentation object that is constructed directly is
|Presentation|.

Example::

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


|SlideCollection| objects
=========================

The |SlideCollection| object is typically encountered as the
:attr:`~pptx.Presentation.slides` member of |Presentation|. It is not intended
to be constructed directly.

.. autoclass:: pptx.presentation.SlideCollection
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:


|ShapeCollection| objects
=========================

The |ShapeCollection| object is typically encountered as the
:attr:`~BaseSlide.shapes` member of |Slide|. It is not intended to be
constructed directly.

.. autoclass:: pptx.presentation.ShapeCollection
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:


``Shape`` objects
=================

The following properties and methods are common to all shapes.

.. autoclass:: pptx.shapes.BaseShape
   :members:
   :member-order: bysource
   :undoc-members:


|TextFrame| objects
===================

.. autoclass:: pptx.shapes.TextFrame()
   :members:
   :member-order: bysource
   :undoc-members:


|Font| objects
===============

The |Font| object is encountered as a property of |Run|, |Paragraph|, and in
future other presentation text objects.

.. autoclass:: pptx.shapes._Font()
   :members:
   :member-order: bysource
   :undoc-members:


|Paragraph| objects
===================

.. autoclass:: pptx.shapes.Paragraph()
   :members:
   :member-order: bysource
   :undoc-members:


|Run| objects
=============

.. autoclass:: pptx.shapes.Run()
   :members:
   :member-order: bysource
   :undoc-members:


:mod:`presentation` Module
==========================

The remaining API classes of the :mod:`presentation` module are described
here.

.. automodule:: pptx.presentation
   :members: BaseSlide, SlideMaster, SlideLayout, Slide
   :member-order: bysource
   :show-inheritance:
   :undoc-members:

.. |Presentation| replace:: :class:`~pptx.Presentation`

.. |SlideCollection| replace:: :class:`SlideCollection`

.. |Slide| replace:: :class:`Slide`

.. |ShapeCollection| replace:: :class:`ShapeCollection`

.. |TextFrame| replace:: :class:`TextFrame`

.. |Paragraph| replace:: :class:`Paragraph`

.. |Run| replace:: :class:`Run`

.. |Font| replace:: :class:`_Font`


