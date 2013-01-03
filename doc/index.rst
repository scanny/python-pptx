.. python-pptx documentation master file, created by
   sphinx-quickstart on Thu Nov 29 13:59:35 2012.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to python-pptx's documentation!
=======================================

Contents:

.. toctree::
   :maxdepth: 2

   modules/index
   resources/index

.. ... inheritance-diagram:: pptx.packaging

.. ... inheritance-diagram:: pptx.presentation

.. .. graphviz::
.. 
..    digraph hierarchy_of_D {
.. 
..    node [color=Blue,fontcolor=Black,font=Courier]
.. 
..     B -> D
..     C -> D
.. 
..     {rank=same; B C }
.. 
..     Part -> PresentationPart
.. 
..     Part -> CollectionPart
.. 
..     packaging -> Part
.. 
..    }


   .. digraph hierarchy_of_D {
   .. 
   .. node [color=Blue,fontcolor=Black,font=Courier]
   .. 
   ..  Part -> PresentationPart
   ..  Part -> CollectionPart
   .. 
   ..  {rank=same; PresentationPart CollectionPart }
   .. 
   ..  PartCollection -> ImageParts
   ..  PartCollection -> SlideLayoutParts
   .. 
   ..  pptx.packaging -> Part
   ..  pptx.packaging -> PartCollection
   .. 
   .. }


   .. digraph foo {
   ..    "bar" -> "baz";
   .. }


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

