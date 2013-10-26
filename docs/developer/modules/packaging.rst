=======================
:mod:`packaging` Module
=======================

Manipulate Open Packaging Convention files

.. automodule:: pptx.packaging

   .. :members: PartTypeSpec
   .. :member-order: bysource
   .. :undoc-members:
   .. :show-inheritance:
   .. :members: PartTypeSpec ContentTypesItem FileSystem BaseFileSystem
                DirectoryFileSystem ZipFileSystem

   A single |Package| instance along with a number of |Part| instances form
   the nodes of the object graph for an Open XML document, and |Relationship|
   instances provide the directed links that connect these nodes into a graph.
   These three classes are discussed first, followed by their supporting
   classes.


|Package| objects
------------------------

Most users of the :mod:`pptx.packaging` module will only need to interact with
the |Package| class. The remaining classes in the module provide specialized
support for |Package|, although one or two contribute to the API via one or
more of their attributes.

.. autoclass:: pptx.packaging.Package
   :members:
   :member-order: bysource
..    :undoc-members:


|Part| objects
---------------------

.. note::
   The :mod:`pptx.presentation` module contains a distinct class that is also
   named |Part|. The two classes are conceptually similar, but have very
   different behavior corresponding to their uses in marshaled and unmarshaled
   packages, respectively.

.. autoclass:: pptx.packaging.Part
   :members:
   :member-order: bysource
   :undoc-members:


|Relationship| objects
-----------------------------

|Relationship| objects associate a *source* with a *target*. The source is
either an instance of |Package| or of |Part|. The target is always an instance
of |Part|. In the *Open Packaging Convention* (OPC), relationships serve to
serialize the direct references to parts that document objects have when
in-memory. Relationships turn out to be a pivotal notion in dealing with OPC
packages because they represent the links in the document part graph.

.. note::
   The :mod:`pptx.presentation` module contains a distinct class that is also
   named |Relationship|. The two classes are conceptually similar, but have
   different behavior corresponding to their use in marshaled and unmarshaled
   packages, respectively.

.. autoclass:: pptx.packaging.Relationship
   :members:
   :member-order: bysource
   :undoc-members:


|PartTypeSpec| objects
-----------------------------

The `ECMA-376 spec`_ defines several characteristics and constant values for
each type of part that may appear in a package. |PartTypeSpec| instances make
that data available as an aid to processing OPC packages.

.. _ECMA-376 spec:
   http://www.ecma-international.org/publications/standards/Ecma-376.htm

.. autoclass:: pptx.packaging.PartTypeSpec
   :members:
   :member-order: bysource
   :undoc-members:

