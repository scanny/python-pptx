=============================
Parts protocol and interfaces
=============================

... general discussion of Part interfaces and protocol ...

Behavior of loaded parts is different from added parts (but shouldn't be,
indicates a coupling to loading) ...

There are essentially three protocols with partially-overlapping interfaces:

* unmarshaling protocol
* marshaling protocol
* manipulation protocol


:class:`Part` objects
=====================

.. currentmodule:: pptx.presentation

.. autoclass:: pptx.presentation.Part
   :members: __new__
   :undoc-members:


.. autoclass:: pptx.presentation.BasePart
   :members:
   :undoc-members:
   :private-members:


.. autoclass:: pptx.presentation.PartCollection
   :members:
   :undoc-members:
   :private-members:
   :show-inheritance:

   Inherits :meth:`__contains__`, :meth:`__iter__`, :meth:`__len__`,
   :meth:`__getitem__`, and :meth:`index` from :class:`Collection`.



Unmarshaling interface
---------------------------------------

The unmarshaling interface provides a protocol the model side uses to load
a graph of parts from :class:`pptx.packaging.Package`.

The following code sample illustrates the protocol::

    # construct a part
    part = Part(reltype, content_type)
    
    # load it from instance of pptx.packaging.Part
    part._load(pkgpart, part_dict)
    
    # construct a part collection
    parts = PartCollection()
    
    # add a part that was loaded from package
    parts._loadpart(part)
    
    # access part attributes
    partname = part.partname
    content_type = part._content_type
    blob = part._blob
    
    # get ElementTree XML element (XML parts only)
    elm = part._element
    
    # access part relationships
    rels = part._relationships



Marshaling interface
---------------------------------------

The marshaling interface provides a protocol to :mod:`pptx.packaging` that
allows it to gather what it needs to save the document to a .pptx file.

The following code sample illustrates the protocol::

    # access relationships
    rels = part._relationships
    
    # access partname
    partname = part.partname
    
    # access content type
    content_type = part._content_type
    
    # access blob
    blob = part._blob



Manipulation interface
---------------------------------------

The manipulation protocol and interfaces are used as the part graph is
manipulated via the library API::

    # add a part
    part = parts.add_part(part)
    
    # delete a part
    raise NotImplementedError('not implemented yet')
    

Class hierarchy
---------------------------------------

Probably outdated, needs fixing ...

::

   +--BasePart
   |  |
   |  +--Presentation
   |  +--BaseSlide
   |  |  |
   |  |  +--SlideMaster
   |  |  +--SlideLayout
   |  |  +--Slide
   |  |  +--... others later
   |  |
   |  +--... others later
   |
   +--Collection
      |
      +--PartCollection
         |
         +--SlideCollection
         +--... others later

