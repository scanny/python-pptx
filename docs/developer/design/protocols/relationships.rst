=====================================
Relationships protocol and interfaces
=====================================

... the relationship-related protocol ...


.. _relationship-related-protocol:

Relationship-related protocol
=============================

Relationship management is provided by a collaboration between
|RelationshipCollection| and |Relationship| objects.
|RelationshipCollection| maintains an immutable ordered sequence of
|Relationship| objects corresponding to the relationships of a part or
package (*see* :doc:`../../resources/about_relationships` *for more on
relationships*).

The following code sample illustrates the protocol::

    # construction
    relationships = RelationshipCollection()
    
    # set sequencing by type with auto-renumbering
    relationships._reltype_ordering = (RT_SLIDE, RT_THEME, ...)
    
    # load relationship from package (unique rId known)
    rel = Relationship(rId, reltype, target_part)
    relationships._additem(rel)
    
    # add new relationship (unique rId not known)
    rId = relationships._next_rId
    rel = Relationship(rId, ...)
    relationships._additem(rel)

    # retrieve relationships of a specified type
    slide_rels = relationships.rels_of_reltype(RT_SLIDE)


|RelationshipCollection| objects
----------------------------------------

.. currentmodule:: pptx.opc.rels

.. autoclass:: pptx.opc.rels.RelationshipCollection
   :members: _additem, _next_rId, _reltype_ordering, rels_of_reltype
   :member-order: bysource
   :undoc-members:
   :show-inheritance:

   Inherits :meth:`__contains__`, :meth:`__iter__`, :meth:`__len__`,
   :meth:`__getitem__`, and :meth:`index` from |Collection|.


|Relationship| objects
-----------------------

.. autoclass:: pptx.opc.rels.Relationship
   :members: _rId, _reltype, _target, _num
   :member-order: bysource
   :undoc-members:

