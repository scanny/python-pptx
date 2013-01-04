=============================
:mod:`presentation` internals
=============================

For now, this page documents the Relationship-related protocol, others will be
added as development proceeds.


Relationship-related protocol
=============================

Relationship management services are provided by a collaboration between
|_RelationshipCollection| and |_Relationship| objects.
|_RelationshipCollection| maintains an immutable ordered sequence of
|_Relationship| objects corresponding to the relationships of a part or
package (*see* :doc:`../resources/about_relationships` *for more on
relationships*).

The following code sample illustrates the protocol::

    # construction
    relationships = _RelationshipCollection()
    
    # set sequencing by type with auto-renumbering
    relationships._reltype_ordering = (RT_SLIDE, RT_THEME, ...)
    
    # load relationship from package (unique rId known)
    rel = _Relationship(rId, reltype, target_part)
    relationships._additem(rel)
    
    # add new relationship (unique rId not known)
    rId = relationships._next_rId
    rel = _Relationship(rId, ...)
    relationships._additem(rel)

    # retrieve relationships of a specified type
    slide_rels = relationships.rels_of_reltype(RT_SLIDE)


|_RelationshipCollection| objects
----------------------------------------

.. currentmodule:: pptx.presentation

.. autoclass:: pptx.presentation._RelationshipCollection
   :show-inheritance:

   Inherits :meth:`__contains__`, :meth:`__iter__`, :meth:`__len__`,
   :meth:`__getitem__`, and :meth:`index` from :class:`Collection`.

   .. :private-members:
   .. :special-members:

   .. automethod:: _additem

   .. autoattribute:: _next_rId

   .. autoattribute:: _reltype_ordering

   .. automethod:: rels_of_reltype


|_Relationship| objects
-----------------------

.. autoclass:: pptx.presentation._Relationship
   :members:
   :member-order: bysource
   :undoc-members:

   .. autoattribute:: _rId

   .. autoattribute:: _reltype

   .. autoattribute:: _target

   .. autoattribute:: _num



.. |_Relationship| replace:: :class:`_Relationship`

.. |_RelationshipCollection| replace:: :class:`_RelationshipCollection`

