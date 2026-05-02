.. _text_api:

Text-related objects
====================


.. currentmodule:: pptx.text.text


|TextFrame| objects
--------------------

.. autoclass:: TextFrame()
   :members:
   :member-order: bysource
   :undoc-members:


|Font| objects
--------------

The |Font| object is encountered as a property of |_Run|, |_Paragraph|, and in
future other presentation text objects.

.. autoclass:: Font()
   :members:
   :member-order: bysource
   :undoc-members:


|_Paragraph| objects
--------------------

.. autoclass:: _Paragraph()
   :members:
   :member-order: bysource
   :undoc-members:


|_Run| objects
--------------

.. autoclass:: _Run()
   :members:
   :member-order: bysource
   :undoc-members:


|_Field| objects
----------------

A ``_Field`` object is returned by :meth:`._Paragraph.add_field`. It corresponds to an
``<a:fld>`` element — an auto-refresh text field for the current slide number, the
current date/time, or the current footer — and exposes the field type plus the
placeholder-display text.

.. autoclass:: _Field()
   :members:
   :member-order: bysource
   :undoc-members:
