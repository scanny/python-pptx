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


|TextFrameRect| objects
-----------------------

A ``TextFrameRect`` is returned by
:attr:`~pptx.shapes.base.BaseShape.text_frame_rect`. It is a 4-element
namedtuple of |Length| values in EMU describing the slide-relative
rectangle PowerPoint allocates for text within a shape (issue #663).

.. autoclass:: TextFrameRect()
   :members:
   :member-order: bysource
   :undoc-members:


|Font| objects
--------------

The |Font| object is encountered as a property of |_Run|, |_Paragraph|, and in
future other presentation text objects.

A run-properties element (``a:rPr``) exposes three independent typeface slots:
``a:latin`` (Latin script), ``a:ea`` (East-Asian / CJK), and ``a:cs`` (complex
script, e.g. Arabic, Hebrew, Thai). These slots are surfaced on |Font| as
:attr:`Font.name`, :attr:`Font.name_ea`, and :attr:`Font.name_cs`
respectively. Each slot can be set independently; PowerPoint selects the
appropriate slot per-character based on the Unicode range of the text being
rendered.

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


|_BulletFormat| objects
-----------------------

The |_BulletFormat| object is a proxy for the bullet-related child elements of
a paragraph's ``a:pPr`` element. It is obtained as the
:attr:`._Paragraph.bullet` property.

.. autoclass:: _BulletFormat()
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
