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


|_Run| objects
--------------

.. autoclass:: _Run()
   :members:
   :member-order: bysource
   :undoc-members:
