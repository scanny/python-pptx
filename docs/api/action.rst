.. _action_api:

Click Action-related Objects
============================

The following classes represent click and hover mouse actions, typically
a hyperlink. Other actions such as navigating to another slide in the
presentation or running a macro are also possible.


|ActionSetting| objects
-----------------------

.. autoclass:: pptx.action.ActionSetting()
   :members:
   :inherited-members:
   :undoc-members:


|Hyperlink| objects
-----------------------

.. autoclass:: pptx.action.Hyperlink()
   :members:
   :inherited-members:
   :undoc-members:


|Sound| objects
---------------

A |Sound| object represents a WAV audio clip embedded on a click or hover
action (the ``a:snd`` child of ``a:hlinkClick`` / ``a:hlinkHover``). Use
:meth:`ActionSetting.set_sound` to attach one and
:meth:`ActionSetting.remove_sound` to delete it.

.. autoclass:: pptx.action.Sound()
   :members:
   :inherited-members:
   :undoc-members:


|Audio| objects
---------------

An |Audio| value object carries an audio bytestream (typically a WAV
file) between user code and the package. It is accepted by
:meth:`ActionSetting.set_sound` for callers that want to control the
``name`` shown for a sound or re-use the same audio across multiple
shapes without re-reading it from disk.

.. autoclass:: pptx.media.Audio()
   :members:
   :undoc-members:
