.. _install:

Installing
==========

|pp| is published on PyPI. Install it with ``pip``::

    pip install python-pptx

Or, with `uv <https://docs.astral.sh/uv/>`_::

    uv pip install python-pptx

To add |pp| as a project dependency with uv::

    uv add python-pptx

Requirements
------------

* Python 3.8 or later
* `lxml <https://pypi.org/project/lxml/>`_
* `Pillow <https://pypi.org/project/Pillow/>`_
* `XlsxWriter <https://pypi.org/project/XlsxWriter/>`_ (used by the charting features)
* `typing_extensions <https://pypi.org/project/typing-extensions/>`_

``pip`` (and ``uv``) resolve and install these dependencies automatically.

Optional dependencies
---------------------

* `python-ooxml-crypto <https://pypi.org/project/python-ooxml-crypto/>`_ — required only
  to open or save password-protected ``.pptx`` files. It is imported lazily the
  first time a ``password`` argument is passed to :func:`pptx.Presentation` or
  :meth:`Presentation.save`. Install it with::

      pip install python-ooxml-crypto

  If you never use the ``password`` argument, |pp| has no extra runtime
  dependency.

Related projects
----------------

|pp| is part of a family of pure-Python libraries for reading and writing
Office Open XML files:

* `python-docx <https://python-docx.readthedocs.io/en/latest/user/install.html>`_ — Microsoft Word (``.docx``) documents
* `python-xlsx <https://python-xlsx.readthedocs.io/en/latest/user/install.html>`_ — Microsoft Excel (``.xlsx``) workbooks
