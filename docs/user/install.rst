.. _install:

Installing
==========

|pp| is hosted on PyPI, so installing with `pip` is simple::

    pip install python-pptx

|pp| depends on the ``lxml`` package and ``Pillow``, the modern version of
the Python Imaging Library (``PIL``). The charting features depend on
``XlsxWriter``. Both ``pip`` and ``easy_install`` will take care of
satisfying these dependencies for you, but if you use the ``setup.py``
installation method you will need to install the dependencies yourself.

Currently |pp| requires Python 2.7 or 3.3 or later. The tests are run against 2.7 and
3.8 on Travis CI.

Dependencies
------------

* Python 2.6, 2.7, 3.3 or later
* lxml
* Pillow
* XlsxWriter (to use charting features)


.. admonition:: Note: font availability on Linux

   On most Linux distros, default PowerPoint fonts are typically unavailable. Thus, you should install them manually by either:

   - Installing ``ttf-mscorefonts-installer`` package (on Ubuntu).
   - Copying font files from the ``C:\Windows\Fonts`` folder of an existing Windows machine.
   - Manually installing the fonts following `this guide <https://wiki.debian.org/ppviewerFonts>`_.


