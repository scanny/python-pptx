.. _install:

Installing
==========

|pp| is hosted on PyPI, so installing with `pip` is simple::

    pip install python-pptx

It can also be installed using ``easy_install``, but that is `not
recommended`_::

    easy_install python-pptx

If neither ``pip`` nor ``easy_install`` is available, it can be installed
manually by downloading the distribution from PyPI, unpacking the tarball,
and running ``setup.py``::

    tar xvzf python-pptx-0.1.0a1.tar.gz
    cd python-pptx-0.1.0a1
    python setup.py install

|pp| depends on the ``lxml`` package and ``Pillow``, the modern version of
the Python Imaging Library (``PIL``). The charting features depend on
``XlsxWriter``. Both ``pip`` and ``easy_install`` will take care of
satisfying these dependencies for you, but if you use the ``setup.py``
installation method you will need to install the dependencies yourself.

Currently |pp| requires Python 2.6, 2.7, 3.3 or 3.4.

Dependencies
------------

* Python 2.6, 2.7, 3.3, or 3.4
* lxml
* Pillow
* XlsxWriter (to use charting features)

.. _`not recommended`:
   https://stackoverflow.com/questions/3220404/why-use-pip-over-easy-install
