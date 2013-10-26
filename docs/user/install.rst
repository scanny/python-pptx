Installing
==========

|pp| is hosted on PyPI, so installation is relatively simple, and just
depends on what installation utilit(y/ies) you have installed.

|pp| may be installed with ``pip`` if you have it available::

    pip install python-pptx

It can also be installed using ``easy_install``::

    easy_install python-pptx

If neither ``pip`` nor ``easy_install`` is available, it can be installed
manually by downloading the distribution from PyPI, unpacking the tarball,
and running ``setup.py``::

    tar xvzf python-pptx-0.1.0a1.tar.gz
    cd python-pptx-0.1.0a1
    python setup.py install

|pp| depends on the ``lxml`` package and the Python Imaging Library (``PIL``).
Both ``pip`` and ``easy_install`` will take care of satisfying those
dependencies for you, but if you use this last method you will need to install
those yourself.


Dependencies
------------

* Python 2.6 or 2.7
* lxml
* Python Imaging Library (PIL)

Currently |pp| requires Python 2.6 or 2.7. Support for earlier versions is not
planned as it complicates future support for Python 3.x, but if you have a
compelling need please reach out via the mailing list. Support for Python 3.x
is underway and should be out shortly.
