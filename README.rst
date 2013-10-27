###########
python-pptx
###########

VERSION: 0.2.6 (seventh beta release)


STATUS (as of June 22 2013)
===========================

Stable beta release. Under active development, with new features added in a new
release roughly once a month.


Documentation
=============

Documentation is hosted on Read The Docs (readthedocs.org) at
https://python-pptx.readthedocs.org/en/latest/. The documentation is now in
reasonably robust shape and is being developed steadily alongside the code.


Reaching out
============

We'd love to hear from you if you like |pp|, want a new feature, find a bug,
need help using it, or just have a word of encouragement.

The **mailing list** for |pp| is python-pptx@googlegroups.com

The **issue tracker** is on github at `scanny/python-pptx`_.

Feature requests are best broached initially on the mailing list, they can be
added to the issue tracker once we've clarified the best approach,
particularly the appropriate API signature.

.. _`scanny/python-pptx`:
   https://github.com/scanny/python-pptx


Installation
============

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

|pp| depends on the ``lxml`` package and the Python Imaging Library
(``PIL``). Both ``pip`` and ``easy_install`` will take care of satisfying
those dependencies for you, but if you use this last method you will need to
install those yourself.


License
=======

Licensed under the `MIT license`_. Short version: this code is copyrighted by
me (Steve Canny), I give you permission to do what you want with it except
remove my name from the credits. See the LICENSE file for specific terms.

.. _MIT license:
   http://www.opensource.org/licenses/mit-license.php

.. |pp| replace:: ``python-pptx``
