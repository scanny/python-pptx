=======================================
Welcome to python-pptx's documentation!
=======================================

:mod:`python-pptx` is a pure-Python library for reading and writing PowerPoint
(``.pptx``) files.


---------------
Checking it out
---------------

Browse :doc:`examples with screenshots <user/examples>` to get a quick idea
what you can do with |pp|.


------
Status
------

The current release is v0.2.0 dated February 10, 2013. The library is under
active development. The current release has the following basic capabilities:

* Round-trip any Open XML presentation (.pptx file) including all its elements
* Add slides
* Populate text placeholders, for example to create a bullet slide
* Add an image to a slide with arbitrary position and size
* Add a textbox to a slide
* Manipulate text font size and bold

Additional capabilities are actively being developed and added on a release
cadence of roughly two-weeks. If you find a feature you need that |pp| doesn't
do yet, reach out via the mailing list or issue tracker and we'll see if we
can jump the queue for you to pop it in there :)

Currently |pp| requires Python 2.6.
Support for earlier versions is not likely to be on the horizon, but if
someone has a compelling need and is willing to pitch in to make it possible,
that could happen. Support for Python 3.x is also planned, but will likely
need to wait until more of the corpus of 2.x libraries get 3.x versions.


------------
Reaching out
------------

We'd love to hear from you if you like |pp|, want a new feature, find a bug,
need help using it, or just have a word of encouragement.

The **mailing list** for |pp| is python.pptx@librelist.com.

The **issue tracker** is on github at `scanny/python-pptx`_.

Feature requests are best broached initially on the mailing list, they can be
added to the issue tracker once we've clarified the best approach,
particularly the appropriate API signature.

.. _`scanny/python-pptx`:
   https://github.com/scanny/python-pptx


----------
Installing
----------

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


-----------------
Running the tests
-----------------

|pp| has a robust test suite, comprising over 180 tests at the time of this
writing, both at the acceptance test and unit test levels. ``unittest2`` is
used for unit tests, with help from ``PyHamcrest`` matchers and the excellent
``mock`` library. ``behave`` is used for acceptance tests.

You can run the tests from the folder containing the extracted source
distribution by issuing the following commands::

    $ nosetests
    .........................
    ----------------------------------------------------------------------
    Ran 178 tests in 2.279s
    
    OK
    
    $ behave
    Feature: Add a text box to a slide
      In order to accommodate a requirement for free-form text on a slide
      As a presentation developer
      I need the ability to place a text box on a slide
      
      Scenario: Add a text box to a slide 
        Given I have a reference to a blank slide
        When I add a text box to the slide's shape collection
        And I save the presentation
        Then the text box appears in the slide

    # ... more output ...

    6 features passed, 0 failed, 0 skipped
    6 scenarios passed, 0 failed, 0 skipped
    24 steps passed, 0 failed, 0 skipped, 0 undefined
    Took 0m0.8s


---------------
Getting Started
---------------

A quick way to get started is by trying out some of
:doc:`the examples <user/examples>` to get a feel for how to use |pp|.

Once you've gotten your feet wet, :doc:`this page <user/index>` can help you
build a better understanding of the object hierarchy that is central to using
|pp|.

The :doc:`user API documentation <user/modules/pptx>` can help you with the
fine details of calling signatures and behaviors.


----

Contents:

.. toctree::
   :titlesonly:
   :maxdepth: 2

   user/index
   developer/index



Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

.. |pp| replace:: ``python-pptx``
