python-pptx
===========

Release v\ |version| (:doc:`Installation <user/install>`)

:mod:`python-pptx` is a pure-Python library for reading and writing PowerPoint
(``.pptx``) files.


Checking it out
---------------

Browse :doc:`examples with screenshots <user/quickstart>` to get a quick idea
what you can do with |pp|.


Feature Support
---------------

|pp| has the following capabilities, with many more on the roadmap:

* Round-trip any Open XML presentation (.pptx file) including all its elements
* Add slides
* Populate text placeholders, for example to create a bullet slide
* Add image to slide at arbitrary position and size
* Add textbox to a slide; manipulate text font size and bold
* Add table to a slide
* Add auto shapes (e.g. polygons, flowchart shapes, etc.) to a slide
* Access and change core document properties such as title and subject

Additional capabilities are actively being developed and added on a release
cadence of roughly once per month. If you find a feature you need that |pp|
doesn't yet have, reach out via the mailing list or issue tracker and we'll see
if we can jump the queue for you to pop it in there :)


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
    Ran 299 tests in 1.473s
    
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

    13 features passed, 0 failed, 0 skipped
    27 scenarios passed, 0 failed, 0 skipped
    108 steps passed, 0 failed, 0 skipped, 0 undefined
    Took 0m1.3s


Getting Started
---------------

A quick way to get started is by trying out some of
:doc:`the examples <user/quickstart>` to get a feel for how to use |pp|.

The :doc:`user API documentation <user/modules/pptx>` can help you with the
fine details of calling signatures and behaviors.


User Guide
----------

.. toctree::
   :maxdepth: 2

   user/install
   user/use-cases
   user/concepts
   user/autoshape-types
   user/quickstart


API Documentation
-----------------

.. toctree::
   :maxdepth: 2

   user/modules/index


Contributor Guide
-----------------

.. toctree::
   :maxdepth: 1

   dev/index



Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
