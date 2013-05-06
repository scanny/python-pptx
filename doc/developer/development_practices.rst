#####################
Development Practices
#####################

Release procedure
=================

* develop branch git status clean

  + complete and commit any pending feature branches
  + stash any incomplete work

* release commit message prepared

  + review and summarize changes since prior release

* release branch committed

  + git flow release start 0.2.2
  + update version number in `pptx/__init__.py`
  + update setup.py (shouldn't usually be any changes)
  + update README.py, in particular, the release history
  + update doc/index.rst
  + confirm docs compile without errors
  + run all tests (behave, nosetests, tox)
  + create trial distribution (make clean sdist)
  + git commit -m 'Prepare v0.2.2 for release'
  + git flow release finish 0.2.2
  + (x) create tag for new version (automatically done by git flow)

* release uploaded to PyPI

  + upload: ``make upload``

* local repo synchronized with github

  + push tags to github: ``$ git push --tags``

* docs regenerated

  + trigger regeneration of docs on RTD.org


Procedure -- Adding a new feature
=================================

* issue added to github issue tracker
* git flow feature branch created

  + ``git flow feature start paragraph-level``

* working analysis documented
* acceptance test failing (not just raising exceptions)
* recursively, outside in:

  + unit test failing
  + next level method(s) written
  + unit test passing

* all tests passing

  + unit
  + acceptance
  + tox
  + visual confirmation of behavior in PowerPoint

* documentation updated as required

  + API additions
  + example code

* feature branch committed, rebased if required
* feature branch merged into develop

  + ``git flow feature finish paragraph-level``

* changes pushed to github
* issue closed


Outside-in layers
-----------------

* API wrapper method (if applicable)
* Internal API method
* objectify manipulation layer
* perhaps others


Creating slide images for documentation
=======================================

* Desired slide created using a test script
* Screenshot file on desktop using Cmd-Shift-4, Space, click
* Screenshot image scaled to 280 x 210px using Photoshop
* Completed image saved in ``doc/_static/img/``


Acceptance testing with ``behave``
==================================

... using *behave* for now for acceptance testing ...


Installation
------------

::

   pip install behave


Tutorial
--------

The `behave tutorial`_ is well worth working through.

.. _behave tutorial:
   http://packages.python.org/behave/tutorial.html

And this more detailed set of `examples and tutorials`_ is great for getting
the practicalities down.

.. _examples and tutorials:
   http://jenisys.github.com/behave.example/index.html


``behave`` Resources
--------------------

* `INVEST in Good Stories, and SMART Tasks`_
* `The Secret Ninja Cucumber Scrolls`_
* `Behavior Driven Outside In Development Explained, Part 1`_
* The `behave website`_ contains excellent documentation on installing and
  using behave.

.. _`INVEST in Good Stories, and SMART Tasks`:
   http://xp123.com/articles/invest-in-good-stories-and-smart-tasks/

.. _`The Secret Ninja Cucumber Scrolls`:
   http://cuke4ninja.com/sec_cucumber_jargon.html

.. _`Behavior Driven Outside In Development Explained, Part 1`:
   http://www.knwang.com/behavior-driven-outside-in-development-explai

.. _behave website:
   http://packages.python.org/behave/index.html


