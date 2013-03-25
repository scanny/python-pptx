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
  + git flow release finish 0.2.2
  + (x) create tag for new version (automatically done by git flow)

* release uploaded to PyPI

  + upload: ``make upload``

* local repo synchronized with github

  + push tags to github

* docs regenerated

  + trigger regeneration of docs on RTD.org



Prodedure -- Adding a new feature
=================================

* add issue to github issue tracker
* create git feature branch
* write acceptance test
* recursively, outside in:

  * write unit test
  * write next level method(s)

* check and adjust documentation as required
* finish feature branch (merge into development)
* close issue


Outside-in layers
-----------------

* API wrapper method (if applicable)
* Internal API method
* objectify manipulation layer
* perhaps others


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


