
Running the test suite
======================

|pp| has a robust test suite, comprising over 600 tests at the time of writing,
both at the acceptance test and unit test levels. ``pytest`` is used for unit
tests, with help from the excellent ``mock`` library. ``behave`` is used for
acceptance tests.

You can run the tests from the source working directory using the following
commands::

    $ py.test
    ============================ test session starts ============================
    platform darwin -- Python 2.7.5 -- pytest-2.4.2
    plugins: cov
    collected 301 items

    tests/test_oxml.py .........................................
    tests/test_presentation.py ...................
    tests/test_spec.py ........
    tests/test_text.py ..........................
    tests/test_util.py ...............
    tests/opc/test_packaging.py .....................................................
    tests/opc/test_rels.py ..............
    tests/parts/test_coreprops.py .
    tests/parts/test_image.py ...........
    tests/parts/test_part.py ........
    tests/parts/test_slides.py ..................
    tests/shapes/test_autoshape.py ....................
    tests/shapes/test_picture.py .
    tests/shapes/test_placeholder.py .
    tests/shapes/test_shape.py ...........
    tests/shapes/test_shapetree.py ..............
    tests/shapes/test_table.py ........................................

    ======================= 301 passed in 2.00 seconds ======================

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
    30 scenarios passed, 0 failed, 0 skipped
    120 steps passed, 0 failed, 0 skipped, 0 undefined
    Took 0m1.3s


Corpus conformance tests
------------------------

``tests/test_conformance_corpus.py`` exercises the library against the
shared feature manifests maintained in the sibling ``loadfix/ooxml-reference-corpus``
repository. For each ``features/pptx/<name>.json`` manifest, the test
re-runs the committed generator script from that repo and feeds the
resulting ``.pptx`` through ``ooxml_validate.conformance.run_feature``.
A ``pass`` status means python-pptx's current output still satisfies
the manifest's assertion block; a ``fail`` status signals drift.

These tests **auto-skip** unless both of the following are present:

1. **Sibling checkout of the corpus** at ``../ooxml-reference-corpus/``
   (i.e. alongside your python-pptx checkout). Clone from
   ``https://github.com/loadfix/ooxml-reference-corpus``.
2. **`ooxml-validate`** importable in the current environment. It is
   not on PyPI; install it as an editable sibling checkout::

       git clone https://github.com/loadfix/ooxml-validate ../ooxml-validate
       pip install -e ../ooxml-validate

Once both are in place, the suite picks up every ``pptx/*.json``
manifest automatically::

    $ pytest tests/test_conformance_corpus.py -v

When the corpus or ``ooxml-validate`` is missing the suite skips
silently, so CI in environments without them still passes.
