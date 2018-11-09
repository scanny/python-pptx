BEHAVE = behave
MAKE   = make
PYTHON = python
SETUP  = $(PYTHON) ./setup.py

.PHONY: accept clean cleandocs coverage docs readme sdist upload

help:
	@echo "Please use \`make <target>' where <target> is one or more of"
	@echo "  accept    run acceptance tests using behave"
	@echo "  clean     delete intermediate work product and start fresh"
	@echo "  cleandocs delete cached HTML documentation and start fresh"
	@echo "  coverage  run nosetests with coverage"
	@echo "  docs      build HTML documentation using Sphinx (incremental)"
	@echo "  opendocs  open local HTML documentation in browser"
	@echo "  readme    update README.html from README.rst"
	@echo "  sdist     generate a source distribution into dist/"
	@echo "  upload    upload distribution tarball to PyPI"

accept:
	$(BEHAVE) --stop

clean:
	find . -type f -name \*.pyc -exec rm {} \;
	find . -type f -name .DS_Store -exec rm {} \;
	rm -rf dist .coverage

cleandocs:
	$(MAKE) -C docs clean

coverage:
	py.test --cov-report term-missing --cov=pptx --cov=tests

docs:
	$(MAKE) -C docs html

opendocs:
	open docs/.build/html/index.html

sdist:
	$(SETUP) sdist

upload:
	$(SETUP) sdist upload
