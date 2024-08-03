BEHAVE = behave
MAKE   = make

.PHONY: help
help:
	@echo "Please use \`make <target>' where <target> is one or more of"
	@echo "  accept       run acceptance tests using behave"
	@echo "  build        generate both sdist and wheel suitable for upload to PyPI"
	@echo "  clean        delete intermediate work product and start fresh"
	@echo "  cleandocs    delete cached HTML documentation and start fresh"
	@echo "  coverage     run nosetests with coverage"
	@echo "  docs         build HTML documentation using Sphinx (incremental)"
	@echo "  opendocs     open local HTML documentation in browser"
	@echo "  test-upload  upload distribution to TestPyPI"
	@echo "  upload       upload distribution tarball to PyPI"

.PHONY: accept
accept:
	$(BEHAVE) --stop

.PHONY: build
build:
	rm -rf dist
	python -m build
	twine check dist/*

.PHONY: clean
clean:
	find . -type f -name \*.pyc -exec rm {} \;
	find . -type f -name .DS_Store -exec rm {} \;
	rm -rf dist .coverage

.PHONY: cleandocs
cleandocs:
	$(MAKE) -C docs clean

.PHONY: coverage
coverage:
	py.test --cov-report term-missing --cov=pptx --cov=tests

.PHONY: docs
docs:
	$(MAKE) -C docs html

.PHONY: opendocs
opendocs:
	open docs/.build/html/index.html

.PHONY: test-upload
test-upload: build
	twine upload --repository testpypi dist/*

.PHONY: upload
upload: clean build
	twine upload dist/*
