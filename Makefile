PYTHON      = $(shell test -x bin/python && echo bin/python || \
                      echo `which python`)
BEHAVE      = behave
SETUP       = $(PYTHON) ./setup.py

# BUILD_NUMBER ?= 1
# PYVERS        = $(shell $(PYTHON) -c 'import sys; print "%s.%s" % sys.version_info[0:2]')
# SIGN_KEY     ?= nerds@simplegeo.com
# SRCDIR       := oauth2
# VIRTUALENV    = $(shell /bin/echo -n `which virtualenv || \
#                                       which virtualenv-$(PYVERS) || \
#                                       which virtualenv$(PYVERS)`)
# VIRTUALENV   += --no-site-packages
# COVERAGE      = $(shell test -x bin/coverage && echo bin/coverage || echo true)
# DEPS         := $(shell find $(PWD)/deps -type f -printf "file://%p ")
# EGG          := $(shell $(SETUP) --fullname)-py$(PYVERS).egg
# EZ_INSTALL    = $(SETUP) easy_install -f "$(DEPS)"
# OS           := $(shell uname)
# PAGER        ?= less
# PLATFORM      = $(shell $(PYTHON) -c "from pkg_resources import get_build_platform; print get_build_platform()")
# PYLINT        = bin/pylint
# ROOT          = $(shell pwd)
# ROOTCMD       = fakeroot
# SDIST        := $(shell $(SETUP) --fullname).tar.gs
# SOURCES      := $(shell find $(SRCDIR) -type f -name \*.py -not -name 'test_*')
# TESTS        := $(shell find $(SRCDIR) -type f -name test_\*.py)
# COVERED      := $(SOURCES)


.PHONY: test sdist clean

help:
	@echo "Please use \`make <target>' where <target> is one or more of"
	@echo "  accept    to run acceptance tests using behave"
	@echo "  clean     to delete intermediate work product and start fresh"
	@echo "  test      to run tests using setup.py"
	@echo "  register  to update metadata (README.rst) on PyPI"
	@echo "  sdist     to generate a source distribution into dist/"
	@echo "  upload    to upload distribution tarball to PyPI"

accept:
	$(BEHAVE)

clean:
	find . -type f -name \*.pyc -exec rm {} \;
	rm -rf dist python_pptx.egg-info .coverage .DS_Store lint.html lint.txt

register:
	$(SETUP) register

sdist:
	$(SETUP) sdist

test:
	$(SETUP) test

upload:
	$(SETUP) sdist upload

# all: egg
# egg: dist/$(EGG)
#
# dist/$(EGG):
# 	$(SETUP) bdist_egg

# xunit.xml: bin/nosetests $(SOURCES) $(TESTS)
# 	$(SETUP) test --with-xunit --xunit-file=$@
#
# bin/nosetests: bin/easy_install
# 	@$(EZ_INSTALL) nose
#
# coverage: .coverage
# 	@$(COVERAGE) html -d $@ $(COVERED)
#
# coverage.xml: .coverage
# 	@$(COVERAGE) xml $(COVERED)
#
# .coverage: $(SOURCES) $(TESTS) bin/coverage bin/nosetests
# 	-@$(COVERAGE) run $(SETUP) test
#
# bin/coverage: bin/easy_install
# 	@$(EZ_INSTALL) coverage
#
# profile: .profile bin/pyprof2html
# 	bin/pyprof2html -o $@ $<
#
# .profile: $(SOURCES) bin/nosetests
# 	-$(SETUP) test -q --with-profile --profile-stats-file=$@
#
# bin/pyprof2html: bin/easy_install bin/
# 	@$(EZ_INSTALL) pyprof2html
#
# docs: $(SOURCES) bin/epydoc
# 	@echo bin/epydoc -q --html --no-frames -o $@ ...
# 	@bin/epydoc -q --html --no-frames -o $@ $(SOURCES)
#
# bin/epydoc: bin/easy_install
# 	@$(EZ_INSTALL) epydoc
#
# bin/pep8: bin/easy_install
# 	@$(EZ_INSTALL) pep8
#
# pep8: bin/pep8
# 	@bin/pep8 --repeat --ignore E225 $(SRCDIR)
#
# pep8.txt: bin/pep8
# 	@bin/pep8 --repeat --ignore E225 $(SRCDIR) > $@
#
# lint: bin/pylint
# 	-$(PYLINT) -f colorized $(SRCDIR)
#
# lint.html: bin/pylint
# 	-$(PYLINT) -f html $(SRCDIR) > $@
#
# lint.txt: bin/pylint
# 	-$(PYLINT) -f parseable $(SRCDIR) > $@
#
# bin/pylint: bin/easy_install
# 	@$(EZ_INSTALL) pylint
#
# README.html: README.mkd | bin/markdown
# 	bin/markdown -e utf-8 $^ -f $@
#
# bin/markdown: bin/easy_install
# 	@$(EZ_INSTALL) Markdown
#
# # Development setup
# rtfm:
# 	$(PAGER) README.mkd
#
# tags: TAGS.gz
#
# TAGS.gz: TAGS
# 	gzip $^
#
# TAGS: $(SOURCES)
# 	ctags -eR .
#
# env: bin/easy_install
#
# bin/easy_install:
# 	$(VIRTUALENV) .
# 	-test -f deps/setuptools* && $@ -U deps/setuptools*
#
# dev: develop
# develop: env
# 	nice -n 20 $(SETUP) develop
# 	@echo "            ---------------------------------------------"
# 	@echo "            To activate the development environment, run:"
# 	@echo "                           . bin/activate"
# 	@echo "            ---------------------------------------------"

# scrap_targets:
# 	@if test "$(OS)" = "Linux"; then $(ROOTCMD) debian/rules clean; fi

# xclean: extraclean
# extraclean: clean
# 	rm -rf bin lib .Python include