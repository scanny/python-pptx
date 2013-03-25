#!/usr/bin/env python

import os
import re

from setuptools import setup
from glob import glob

# Read the version from pptx.__version__ without importing the package
# (and thus attempting to import packages it depends on that may not be
# installed yet)
thisdir = os.path.dirname(__file__)
init_py_path = os.path.join(thisdir, 'pptx', '__init__.py')
version = re.search("__version__ = '([^']+)'",
                    open(init_py_path).read()).group(1)



NAME         = 'python-pptx'
VERSION      = version
DESCRIPTION  = 'Generate and manipulate Open XML PowerPoint (.pptx) files'
AUTHOR       = 'Steve Canny'
AUTHOR_EMAIL = 'python.pptx@librelist.com'
URL          = 'http://github.com/scanny/python-pptx'
LICENSE      = 'MIT'
PACKAGES     = ['pptx']
PACKAGE_DATA = {'pptx': ['templates/*']}

INSTALL_REQUIRES = ['lxml', 'PIL']
TEST_SUITE       = 'test'
TESTS_REQUIRE    = ['unittest2', 'mock', 'PyHamcrest', 'behave']

CLASSIFIERS =\
    [ 'Development Status :: 4 - Beta'
    , 'Environment :: Console'
    , 'Intended Audience :: Developers'
    , 'License :: OSI Approved :: MIT License'
    , 'Operating System :: OS Independent'
    , 'Programming Language :: Python'
    , 'Programming Language :: Python :: 2'
    , 'Programming Language :: Python :: 2.6'
    , 'Programming Language :: Python :: 2.7'
    , 'Topic :: Office/Business :: Office Suites'
    , 'Topic :: Software Development :: Libraries'
    ]

readme = os.path.join(os.path.dirname(__file__), 'README.rst')
LONG_DESCRIPTION = open(readme).read()


params =\
    { 'name'             : NAME
    , 'version'          : VERSION
    , 'description'      : DESCRIPTION
    , 'long_description' : LONG_DESCRIPTION
    , 'author'           : AUTHOR
    , 'author_email'     : AUTHOR_EMAIL
    , 'url'              : URL
    , 'license'          : LICENSE
    , 'packages'         : PACKAGES
    , 'package_data'     : PACKAGE_DATA
    , 'install_requires' : INSTALL_REQUIRES
    , 'tests_require'    : TESTS_REQUIRE
    , 'test_suite'       : TEST_SUITE
    , 'classifiers'      : CLASSIFIERS
    }

setup(**params)

