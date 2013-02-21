#!/usr/bin/env python

import os

from setuptools import setup
from glob import glob

execfile('pptx/version.py')

NAME         = 'python-pptx'
VERSION      = __version__
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

