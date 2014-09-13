#!/usr/bin/env python

import os
import re

from setuptools import find_packages, setup

# Read the version from pptx.__version__ without importing the package
# (and thus attempting to import packages it depends on that may not be
# installed yet)
thisdir = os.path.dirname(__file__)
init_py = os.path.join(thisdir, 'pptx', '__init__.py')
version = re.search("__version__ = '([^']+)'", open(init_py).read()).group(1)
license = os.path.join(thisdir, 'LICENSE')


NAME = 'python-pptx'
VERSION = version
DESCRIPTION = 'Generate and manipulate Open XML PowerPoint (.pptx) files'
KEYWORDS = 'powerpoint ppt pptx office open xml'
AUTHOR = 'Steve Canny'
AUTHOR_EMAIL = 'python-pptx@googlegroups.com'
URL = 'http://github.com/scanny/python-pptx'
LICENSE = open(license).read()
PACKAGES = find_packages(exclude=['tests', 'tests.*'])
PACKAGE_DATA = {'pptx': ['templates/*']}

INSTALL_REQUIRES = ['lxml>=2.3.2', 'Pillow>=2.0', 'XlsxWriter>=0.5.7']
TEST_SUITE = 'tests'
TESTS_REQUIRE = ['behave', 'mock', 'pyparsing>=2.0.1', 'pytest', 'unittest2']

CLASSIFIERS = [
    'Development Status :: 4 - Beta',
    'Environment :: Console',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: MIT License',
    'Operating System :: OS Independent',
    'Programming Language :: Python',
    'Programming Language :: Python :: 2',
    'Programming Language :: Python :: 2.6',
    'Programming Language :: Python :: 2.7',
    'Topic :: Office/Business :: Office Suites',
    'Topic :: Software Development :: Libraries'
]

readme = os.path.join(thisdir, 'README.rst')
history = os.path.join(thisdir, 'HISTORY.rst')
LONG_DESCRIPTION = open(readme).read() + '\n\n' + open(history).read()


params = {
    'name':             NAME,
    'version':          VERSION,
    'description':      DESCRIPTION,
    'keywords':         KEYWORDS,
    'long_description': LONG_DESCRIPTION,
    'author':           AUTHOR,
    'author_email':     AUTHOR_EMAIL,
    'url':              URL,
    'license':          LICENSE,
    'packages':         PACKAGES,
    'package_data':     PACKAGE_DATA,
    'install_requires': INSTALL_REQUIRES,
    'tests_require':    TESTS_REQUIRE,
    'test_suite':       TEST_SUITE,
    'classifiers':      CLASSIFIERS,
}

setup(**params)
