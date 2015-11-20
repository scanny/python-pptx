#!/usr/bin/env python

import os
import re

from setuptools import find_packages, setup


def ascii_bytes_from(path, *paths):
    """
    Return the ASCII characters in the file specified by *path* and *paths*.
    The file path is determined by concatenating *path* and any members of
    *paths* with a directory separator in between.
    """
    file_path = os.path.join(path, *paths)
    with open(file_path) as f:
        ascii_bytes = f.read()
    return ascii_bytes

# read required text from files
thisdir = os.path.dirname(__file__)
init_py = ascii_bytes_from(thisdir, 'pptx', '__init__.py')
readme = ascii_bytes_from(thisdir, 'README.rst')
history = ascii_bytes_from(thisdir, 'HISTORY.rst')
license = ascii_bytes_from(thisdir, 'LICENSE')

# Read the version from pptx.__version__ without importing the package
# (and thus attempting to import packages it depends on that may not be
# installed yet)
version = re.search("__version__ = '([^']+)'", init_py).group(1)


NAME = 'python-pptx'
VERSION = version
DESCRIPTION = 'Generate and manipulate Open XML PowerPoint (.pptx) files'
KEYWORDS = 'powerpoint ppt pptx office open xml'
AUTHOR = 'Steve Canny'
AUTHOR_EMAIL = 'python-pptx@googlegroups.com'
URL = 'http://github.com/scanny/python-pptx'
LICENSE = license
PACKAGES = find_packages(exclude=['tests', 'tests.*'])
PACKAGE_DATA = {'pptx': ['templates/*']}

INSTALL_REQUIRES = ['lxml>=3.1.0', 'Pillow>=2.6.1', 'XlsxWriter>=0.5.7']
TEST_SUITE = 'tests'
TESTS_REQUIRE = ['behave', 'mock', 'pyparsing>=2.0.1', 'pytest']

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
    'Programming Language :: Python :: 3',
    'Programming Language :: Python :: 3.3',
    'Programming Language :: Python :: 3.4',
    'Topic :: Office/Business :: Office Suites',
    'Topic :: Software Development :: Libraries'
]

LONG_DESCRIPTION = readme + '\n\n' + history


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
