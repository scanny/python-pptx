# encoding: utf-8

"""Provides Python 2/3 compatibility objects."""

import sys

import collections

try:
    Container = collections.abc.Container
    Mapping = collections.abc.Mapping
    Sequence = collections.abc.Sequence
except AttributeError:
    Container = collections.Container
    Mapping = collections.Mapping
    Sequence = collections.Sequence

if sys.version_info >= (3, 0):
    from .python3 import (  # noqa
        BytesIO,
        is_integer,
        is_string,
        is_unicode,
        to_unicode,
        Unicode,
    )
else:
    from .python2 import (  # noqa
        BytesIO,
        is_integer,
        is_string,
        is_unicode,
        to_unicode,
        Unicode,
    )
