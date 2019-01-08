# encoding: utf-8

"""
Provides Python 3 compatibility objects
"""

from io import BytesIO  # noqa


def is_integer(obj):
    """
    Return True if *obj* is an int, False otherwise.
    """
    return isinstance(obj, int)


def is_string(obj):
    """
    Return True if *obj* is a string, False otherwise.
    """
    return isinstance(obj, str)


def is_unicode(obj):
    """
    Return True if *obj* is a unicode string, False otherwise.
    """
    return isinstance(obj, str)


def to_unicode(text):
    """
    Return *text* as a unicode string. All text in Python 3 is unicode, so
    this just returns *text* unchanged.
    """
    if not isinstance(text, str):
        tmpl = 'expected unicode string, got %s value %s'
        raise TypeError(tmpl % (type(text), text))
    return text


Unicode = str
