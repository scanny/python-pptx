# encoding: utf-8

"""
Provides Python 2 compatibility objects
"""

from StringIO import StringIO as BytesIO  # noqa


def is_integer(obj):
    """
    Return True if *obj* is an integer (int, long), False otherwise.
    """
    return isinstance(obj, (int, long))


def is_string(obj):
    """
    Return True if *obj* is a string, False otherwise.
    """
    return isinstance(obj, basestring)


def is_unicode(obj):
    """
    Return True if *obj* is a unicode string, False otherwise.
    """
    return isinstance(obj, unicode)


def to_unicode(text):
    """
    Return *text* as a unicode string. *text* can be a 7-bit ASCII string,
    a UTF-8 encoded 8-bit string, or unicode. String values are converted to
    unicode assuming UTF-8 encoding. Unicode values are returned unchanged.
    """
    # both str and unicode inherit from basestring
    if not isinstance(text, basestring):
        tmpl = 'expected UTF-8 encoded string or unicode, got %s value %s'
        raise TypeError(tmpl % (type(text), text))
    # return unicode strings unchanged
    if isinstance(text, unicode):
        return text
    # otherwise assume UTF-8 encoding, which also works for ASCII
    return unicode(text, 'utf-8')


Unicode = unicode
