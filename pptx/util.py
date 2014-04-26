# encoding: utf-8

"""
Utility functions and classes that come in handy when working with PowerPoint
and Open XML.
"""

import platform


class BaseLength(int):
    """
    Base class for length classes Inches, Emu, Cm, Mm, Pt, and Px. Provides
    properties for converting length values to convenient units.
    """

    _EMUS_PER_INCH = 914400
    _EMUS_PER_CENTIPOINT = 127
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PX = 9525 if platform.system() == 'Windows' else 12700

    def __new__(cls, emu):
        return int.__new__(cls, emu)

    @property
    def inches(self):
        """
        Floating point length in inches
        """
        return self / float(self._EMUS_PER_INCH)

    @property
    def centipoints(self):
        """
        Integer length in hundredths of a point (1/7200 inch). Used
        internally because PowerPoint stores font size in centipoints.
        """
        return self / self._EMUS_PER_CENTIPOINT

    @property
    def cm(self):
        """
        Floating point length in centimeters
        """
        return self / float(self._EMUS_PER_CM)

    @property
    def emu(self):
        """
        Integer length in English Metric Units
        """
        return self

    @property
    def mm(self):
        """
        Floating point length in millimeters
        """
        return self / float(self._EMUS_PER_MM)

    @property
    def px(self):
        """
        Integer length in pixels. Note this value is platform dependent,
        using 96 pixels/inch on Windows, 72 pixels/inch on all other
        platforms.
        """
        # round can somtimes return values like x.999999 which are truncated
        # to x by int(); adding the 0.1 prevents this
        return int(round(self / float(self._EMUS_PER_PX)) + 0.1)


class Inches(BaseLength):
    """
    Convenience constructor for length in inches
    """
    def __new__(cls, inches):
        emu = int(inches * BaseLength._EMUS_PER_INCH)
        return BaseLength.__new__(cls, emu)


class Centipoints(BaseLength):
    """
    Convenience constructor for length in hundredths of a point
    """
    def __new__(cls, centipoints):
        emu = int(centipoints * BaseLength._EMUS_PER_CENTIPOINT)
        return BaseLength.__new__(cls, emu)


class Cm(BaseLength):
    """
    Convenience constructor for length in centimeters
    """
    def __new__(cls, cm):
        emu = int(cm * BaseLength._EMUS_PER_CM)
        return BaseLength.__new__(cls, emu)


class Emu(BaseLength):
    """
    Convenience constructor for length in english metric units
    """
    def __new__(cls, emu):
        return BaseLength.__new__(cls, int(emu))


class Mm(BaseLength):
    """
    Convenience constructor for length in millimeters
    """
    def __new__(cls, mm):
        emu = int(mm * BaseLength._EMUS_PER_MM)
        return BaseLength.__new__(cls, emu)


class Pt(int):
    """
    Convenience value class for specifying a length in points
    """

    _UNITS_PER_POINT = 100

    def __new__(cls, pts):
        units = int(pts * Pt._UNITS_PER_POINT)
        return int.__new__(cls, units)


class Px(BaseLength):
    """
    Convenience constructor for length in pixels
    """
    def __new__(cls, px):
        emu = int(px * BaseLength._EMUS_PER_PX)
        return BaseLength.__new__(cls, emu)


class Collection(object):
    """
    Base class for collection classes. May also be used for part collections
    that don't yet have any custom methods.

    Has the following characteristics.:

    * Container (implements __contains__)
    * Iterable (delegates __iter__ to |list|)
    * Sized (implements __len__)
    * Sequence (delegates __getitem__ to |list|)
    """
    def __init__(self):
        super(Collection, self).__init__()
        self._values_ = []

    @property
    def _values(self):
        """Return read-only reference to collection values (list)."""
        return self._values_

    def __contains__(self, item):  # __iter__ would do this job by itself
        """Supports 'in' operator (e.g. 'x in collection')."""
        return (item in self._values_)

    def __getitem__(self, key):
        """Provides indexed access, (e.g. 'collection[0]')."""
        return self._values_.__getitem__(key)

    def __iter__(self):
        """Supports iteration (e.g. 'for x in collection: pass')."""
        return self._values_.__iter__()

    def __len__(self):
        """Supports len() function (e.g. 'len(collection) == 1')."""
        return len(self._values_)

    def index(self, item):
        """Supports index method (e.g. '[1, 2, 3].index(2) == 1')."""
        return self._values_.index(item)


def lazyproperty(f):
    """
    @lazyprop decorator. Decorated method will be called only on first access
    to calculate a cached property value. After that, the cached value is
    returned.
    """
    cache_attr_name = '_%s' % f.__name__  # like '_foobar' for prop 'foobar'
    docstring = f.__doc__

    def get_prop_value(obj):
        try:
            return getattr(obj, cache_attr_name)
        except AttributeError:
            value = f(obj)
            setattr(obj, cache_attr_name, value)
            return value

    return property(get_prop_value, doc=docstring)


def to_unicode(text):
    """
    Return *text* as a unicode string.

    *text* can be a 7-bit ASCII string, a UTF-8 encoded 8-bit string, or
    unicode. String values are converted to unicode assuming UTF-8 encoding.
    Unicode values are returned unchanged.
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
