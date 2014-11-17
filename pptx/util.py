# encoding: utf-8

"""
Utility functions and classes that come in handy when working with PowerPoint
and Open XML.
"""

from __future__ import absolute_import, division, print_function

import platform

from warnings import warn


class Length(int):
    """
    Base class for length classes Inches, Emu, Cm, Mm, Pt, and Px. Provides
    properties for converting length values to convenient units.
    """

    _EMUS_PER_INCH = 914400
    _EMUS_PER_CENTIPOINT = 127
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PT = 12700
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
        return self // self._EMUS_PER_CENTIPOINT

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
    def pt(self):
        """
        Floating point length in points
        """
        return self / float(self._EMUS_PER_PT)

    @property
    def px(self):
        """
        Integer length in pixels. Note this value is platform dependent,
        using 96 pixels/inch on Windows, 72 pixels/inch on all other
        platforms.
        """
        msg = (
            'Length.px property is deprecated and will be removed in a futur'
            'e release.'
        )
        warn(msg, UserWarning, stacklevel=2)

        # round can somtimes return values like x.999999 which are truncated
        # to x by int(); adding the 0.1 prevents this
        return int(round(self / float(self._EMUS_PER_PX)) + 0.1)


class Inches(Length):
    """
    Convenience constructor for length in inches
    """
    def __new__(cls, inches):
        emu = int(inches * Length._EMUS_PER_INCH)
        return Length.__new__(cls, emu)


class Centipoints(Length):
    """
    Convenience constructor for length in hundredths of a point
    """
    def __new__(cls, centipoints):
        emu = int(centipoints * Length._EMUS_PER_CENTIPOINT)
        return Length.__new__(cls, emu)


class Cm(Length):
    """
    Convenience constructor for length in centimeters
    """
    def __new__(cls, cm):
        emu = int(cm * Length._EMUS_PER_CM)
        return Length.__new__(cls, emu)


class Emu(Length):
    """
    Convenience constructor for length in english metric units
    """
    def __new__(cls, emu):
        return Length.__new__(cls, int(emu))


class Mm(Length):
    """
    Convenience constructor for length in millimeters
    """
    def __new__(cls, mm):
        emu = int(mm * Length._EMUS_PER_MM)
        return Length.__new__(cls, emu)


class Pt(Length):
    """
    Convenience value class for specifying a length in points
    """
    def __new__(cls, points):
        emu = int(points * Length._EMUS_PER_PT)
        return Length.__new__(cls, emu)


class Px(Length):
    """
    Convenience constructor for length in pixels
    """
    def __new__(cls, px):
        """
        Deprecated. Use one of the other Length classes.
        """
        msg = (
            'Px class is deprecated and will be removed in a future release.'
        )
        warn(msg, UserWarning, stacklevel=2)

        emu = int(px * Length._EMUS_PER_PX)
        return Length.__new__(cls, emu)


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
