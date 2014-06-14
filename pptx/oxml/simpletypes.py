# encoding: utf-8

"""
Simple type classes, providing validation and format translation for values
stored in XML element attributes. Naming generally corresponds to the simple
type in the associated XML schema.
"""

from __future__ import absolute_import

from ..util import Emu


class BaseSimpleType(object):

    @classmethod
    def from_xml(cls, str_value):
        return cls.convert_from_xml(str_value)

    @classmethod
    def to_xml(cls, value):
        cls.validate(value)
        str_value = cls.convert_to_xml(value)
        return str_value

    @classmethod
    def validate_int(cls, value):
        if not isinstance(value, int):
            raise TypeError(
                "value must be <type 'int'>, got %s" % type(value)
            )

    @classmethod
    def validate_int_in_range(cls, value, min_inclusive, max_inclusive):
        cls.validate_int(value)
        if value < min_inclusive or value > max_inclusive:
            raise ValueError(
                "value must be in range %d to %d inclusive, got %d" %
                (min_inclusive, max_inclusive, value)
            )

    @classmethod
    def validate_string(cls, value):
        if isinstance(value, str):
            return
        try:
            if isinstance(value, basestring):
                return
        except NameError:  # means we're on Python 3
            pass
        raise TypeError(
            "value must be a string, got %s" % type(value)
        )


class BaseStringType(BaseSimpleType):

    @classmethod
    def convert_from_xml(cls, str_value):
        return str_value

    @classmethod
    def convert_to_xml(cls, value):
        return value

    @classmethod
    def validate(cls, value):
        cls.validate_string(value)


class BaseIntType(BaseSimpleType):

    @classmethod
    def convert_from_xml(cls, str_value):
        return int(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        return str(value)

    @classmethod
    def validate(cls, value):
        cls.validate_int(value)


class XsdString(BaseStringType):
    pass


class XsdUnsignedInt(BaseIntType):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(value, 0, 4294967295)


class ST_SlideId(XsdUnsignedInt):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(value, 256, 2147483647)


class ST_SlideSizeCoordinate(BaseIntType):

    @classmethod
    def convert_from_xml(cls, str_value):
        return Emu(str_value)

    @classmethod
    def validate(cls, value):
        cls.validate_int(value)
        if value < 914400 or value > 51206400:
            raise ValueError(
                "value must be in range(914400, 51206400) (1-56 inches), got"
                " %d" % value
            )
