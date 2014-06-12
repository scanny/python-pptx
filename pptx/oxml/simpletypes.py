# encoding: utf-8

"""
Simple type classes, providing validation and format translation for values
stored in XML element attributes. Naming generally corresponds to the simple
type in the associated XML schema.
"""

from __future__ import absolute_import, print_function

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
            return value
        try:
            if isinstance(value, basestring):
                return value
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


class XsdAnyUri(BaseStringType):
    """
    There's a regular expression this is supposed to meet but so far thinking
    spending cycles on validating wouldn't be worth it for the number of
    programming errors it would catch.
    """


class XsdBoolean(BaseSimpleType):

    @classmethod
    def convert_from_xml(cls, str_value):
        return str_value in ('1', 'true')

    @classmethod
    def convert_to_xml(cls, value):
        return {True: '1', False: '0'}[value]

    @classmethod
    def validate(cls, value):
        if value not in (True, False):
            raise TypeError(
                "only True or False (and possibly None) may be assigned, got"
                " '%s'" % value
            )


class XsdId(BaseStringType):
    """
    String that must begin with a letter or underscore and cannot contain any
    colons. Not fully validated because not used in external API.
    """
    pass


class XsdInt(BaseIntType):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(value, -2147483648, 2147483647)


class XsdLong(BaseIntType):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(
            value, -9223372036854775808, 9223372036854775807
        )


class XsdString(BaseStringType):
    pass


class XsdUnsignedInt(BaseIntType):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(value, 0, 4294967295)


class ST_ContentType(XsdString):
    """
    Has a pretty wicked regular expression it needs to match in the schema,
    but figuring it's not worth the trouble or run time to identify
    a programming error (as opposed to a user/runtime error).
    """
    pass


class ST_Coordinate(BaseSimpleType):

    @classmethod
    def convert_from_xml(cls, str_value):
        if 'i' in str_value or 'm' in str_value or 'p' in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        return int(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        return str(value)

    @classmethod
    def validate(cls, value):
        ST_CoordinateUnqualified.validate(value)


class ST_Coordinate32(BaseSimpleType):
    """
    xsd:union of ST_Coordinate32Unqualified, ST_UniversalMeasure
    """
    @classmethod
    def convert_from_xml(cls, str_value):
        if 'i' in str_value or 'm' in str_value or 'p' in str_value:
            return ST_UniversalMeasure.convert_from_xml(str_value)
        return ST_Coordinate32Unqualified.convert_from_xml(str_value)

    @classmethod
    def convert_to_xml(cls, value):
        return ST_Coordinate32Unqualified.convert_to_xml(value)

    @classmethod
    def validate(cls, value):
        ST_Coordinate32Unqualified.validate(value)


class ST_Coordinate32Unqualified(XsdInt):
    pass


class ST_CoordinateUnqualified(XsdLong):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(value, -27273042329600, 27273042316900)


class ST_DrawingElementId(XsdUnsignedInt):
    pass


class ST_Extension(XsdString):
    """
    Has a regular expression it needs to match in the schema, but figuring
    it's not worth the trouble or run time to identify a programming error
    (as opposed to a user/runtime error).
    """
    pass


class ST_HexColorRGB(BaseStringType):

    @classmethod
    def convert_to_xml(cls, value):
        """
        Keep alpha characters all uppercase just for consistency.
        """
        return value.upper()

    @classmethod
    def validate(cls, value):
        # must be string ---------------
        str_value = cls.validate_string(value)

        # must be 6 chars long----------
        if len(str_value) != 6:
            raise ValueError(
                "RGB string must be six characters long, got '%s'"
                % str_value
            )

        # must parse as hex int --------
        try:
            int(str_value, 16)
        except ValueError:
            raise ValueError(
                "RGB string must be valid hex string, got '%s'"
                % str_value
            )


class ST_LineWidth(XsdInt):

    @classmethod
    def convert_from_xml(cls, str_value):
        return Emu(int(str_value))

    @classmethod
    def validate(cls, value):
        super(ST_LineWidth, cls).validate(value)
        if value < 0 or value > 20116800:
            raise ValueError(
                "value must be in range 0-20116800 inclusive (0-1584 points)"
                ", got %d" % value
            )


class ST_Percentage(BaseIntType):
    """
    String value can be either an integer, representing 1000ths of a percent,
    or a floating point literal with a '%' suffix.
    """
    @classmethod
    def convert_from_xml(cls, str_value):
        if '%' in str_value:
            return cls._convert_from_percent_literal(str_value)
        return int(str_value)

    @classmethod
    def _convert_from_percent_literal(cls, str_value):
        float_part = str_value[:-1]  # trim off '%' character
        percent_value = float(float_part)
        int_value = int(round(percent_value * 1000))
        return int_value


class ST_PositiveCoordinate(XsdLong):

    @classmethod
    def validate(cls, value):
        cls.validate_int_in_range(value, 0, 27273042316900)


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


class ST_TargetMode(XsdString):
    """
    The valid values for the ``TargetMode`` attribute in a Relationship
    element, either 'External' or 'Internal'.
    """
    @classmethod
    def validate(cls, value):
        cls.validate_string(value)
        if value not in ('External', 'Internal'):
            raise ValueError(
                "must be one of 'Internal' or 'External', got '%s'" % value
            )


class ST_TextFontSize(BaseIntType):

    @classmethod
    def validate(cls, value):
        cls.validate_int(value)
        if value < 100 or value > 400000:
            raise ValueError(
                "value must be in range 100 -> 400000 (1-4000 points), got"
                " %d" % value
            )


class ST_TextTypeface(XsdString):
    pass


class ST_UniversalMeasure(BaseSimpleType):

    @classmethod
    def convert_from_xml(cls, str_value):
        float_part, units_part = str_value[:-2], str_value[-2:]
        quantity = float(float_part)
        multiplier = {
            'mm': 36000, 'cm': 360000, 'in': 914400, 'pt': 12700,
            'pc': 152400, 'pi': 152400
        }[units_part]
        emu_value = int(round(quantity * multiplier))
        return emu_value
