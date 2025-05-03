"""Enumerations used by text and related objects."""

from __future__ import annotations
from enum import Enum, auto

from pptx.enum.base import BaseEnum, BaseXmlEnum


class MSO_AUTO_SIZE(BaseEnum):
    """Determines the type of automatic sizing allowed.

    The following names can be used to specify the automatic sizing behavior used to fit a shape's
    text within the shape bounding box, for example::

        from pptx.enum.text import MSO_AUTO_SIZE

        shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    The word-wrap setting of the text frame interacts with the auto-size setting to determine the
    specific auto-sizing behavior.

    Note that `TextFrame.auto_size` can also be set to |None|, which removes the auto size setting
    altogether. This causes the setting to be inherited, either from the layout placeholder, in the
    case of a placeholder shape, or from the theme.

    MS API Name: `MsoAutoSize`

    http://msdn.microsoft.com/en-us/library/office/ff865367(v=office.15).aspx
    """

    NONE = (
        0,
        "No automatic sizing of the shape or text will be done.\n\nText can freely extend beyond"
        " the horizontal and vertical edges of the shape bounding box.",
    )
    """No automatic sizing of the shape or text will be done.

    Text can freely extend beyond the horizontal and vertical edges of the shape bounding box.
    """

    SHAPE_TO_FIT_TEXT = (
        1,
        "The shape height and possibly width are adjusted to fit the text.\n\nNote this setting"
        " interacts with the TextFrame.word_wrap property setting. If word wrap is turned on,"
        " only the height of the shape will be adjusted; soft line breaks will be used to fit the"
        " text horizontally.",
    )
    """The shape height and possibly width are adjusted to fit the text.

    Note this setting interacts with the TextFrame.word_wrap property setting. If word wrap is
    turned on, only the height of the shape will be adjusted; soft line breaks will be used to fit
    the text horizontally.
    """

    TEXT_TO_FIT_SHAPE = (
        2,
        "The font size is reduced as necessary to fit the text within the shape.",
    )
    """The font size is reduced as necessary to fit the text within the shape."""

    MIXED = (-2, "Return value only; indicates a combination of automatic sizing schemes are used.")
    """Return value only; indicates a combination of automatic sizing schemes are used."""


class MSO_TEXT_UNDERLINE_TYPE(BaseXmlEnum):
    """
    Indicates the type of underline for text. Used with
    :attr:`.Font.underline` to specify the style of text underlining.

    Alias: ``MSO_UNDERLINE``

    Example::

        from pptx.enum.text import MSO_UNDERLINE

        run.font.underline = MSO_UNDERLINE.DOUBLE_LINE

    MS API Name: `MsoTextUnderlineType`

    http://msdn.microsoft.com/en-us/library/aa432699.aspx
    """

    NONE = (0, "none", "Specifies no underline.")
    """Specifies no underline."""

    DASH_HEAVY_LINE = (8, "dashHeavy", "Specifies a dash underline.")
    """Specifies a dash underline."""

    DASH_LINE = (7, "dash", "Specifies a dash line underline.")
    """Specifies a dash line underline."""

    DASH_LONG_HEAVY_LINE = (10, "dashLongHeavy", "Specifies a long heavy line underline.")
    """Specifies a long heavy line underline."""

    DASH_LONG_LINE = (9, "dashLong", "Specifies a dashed long line underline.")
    """Specifies a dashed long line underline."""

    DOT_DASH_HEAVY_LINE = (12, "dotDashHeavy", "Specifies a dot dash heavy line underline.")
    """Specifies a dot dash heavy line underline."""

    DOT_DASH_LINE = (11, "dotDash", "Specifies a dot dash line underline.")
    """Specifies a dot dash line underline."""

    DOT_DOT_DASH_HEAVY_LINE = (
        14,
        "dotDotDashHeavy",
        "Specifies a dot dot dash heavy line underline.",
    )
    """Specifies a dot dot dash heavy line underline."""

    DOT_DOT_DASH_LINE = (13, "dotDotDash", "Specifies a dot dot dash line underline.")
    """Specifies a dot dot dash line underline."""

    DOTTED_HEAVY_LINE = (6, "dottedHeavy", "Specifies a dotted heavy line underline.")
    """Specifies a dotted heavy line underline."""

    DOTTED_LINE = (5, "dotted", "Specifies a dotted line underline.")
    """Specifies a dotted line underline."""

    DOUBLE_LINE = (3, "dbl", "Specifies a double line underline.")
    """Specifies a double line underline."""

    HEAVY_LINE = (4, "heavy", "Specifies a heavy line underline.")
    """Specifies a heavy line underline."""

    SINGLE_LINE = (2, "sng", "Specifies a single line underline.")
    """Specifies a single line underline."""

    WAVY_DOUBLE_LINE = (17, "wavyDbl", "Specifies a wavy double line underline.")
    """Specifies a wavy double line underline."""

    WAVY_HEAVY_LINE = (16, "wavyHeavy", "Specifies a wavy heavy line underline.")
    """Specifies a wavy heavy line underline."""

    WAVY_LINE = (15, "wavy", "Specifies a wavy line underline.")
    """Specifies a wavy line underline."""

    WORDS = (1, "words", "Specifies underlining words.")
    """Specifies underlining words."""

    MIXED = (-2, "", "Specifies a mix of underline types (read-only).")
    """Specifies a mix of underline types (read-only)."""


MSO_UNDERLINE = MSO_TEXT_UNDERLINE_TYPE


class MSO_VERTICAL_ANCHOR(BaseXmlEnum):
    """Specifies the vertical alignment of text in a text frame.

    Used with the `.vertical_anchor` property of the |TextFrame| object. Note that the
    `vertical_anchor` property can also have the value None, indicating there is no directly
    specified vertical anchor setting and its effective value is inherited from its placeholder if
    it has one or from the theme. |None| may also be assigned to remove an explicitly specified
    vertical anchor setting.

    MS API Name: `MsoVerticalAnchor`

    http://msdn.microsoft.com/en-us/library/office/ff865255.aspx
    """

    TOP = (1, "t", "Aligns text to top of text frame")
    """Aligns text to top of text frame"""

    MIDDLE = (3, "ctr", "Centers text vertically")
    """Centers text vertically"""

    BOTTOM = (4, "b", "Aligns text to bottom of text frame")
    """Aligns text to bottom of text frame"""

    MIXED = (-2, "", "Return value only; indicates a combination of the other states.")
    """Return value only; indicates a combination of the other states."""


MSO_ANCHOR = MSO_VERTICAL_ANCHOR


class PP_PARAGRAPH_ALIGNMENT(BaseXmlEnum):
    """Specifies the horizontal alignment for one or more paragraphs.

    Alias: `PP_ALIGN`

    Example::

        from pptx.enum.text import PP_ALIGN

        shape.paragraphs[0].alignment = PP_ALIGN.CENTER

    MS API Name: `PpParagraphAlignment`

    http://msdn.microsoft.com/en-us/library/office/ff745375(v=office.15).aspx
    """

    CENTER = (2, "ctr", "Center align")
    """Center align"""

    DISTRIBUTE = (
        5,
        "dist",
        "Evenly distributes e.g. Japanese characters from left to right within a line",
    )
    """Evenly distributes e.g. Japanese characters from left to right within a line"""

    JUSTIFY = (
        4,
        "just",
        "Justified, i.e. each line both begins and ends at the margin.\n\nSpacing between words"
        " is adjusted such that the line exactly fills the width of the paragraph.",
    )
    """Justified, i.e. each line both begins and ends at the margin.

    Spacing between words is adjusted such that the line exactly fills the width of the paragraph.
    """

    JUSTIFY_LOW = (7, "justLow", "Justify using a small amount of space between words.")
    """Justify using a small amount of space between words."""

    LEFT = (1, "l", "Left aligned")
    """Left aligned"""

    RIGHT = (3, "r", "Right aligned")
    """Right aligned"""

    THAI_DISTRIBUTE = (6, "thaiDist", "Thai distributed")
    """Thai distributed"""

    MIXED = (-2, "", "Multiple alignments are present in a set of paragraphs (read-only).")
    """Multiple alignments are present in a set of paragraphs (read-only)."""


PP_ALIGN = PP_PARAGRAPH_ALIGNMENT


class BulletStyleType(Enum):
    """Types of bullet styles."""
    NO_BULLET = auto()
    CUSTOM = auto()
    NUMBERED = auto()
    DEFAULT = auto()


class MSO_NUMBERED_BULLET_STYLE(BaseXmlEnum):
    """Specifies the style of numbered bullets.

    Example::

        from pptx.enum.text import MSO_NUMBERED_BULLET_STYLE

        shape.paragraphs[0].bullet = MSO_NUMBERED_BULLET_STYLE.BULLET_ALPHA_UC_PERIOD

    MS API Name: `MsoNumberedBulletStyle`

    https://learn.microsoft.com/sv-se/office/vba/api/office.msonumberedbulletstyle
    """

    ALPHA_LC_PAREN_BOTH = (8, "alphaLCParenBoth", "Lowercase alphabetical bullet with opening and closing parentheses.")
    ALPHA_LC_PAREN_RIGHT = (9, "alphaLCParenRight", "Lowercase alphabetical bullet with closing parenthesis.")
    ALPHA_LC_PERIOD = (0, "alphaLCPeriod", "Lowercase alphabetical bullet with period.|")
    ALPHA_UC_PAREN_BOTH = (10, "alphaUCParenBoth", "Uppercase alphabetical bullet with opening and closing parentheses.")
    ALPHA_UC_PAREN_RIGHT = (11, "alphaUCParenRight", "Uppercase alphabetical bullet with closing parenthesis.")
    ALPHA_UC_PERIOD = (1, "alphaUCPeriod", "Uppercase alphabetical bullet with period.")
    ARABIC_ABJAD_DASH = (24, "arabicAbjadDash", "Arabic Abjad bullet with a dash.")
    ARABIC_ALPHA_DASH = (23, "arabicAlphaDash", "Arabic alphabetical bullet with a dash.")
    ARABIC_DB_PERIOD = (29, "arabicDBPeriod", "Arabic DB bullet with period.")
    ARABIC_DB_PLAIN = (28, "arabicDBPlain", "Plain Arabic DB bullet.")
    ARABIC_PAREN_BOTH = (12, "arabicParenBoth", "Arabic bullet with opening and closing parentheses.")
    ARABIC_PAREN_RIGHT = (2, "arabicParenRight", "Arabic bullet with closing parenthesis.")
    ARABIC_PERIOD = (3, "arabicPeriod", "Arabic bullet with period.")
    ARABIC_PLAIN = (13, "arabicPlain", "Plain Arabic bullet.")
    CIRCLE_NUM_DB_PLAIN = (18, "circleNumDBPlain", "Circled number bullet.")
    CIRCLE_NUM_WD_BLACK_PLAIN = (20, "circleNumWDBlackPlain", "Circled number WD black bullet.")
    CIRCLE_NUM_WD_WHITE_PLAIN = (19, "circleNumWDWhitePlain", "Circled number WD white bullet.")
    HEBREW_ALPHA_DASH = (25, "hebrewAlphaDash", "Hebrew alphabetical bullet with dash.")
    HINDI_ALPHA1_PERIOD = (40, "hindiAlpha1Period", "Hindi alphabetical bullet 1 with period.")
    HINDI_ALPHA_PERIOD = (36, "hindiAlphaPeriod", "Hindi alphabetical bullet with period.")
    HINDI_NUM_PAREN_RIGHT = (39, "hindiNumParenRight", "Hindi numbered bullet with closing parenthesis.")
    HINDI_NUM_PERIOD = (37, "hindiNumPeriod", "Hindi numbered bullet with period.")
    KANJI_KOREAN_PERIOD = (27, "kanjiKoreanPeriod", "Korean Kanji bullet with period.")
    KANJI_KOREAN_PLAIN = (26, "kanjiKoreanPlain", "Korean Kanji bullet.")
    KANJI_SIMP_CHIN_DB_PERIOD = (38, "kanjiSimpChinDBPeriod", "Simplified Chinese Kanji bulllet with period.")
    ROMAN_LC_PAREN_BOTH = (4, "romanLCParenBoth", "Lowercase roman bullet with opening and closing parentheses.")
    ROMAN_LC_PAREN_RIGHT = (5, "romanLCParenRight", "Lowercase roman bullet with closing parenthesis.")
    ROMAN_LC_PERIOD = (6, "romanLCPeriod", "Lowercase roman bullet with period.")
    ROMAN_UC_PAREN_BOTH = (14, "romanUCParenBoth", "Uppercase roman bullet with opening and closing parentheses.")
    ROMAN_UC_PAREN_RIGHT = (15, "romanUCParenRight", "Uppercase roman bullet with closing parenthesis.")
    ROMAN_UC_PERIOD = (7, "romanUCPeriod", "Uppercase roman bullet with period.")
    SIMP_CHIN_PERIOD = (17, "simpChinPeriod", "Simplified Chinese bulllet with period.")
    SIMP_CHIN_PLAIN = (16, "simpChinPlain", "Simplified Chinese bullet.")
    STYLE_MIXED = (-2, "styleMixed", "Return value only; indicates a combination of the other states. ")
    THAI_ALPHA_PAREN_BOTH = (32, "thaiAlphaParenBoth", "Thai alphabetical bullet with opening and closing parentheses.")
    THAI_ALPHA_PAREN_RIGHT = (31, "thaiAlphaParenRight", "Thai alphabetical bullet with closing parenthesis.")
    THAI_ALPHA_PERIOD = (30, "thaiAlphaPeriod", "Thai alphabetical bullet with period.")
    THAI_NUM_PAREN_BOTH = (35, "thaiNumParenBoth", "Thai numerical bullet with opening and closing parentheses.")
    THAI_NUM_PAREN_RIGHT = (34, "thaiNumParenRight", "Thai numerical bullet with closing parenthesis.")
    THAI_NUM_PERIOD = (33, "thaiNumPeriod", "Thai numerical bullet with period.")
    TRAD_CHIN_PERIOD = (22, "tradChinPeriod", "Traditional Chinese bulllet with period.")
    TRAD_CHIN_PLAIN = (21, "tradChinPlain", "Traditional Chinese bulllet.")
