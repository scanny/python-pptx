"""Enumerations used by text and related objects."""

from __future__ import annotations

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


class MSO_TEXT_STRIKE_TYPE(BaseXmlEnum):
    """Indicates the type of strikethrough for text.

    Used with :attr:`.Font.strikethrough` to specify the style of text
    strikethrough.

    Alias: ``MSO_STRIKE``

    Example::

        from pptx.enum.text import MSO_STRIKE

        run.font.strikethrough = MSO_STRIKE.DOUBLE_LINE

    MS API Name: `MsoTextStrikeType`

    http://msdn.microsoft.com/en-us/library/aa432639.aspx

    .. versionadded:: 2026.05.0
    """

    NONE = (0, "noStrike", "Specifies no strikethrough.")
    """Specifies no strikethrough."""

    SINGLE_LINE = (1, "sngStrike", "Specifies a single line strikethrough.")
    """Specifies a single line strikethrough."""

    DOUBLE_LINE = (2, "dblStrike", "Specifies a double line strikethrough.")
    """Specifies a double line strikethrough."""

    MIXED = (-2, "", "Specifies a mix of strikethrough types (read-only).")
    """Specifies a mix of strikethrough types (read-only)."""


MSO_STRIKE = MSO_TEXT_STRIKE_TYPE


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


class PP_AUTO_NUMBER_SCHEME(BaseXmlEnum):
    """Specifies the numbering scheme used to generate automatic numbers for a paragraph bullet.

    Used with :attr:`._BulletFormat.number_scheme` and as the ``scheme`` argument to
    :meth:`._BulletFormat.auto_number`.

    Alias: ``PP_AUTO_NUMBER``

    Example::

        from pptx.enum.text import PP_AUTO_NUMBER

        paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)

    MS API Name: `PpAutoNumberScheme` (mapped from OOXML ``ST_TextAutonumberScheme``).

    .. versionadded:: 2026.05.0
    """

    ALPHA_LC_PAREN_BOTH = (
        0,
        "alphaLcParenBoth",
        "Lowercase alphabetical characters with both parentheses, e.g. (a), (b), (c).",
    )
    """Lowercase alphabetical characters with both parentheses, e.g. (a), (b), (c)."""

    ALPHA_UC_PAREN_BOTH = (
        1,
        "alphaUcParenBoth",
        "Uppercase alphabetical characters with both parentheses, e.g. (A), (B), (C).",
    )
    """Uppercase alphabetical characters with both parentheses, e.g. (A), (B), (C)."""

    ALPHA_LC_PAREN_R = (
        2,
        "alphaLcParenR",
        "Lowercase alphabetical characters with a right parenthesis, e.g. a), b), c).",
    )
    """Lowercase alphabetical characters with a right parenthesis, e.g. a), b), c)."""

    ALPHA_UC_PAREN_R = (
        3,
        "alphaUcParenR",
        "Uppercase alphabetical characters with a right parenthesis, e.g. A), B), C).",
    )
    """Uppercase alphabetical characters with a right parenthesis, e.g. A), B), C)."""

    ALPHA_LC_PERIOD = (
        4,
        "alphaLcPeriod",
        "Lowercase alphabetical characters with a period, e.g. a., b., c.",
    )
    """Lowercase alphabetical characters with a period, e.g. a., b., c."""

    ALPHA_UC_PERIOD = (
        5,
        "alphaUcPeriod",
        "Uppercase alphabetical characters with a period, e.g. A., B., C.",
    )
    """Uppercase alphabetical characters with a period, e.g. A., B., C."""

    ARABIC_PAREN_BOTH = (
        6,
        "arabicParenBoth",
        "Arabic numerals with both parentheses, e.g. (1), (2), (3).",
    )
    """Arabic numerals with both parentheses, e.g. (1), (2), (3)."""

    ARABIC_PAREN_R = (
        7,
        "arabicParenR",
        "Arabic numerals with a right parenthesis, e.g. 1), 2), 3).",
    )
    """Arabic numerals with a right parenthesis, e.g. 1), 2), 3)."""

    ARABIC_PERIOD = (8, "arabicPeriod", "Arabic numerals with a period, e.g. 1., 2., 3.")
    """Arabic numerals with a period, e.g. 1., 2., 3."""

    ARABIC_PLAIN = (9, "arabicPlain", "Arabic numerals without punctuation, e.g. 1, 2, 3.")
    """Arabic numerals without punctuation, e.g. 1, 2, 3."""

    ROMAN_LC_PAREN_BOTH = (
        10,
        "romanLcParenBoth",
        "Lowercase Roman numerals with both parentheses, e.g. (i), (ii), (iii).",
    )
    """Lowercase Roman numerals with both parentheses, e.g. (i), (ii), (iii)."""

    ROMAN_UC_PAREN_BOTH = (
        11,
        "romanUcParenBoth",
        "Uppercase Roman numerals with both parentheses, e.g. (I), (II), (III).",
    )
    """Uppercase Roman numerals with both parentheses, e.g. (I), (II), (III)."""

    ROMAN_LC_PAREN_R = (
        12,
        "romanLcParenR",
        "Lowercase Roman numerals with a right parenthesis, e.g. i), ii), iii).",
    )
    """Lowercase Roman numerals with a right parenthesis, e.g. i), ii), iii)."""

    ROMAN_UC_PAREN_R = (
        13,
        "romanUcParenR",
        "Uppercase Roman numerals with a right parenthesis, e.g. I), II), III).",
    )
    """Uppercase Roman numerals with a right parenthesis, e.g. I), II), III)."""

    ROMAN_LC_PERIOD = (
        14,
        "romanLcPeriod",
        "Lowercase Roman numerals with a period, e.g. i., ii., iii.",
    )
    """Lowercase Roman numerals with a period, e.g. i., ii., iii."""

    ROMAN_UC_PERIOD = (
        15,
        "romanUcPeriod",
        "Uppercase Roman numerals with a period, e.g. I., II., III.",
    )
    """Uppercase Roman numerals with a period, e.g. I., II., III."""

    CIRCLE_NUM_DB_PLAIN = (
        16,
        "circleNumDbPlain",
        "Double-byte circled numbers, 1 through 10.",
    )
    """Double-byte circled numbers, 1 through 10."""

    CIRCLE_NUM_WD_BLACK_PLAIN = (
        17,
        "circleNumWdBlackPlain",
        "Wingdings black circled numbers.",
    )
    """Wingdings black circled numbers."""

    CIRCLE_NUM_WD_WHITE_PLAIN = (
        18,
        "circleNumWdWhitePlain",
        "Wingdings white circled numbers.",
    )
    """Wingdings white circled numbers."""

    ARABIC_DB_PERIOD = (
        19,
        "arabicDbPeriod",
        "Double-byte Arabic numerals with a double-byte period.",
    )
    """Double-byte Arabic numerals with a double-byte period."""

    ARABIC_DB_PLAIN = (20, "arabicDbPlain", "Double-byte Arabic numerals.")
    """Double-byte Arabic numerals."""

    EA_1_CHS_PERIOD = (
        21,
        "ea1ChsPeriod",
        "East-Asian Simplified Chinese with single-byte period.",
    )
    """East-Asian Simplified Chinese with single-byte period."""

    EA_1_CHS_PLAIN = (22, "ea1ChsPlain", "East-Asian Simplified Chinese.")
    """East-Asian Simplified Chinese."""

    EA_1_CHT_PERIOD = (
        23,
        "ea1ChtPeriod",
        "East-Asian Traditional Chinese with single-byte period.",
    )
    """East-Asian Traditional Chinese with single-byte period."""

    EA_1_CHT_PLAIN = (24, "ea1ChtPlain", "East-Asian Traditional Chinese.")
    """East-Asian Traditional Chinese."""

    EA_1_JPN_CHS_DB_PERIOD = (
        25,
        "ea1JpnChsDbPeriod",
        "East-Asian Japanese double-byte period.",
    )
    """East-Asian Japanese double-byte period."""

    EA_1_JPN_KOR_PLAIN = (26, "ea1JpnKorPlain", "East-Asian Japanese/Korean.")
    """East-Asian Japanese/Korean."""

    EA_1_JPN_KOR_PERIOD = (
        27,
        "ea1JpnKorPeriod",
        "East-Asian Japanese/Korean with single-byte period.",
    )
    """East-Asian Japanese/Korean with single-byte period."""

    ARABIC_1_MINUS = (28, "arabic1Minus", "Bidi Arabic 1 (AraAlpha) with ANSI minus symbol.")
    """Bidi Arabic 1 (AraAlpha) with ANSI minus symbol."""

    ARABIC_2_MINUS = (29, "arabic2Minus", "Bidi Arabic 2 (AraAbjad) with ANSI minus symbol.")
    """Bidi Arabic 2 (AraAbjad) with ANSI minus symbol."""

    HEBREW_2_MINUS = (30, "hebrew2Minus", "Bidi Hebrew 2 with ANSI minus symbol.")
    """Bidi Hebrew 2 with ANSI minus symbol."""

    THAI_ALPHA_PERIOD = (31, "thaiAlphaPeriod", "Thai alphabet with a period.")
    """Thai alphabet with a period."""

    THAI_ALPHA_PAREN_R = (32, "thaiAlphaParenR", "Thai alphabet with a right parenthesis.")
    """Thai alphabet with a right parenthesis."""

    THAI_ALPHA_PAREN_BOTH = (33, "thaiAlphaParenBoth", "Thai alphabet with both parentheses.")
    """Thai alphabet with both parentheses."""

    THAI_NUM_PERIOD = (34, "thaiNumPeriod", "Thai numerals with a period.")
    """Thai numerals with a period."""

    THAI_NUM_PAREN_R = (35, "thaiNumParenR", "Thai numerals with a right parenthesis.")
    """Thai numerals with a right parenthesis."""

    THAI_NUM_PAREN_BOTH = (36, "thaiNumParenBoth", "Thai numerals with both parentheses.")
    """Thai numerals with both parentheses."""

    HINDI_ALPHA_PERIOD = (37, "hindiAlphaPeriod", "Hindi alphabet with a period.")
    """Hindi alphabet with a period."""

    HINDI_NUM_PERIOD = (38, "hindiNumPeriod", "Hindi numerals with a period.")
    """Hindi numerals with a period."""

    HINDI_NUM_PAREN_R = (39, "hindiNumParenR", "Hindi numerals with a right parenthesis.")
    """Hindi numerals with a right parenthesis."""

    HINDI_ALPHA_1_PERIOD = (40, "hindiAlpha1Period", "Hindi alphabet period (variant).")
    """Hindi alphabet period (variant)."""


PP_AUTO_NUMBER = PP_AUTO_NUMBER_SCHEME
