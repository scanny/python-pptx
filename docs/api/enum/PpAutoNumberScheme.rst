.. _PpAutoNumberScheme:

``PP_AUTO_NUMBER_SCHEME``
=========================

Specifies the numbering scheme used to generate automatic numbers for a
paragraph bullet.

Alias: ``PP_AUTO_NUMBER``

Example::

    from pptx.enum.text import PP_AUTO_NUMBER

    paragraph.bullet.auto_number(PP_AUTO_NUMBER.ARABIC_PERIOD)

----

ALPHA_LC_PAREN_BOTH
    Lowercase alphabetical characters with both parentheses, e.g. (a), (b), (c).

ALPHA_UC_PAREN_BOTH
    Uppercase alphabetical characters with both parentheses, e.g. (A), (B), (C).

ALPHA_LC_PAREN_R
    Lowercase alphabetical characters with a right parenthesis, e.g. a), b), c).

ALPHA_UC_PAREN_R
    Uppercase alphabetical characters with a right parenthesis, e.g. A), B), C).

ALPHA_LC_PERIOD
    Lowercase alphabetical characters with a period, e.g. a., b., c.

ALPHA_UC_PERIOD
    Uppercase alphabetical characters with a period, e.g. A., B., C.

ARABIC_PAREN_BOTH
    Arabic numerals with both parentheses, e.g. (1), (2), (3).

ARABIC_PAREN_R
    Arabic numerals with a right parenthesis, e.g. 1), 2), 3).

ARABIC_PERIOD
    Arabic numerals with a period, e.g. 1., 2., 3.

ARABIC_PLAIN
    Arabic numerals without punctuation, e.g. 1, 2, 3.

ROMAN_LC_PAREN_BOTH
    Lowercase Roman numerals with both parentheses, e.g. (i), (ii), (iii).

ROMAN_UC_PAREN_BOTH
    Uppercase Roman numerals with both parentheses, e.g. (I), (II), (III).

ROMAN_LC_PAREN_R
    Lowercase Roman numerals with a right parenthesis, e.g. i), ii), iii).

ROMAN_UC_PAREN_R
    Uppercase Roman numerals with a right parenthesis, e.g. I), II), III).

ROMAN_LC_PERIOD
    Lowercase Roman numerals with a period, e.g. i., ii., iii.

ROMAN_UC_PERIOD
    Uppercase Roman numerals with a period, e.g. I., II., III.

CIRCLE_NUM_DB_PLAIN
    Double-byte circled numbers, 1 through 10.

CIRCLE_NUM_WD_BLACK_PLAIN
    Wingdings black circled numbers.

CIRCLE_NUM_WD_WHITE_PLAIN
    Wingdings white circled numbers.

ARABIC_DB_PERIOD
    Double-byte Arabic numerals with a double-byte period.

ARABIC_DB_PLAIN
    Double-byte Arabic numerals.

EA_1_CHS_PERIOD
    East-Asian Simplified Chinese with single-byte period.

EA_1_CHS_PLAIN
    East-Asian Simplified Chinese.

EA_1_CHT_PERIOD
    East-Asian Traditional Chinese with single-byte period.

EA_1_CHT_PLAIN
    East-Asian Traditional Chinese.

EA_1_JPN_CHS_DB_PERIOD
    East-Asian Japanese double-byte period.

EA_1_JPN_KOR_PLAIN
    East-Asian Japanese/Korean.

EA_1_JPN_KOR_PERIOD
    East-Asian Japanese/Korean with single-byte period.

ARABIC_1_MINUS
    Bidi Arabic 1 (AraAlpha) with ANSI minus symbol.

ARABIC_2_MINUS
    Bidi Arabic 2 (AraAbjad) with ANSI minus symbol.

HEBREW_2_MINUS
    Bidi Hebrew 2 with ANSI minus symbol.

THAI_ALPHA_PERIOD
    Thai alphabet with a period.

THAI_ALPHA_PAREN_R
    Thai alphabet with a right parenthesis.

THAI_ALPHA_PAREN_BOTH
    Thai alphabet with both parentheses.

THAI_NUM_PERIOD
    Thai numerals with a period.

THAI_NUM_PAREN_R
    Thai numerals with a right parenthesis.

THAI_NUM_PAREN_BOTH
    Thai numerals with both parentheses.

HINDI_ALPHA_PERIOD
    Hindi alphabet with a period.

HINDI_NUM_PERIOD
    Hindi numerals with a period.

HINDI_NUM_PAREN_R
    Hindi numerals with a right parenthesis.

HINDI_ALPHA_1_PERIOD
    Hindi alphabet period (variant).
