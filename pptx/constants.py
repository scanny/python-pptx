# -*- coding: utf-8 -*-
#
# constants.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Constant values modeled after those in the MS Office API.
"""


class MSO_AUTO_SHAPE_TYPE(object):
    """
    Constants corresponding to the msoAutoShapeType enumeration in the
    MS API. Standard abbreviation is 'MAST', e.g.:

        from pptx.spec import MSO_AUTO_SHAPE_TYPE as MAST

    """
    # msoAutoShapeType -----------------
    ACTION_BUTTON_BACK_OR_PREVIOUS = 129
    ACTION_BUTTON_BEGINNING = 131
    ACTION_BUTTON_CUSTOM = 125
    ACTION_BUTTON_DOCUMENT = 134
    ACTION_BUTTON_END = 132
    ACTION_BUTTON_FORWARD_OR_NEXT = 130
    ACTION_BUTTON_HELP = 127
    ACTION_BUTTON_HOME = 126
    ACTION_BUTTON_INFORMATION = 128
    ACTION_BUTTON_MOVIE = 136
    ACTION_BUTTON_RETURN = 133
    ACTION_BUTTON_SOUND = 135
    ARC = 25
    BALLOON = 137
    BENT_ARROW = 41
    BENT_UP_ARROW = 44
    BEVEL = 15
    BLOCK_ARC = 20
    CAN = 13
    CHART_PLUS = 182
    CHART_STAR = 181
    CHART_X = 180
    CHEVRON = 52
    CHORD = 161
    CIRCULAR_ARROW = 60
    CLOUD = 179
    CLOUD_CALLOUT = 108
    CORNER = 162
    CORNER_TABS = 169
    CROSS = 11
    CUBE = 14
    CURVED_DOWN_ARROW = 48
    CURVED_DOWN_RIBBON = 100
    CURVED_LEFT_ARROW = 46
    CURVED_RIGHT_ARROW = 45
    CURVED_UP_ARROW = 47
    CURVED_UP_RIBBON = 99
    DECAGON = 144
    DIAGONAL_STRIPE = 141
    DIAMOND = 4
    DODECAGON = 146
    DONUT = 18
    DOUBLE_BRACE = 27
    DOUBLE_BRACKET = 26
    DOUBLE_WAVE = 104
    DOWN_ARROW = 36
    DOWN_ARROW_CALLOUT = 56
    DOWN_RIBBON = 98
    EXPLOSION1 = 89
    EXPLOSION2 = 90
    FLOWCHART_ALTERNATE_PROCESS = 62
    FLOWCHART_CARD = 75
    FLOWCHART_COLLATE = 79
    FLOWCHART_CONNECTOR = 73
    FLOWCHART_DATA = 64
    FLOWCHART_DECISION = 63
    FLOWCHART_DELAY = 84
    FLOWCHART_DIRECT_ACCESS_STORAGE = 87
    FLOWCHART_DISPLAY = 88
    FLOWCHART_DOCUMENT = 67
    FLOWCHART_EXTRACT = 81
    FLOWCHART_INTERNAL_STORAGE = 66
    FLOWCHART_MAGNETIC_DISK = 86
    FLOWCHART_MANUAL_INPUT = 71
    FLOWCHART_MANUAL_OPERATION = 72
    FLOWCHART_MERGE = 82
    FLOWCHART_MULTIDOCUMENT = 68
    FLOWCHART_OFFLINE_STORAGE = 139
    FLOWCHART_OFFPAGE_CONNECTOR = 74
    FLOWCHART_OR = 78
    FLOWCHART_PREDEFINED_PROCESS = 65
    FLOWCHART_PREPARATION = 70
    FLOWCHART_PROCESS = 61
    FLOWCHART_PUNCHED_TAPE = 76
    FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 85
    FLOWCHART_SORT = 80
    FLOWCHART_STORED_DATA = 83
    FLOWCHART_SUMMING_JUNCTION = 77
    FLOWCHART_TERMINATOR = 69
    FOLDED_CORNER = 16
    FRAME = 158
    FUNNEL = 174
    GEAR_6 = 172
    GEAR_9 = 173
    HALF_FRAME = 159
    HEART = 21
    HEPTAGON = 145
    HEXAGON = 10
    HORIZONTAL_SCROLL = 102
    ISOSCELES_TRIANGLE = 7
    LEFT_ARROW = 34
    LEFT_ARROW_CALLOUT = 54
    LEFT_BRACE = 31
    LEFT_BRACKET = 29
    LEFT_CIRCULAR_ARROW = 176
    LEFT_RIGHT_ARROW = 37
    LEFT_RIGHT_ARROW_CALLOUT = 57
    LEFT_RIGHT_CIRCULAR_ARROW = 177
    LEFT_RIGHT_RIBBON = 140
    LEFT_RIGHT_UP_ARROW = 40
    LEFT_UP_ARROW = 43
    LIGHTNING_BOLT = 22
    LINE_CALLOUT_1 = 109
    LINE_CALLOUT_1_ACCENT_BAR = 113
    LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 121
    LINE_CALLOUT_1_NO_BORDER = 117
    LINE_CALLOUT_2 = 110
    LINE_CALLOUT_2_ACCENT_BAR = 114
    LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 122
    LINE_CALLOUT_2_NO_BORDER = 118
    LINE_CALLOUT_3 = 111
    LINE_CALLOUT_3_ACCENT_BAR = 115
    LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 123
    LINE_CALLOUT_3_NO_BORDER = 119
    LINE_CALLOUT_4 = 112
    LINE_CALLOUT_4_ACCENT_BAR = 116
    LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 124
    LINE_CALLOUT_4_NO_BORDER = 120
    LINE_INVERSE = 183
    MATH_DIVIDE = 166
    MATH_EQUAL = 167
    MATH_MINUS = 164
    MATH_MULTIPLY = 165
    MATH_NOT_EQUAL = 168
    MATH_PLUS = 163
    MOON = 24
    NON_ISOSCELES_TRAPEZOID = 143
    NOTCHED_RIGHT_ARROW = 50
    NO_SYMBOL = 19
    OCTAGON = 6
    OVAL = 9
    OVAL_CALLOUT = 107
    PARALLELOGRAM = 2
    PENTAGON = 51
    PIE = 142
    PIE_WEDGE = 175
    PLAQUE = 28
    PLAQUE_TABS = 171
    QUAD_ARROW = 39
    QUAD_ARROW_CALLOUT = 59
    RECTANGLE = 1
    RECTANGULAR_CALLOUT = 105
    REGULAR_PENTAGON = 12
    RIGHT_ARROW = 33
    RIGHT_ARROW_CALLOUT = 53
    RIGHT_BRACE = 32
    RIGHT_BRACKET = 30
    RIGHT_TRIANGLE = 8
    ROUNDED_RECTANGLE = 5
    ROUNDED_RECTANGULAR_CALLOUT = 106
    ROUND_1_RECTANGLE = 151
    ROUND_2_DIAG_RECTANGLE = 153
    ROUND_2_SAME_RECTANGLE = 152
    SMILEY_FACE = 17
    SNIP_1_RECTANGLE = 155
    SNIP_2_DIAG_RECTANGLE = 157
    SNIP_2_SAME_RECTANGLE = 156
    SNIP_ROUND_RECTANGLE = 154
    SQUARE_TABS = 170
    STAR_10_POINT = 149
    STAR_12_POINT = 150
    STAR_16_POINT = 94
    STAR_24_POINT = 95
    STAR_32_POINT = 96
    STAR_4_POINT = 91
    STAR_5_POINT = 92
    STAR_6_POINT = 147
    STAR_7_POINT = 148
    STAR_8_POINT = 93
    STRIPED_RIGHT_ARROW = 49
    SUN = 23
    SWOOSH_ARROW = 178
    TEAR = 160
    TRAPEZOID = 3
    UP_ARROW = 35
    UP_ARROW_CALLOUT = 55
    UP_DOWN_ARROW = 38
    UP_DOWN_ARROW_CALLOUT = 58
    UP_RIBBON = 97
    U_TURN_ARROW = 42
    VERTICAL_SCROLL = 101
    WAVE = 103


class MSO(object):
    """
    Constants corresponding to things like ``msoAnchorMiddle`` in the MS
    Office API.
    """
    # MsoVerticalAnchor ----------------
    ANCHOR_TOP = 1
    ANCHOR_TOP_BASELINE = 2
    ANCHOR_MIDDLE = 3
    ANCHOR_BOTTOM = 4
    ANCHOR_BOTTOM_BASELINE = 5
    VERTICAL_ANCHOR_MIXED = -2

    # Shape Types ----------------------

    # shapes recognized so far
    AUTO_SHAPE = 1
    PICTURE = 13
    PLACEHOLDER = 14
    TEXT_BOX = 17
    TABLE = 19
    # shape type backlog (in implementation sequence)
    GROUP = 6
    CHART = 3
    # shapes left to be recognized
    CALLOUT = 2  # not sure why, but callout auto shapes are distinguished
    CANVAS = 20
    COMMENT = 4
    DIAGRAM = 21
    EMBEDDED_OLE_OBJECT = 7
    FORM_CONTROL = 8
    FREEFORM = 5
    IGX_GRAPHIC = 24  # SmartArt graphic
    INK = 22
    INK_COMMENT = 23
    LINE = 9
    LINKED_OLE_OBJECT = 10
    LINKED_PICTURE = 11
    MEDIA = 16
    OLE_CONTROL_OBJECT = 12
    SCRIPT_ANCHOR = 18
    SHAPE_TYPE_MIXED = -2
    TEXT_EFFECT = 15
    WEB_VIDEO = 26

    # Connector Shapes =====================

    # msoShapeMixed -2
    #     Return value only; indicates a combination of the other states.
    # msoShapeNotPrimitive 138
    #     Not supported.

    # 'bentConnector2'
    # 'bentConnector3'
    # 'bentConnector4'
    # 'bentConnector5'

    # 'curvedConnector2'
    # 'curvedConnector3'
    # 'curvedConnector4'
    # 'curvedConnector5'

    # 'line'

    # 'straightConnector1'


class PP(object):
    """
    Constants specific to PowerPoint (PresentationML), with names like
    ``PP.ALIGN_CENTER`` corresponding to names like ``ppAlignCenter`` in the
    Microsoft Office API.
    """
    # PpParagraphAlignment -------------
    ALIGN_CENTER = 2
    ALIGN_DISTRIBUTE = 5
    ALIGN_JUSTIFY = 4
    ALIGN_JUSTIFY_LOW = 7
    ALIGN_LEFT = 1
    ALIGNMENT_MIXED = -2
    ALIGN_RIGHT = 3
    ALIGN_THAI_DISTRIBUTE = 6


class TEXT_ALIGN_TYPE(object):
    """
    Constants containing the valid values of ST_TextAlignType in Open XML.
    These values appear in the ``algn`` attribute of the ``<a:pPr>`` element
    and other places in Open XML.
    """
    # ST_TextAlignType -----------------
    CENTER = 'ctr'
    DISTRIBUTE = 'dist'
    JUSTIFY = 'just'
    JUSTIFY_LOW = 'justLow'
    LEFT = 'l'
    RIGHT = 'r'
    THAI_DISTRIBUTE = 'thaiDist'


class TEXT_ANCHORING_TYPE(object):
    """
    Constants containing the valid values of ST_TextAnchoringType in Open
    XML. These values appear in the ``anchor`` attribute of the ``<a:tcPr>``
    element of a table cell and the ``<a:bodyPr>`` element of a text frame.
    Specifies the vertical alignment of text in the frame.
    """
    # ST_TextAnchoringType -------------
    TOP = 't'
    MIDDLE = 'ctr'
    BOTTOM = 'b'
    JUSTIFY = 'just'
    DISTRIBUTE = 'dist'
