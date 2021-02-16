# encoding: utf-8

"""Enumerations used by DrawingML objects."""

from __future__ import absolute_import, division, print_function, unicode_literals

from .base import (
    alias,
    Enumeration,
    EnumMember,
    ReturnValueOnlyEnumMember,
    XmlEnumeration,
    XmlMappedEnumMember,
)


class MSO_COLOR_TYPE(Enumeration):
    """
    Specifies the color specification scheme

    Example::

        from pptx.enum.dml import MSO_COLOR_TYPE

        assert shape.fill.fore_color.type == MSO_COLOR_TYPE.SCHEME
    """

    __ms_name__ = "MsoColorType"

    __url__ = (
        "http://msdn.microsoft.com/en-us/library/office/ff864912(v=office.15" ").aspx"
    )

    __members__ = (
        EnumMember("RGB", 1, "Color is specified by an |RGBColor| value"),
        EnumMember("SCHEME", 2, "Color is one of the preset theme colors"),
        EnumMember(
            "HSL",
            101,
            """
            Color is specified using Hue, Saturation, and Luminosity values
            """,
        ),
        EnumMember(
            "PRESET",
            102,
            """
            Color is specified using a named built-in color
            """,
        ),
        EnumMember(
            "SCRGB",
            103,
            """
            Color is an scRGB color, a wide color gamut RGB color space
            """,
        ),
        EnumMember(
            "SYSTEM",
            104,
            """
            Color is one specified by the operating system, such as the
            window background color.
            """,
        ),
    )


@alias("MSO_FILL")
class MSO_FILL_TYPE(Enumeration):
    """
    Specifies the type of bitmap used for the fill of a shape.

    Alias: ``MSO_FILL``

    Example::

        from pptx.enum.dml import MSO_FILL

        assert shape.fill.type == MSO_FILL.SOLID
    """

    __ms_name__ = "MsoFillType"

    __url__ = "http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx"

    __members__ = (
        EnumMember(
            "BACKGROUND",
            5,
            """
            The shape is transparent, such that whatever is behind the shape
            shows through. Often this is the slide background, but if
            a visible shape is behind, that will show through.
            """,
        ),
        EnumMember("GRADIENT", 3, "Shape is filled with a gradient"),
        EnumMember(
            "GROUP",
            101,
            "Shape is part of a group and should inherit the "
            "fill properties of the group.",
        ),
        EnumMember("PATTERNED", 2, "Shape is filled with a pattern"),
        EnumMember("PICTURE", 6, "Shape is filled with a bitmapped image"),
        EnumMember("SOLID", 1, "Shape is filled with a solid color"),
        EnumMember("TEXTURED", 4, "Shape is filled with a texture"),
    )


@alias("MSO_LINE")
class MSO_LINE_DASH_STYLE(XmlEnumeration):
    """Specifies the dash style for a line.

    Alias: ``MSO_LINE``

    Example::

        from pptx.enum.dml import MSO_LINE

        shape.line.dash_style == MSO_LINE.DASH_DOT_DOT
    """

    __ms_name__ = "MsoLineDashStyle"

    __url__ = (
        "https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/mso"
        "linedashstyle-enumeration-office"
    )

    __members__ = (
        XmlMappedEnumMember("DASH", 4, "dash", "Line consists of dashes only."),
        XmlMappedEnumMember("DASH_DOT", 5, "dashDot", "Line is a dash-dot pattern."),
        XmlMappedEnumMember(
            "DASH_DOT_DOT", 6, "lgDashDotDot", "Line is a dash-dot-dot patte" "rn."
        ),
        XmlMappedEnumMember("LONG_DASH", 7, "lgDash", "Line consists of long dashes."),
        XmlMappedEnumMember(
            "LONG_DASH_DOT", 8, "lgDashDot", "Line is a long dash-dot patter" "n."
        ),
        XmlMappedEnumMember("ROUND_DOT", 3, "dot", "Line is made up of round dots."),
        XmlMappedEnumMember("SOLID", 1, "solid", "Line is solid."),
        XmlMappedEnumMember(
            "SQUARE_DOT", 2, "sysDash", "Line is made up of square dots."
        ),
        XmlMappedEnumMember("SYS_DOT", 2, "sysDot", "Line is solid."),
        ReturnValueOnlyEnumMember("DASH_STYLE_MIXED", -2, "Not supported."),
    )


@alias("MSO_PATTERN")
class MSO_PATTERN_TYPE(XmlEnumeration):
    """Specifies the fill pattern used in a shape.

    Alias: ``MSO_PATTERN``

    Example::

        from pptx.enum.dml import MSO_PATTERN

        fill = shape.fill
        fill.patterned()
        fill.pattern == MSO_PATTERN.WAVE
    """

    __ms_name__ = "MsoPatternType"

    __url__ = (
        "https://msdn.microsoft.com/VBA/Office-Shared-VBA/articles/msopatter"
        "ntype-enumeration-office"
    )

    __members__ = (
        XmlMappedEnumMember("CROSS", 51, "cross", "Cross"),
        XmlMappedEnumMember(
            "DARK_DOWNWARD_DIAGONAL", 15, "dkDnDiag", "Dark Downward Diagona" "l"
        ),
        XmlMappedEnumMember("DARK_HORIZONTAL", 13, "dkHorz", "Dark Horizontal"),
        XmlMappedEnumMember(
            "DARK_UPWARD_DIAGONAL", 16, "dkUpDiag", "Dark Upward Diagonal"
        ),
        XmlMappedEnumMember("DARK_VERTICAL", 14, "dkVert", "Dark Vertical"),
        XmlMappedEnumMember(
            "DASHED_DOWNWARD_DIAGONAL", 28, "dashDnDiag", "Dashed Downward D" "iagonal"
        ),
        XmlMappedEnumMember("DASHED_HORIZONTAL", 32, "dashHorz", "Dashed Horizontal"),
        XmlMappedEnumMember(
            "DASHED_UPWARD_DIAGONAL", 27, "dashUpDiag", "Dashed Upward Diago" "nal"
        ),
        XmlMappedEnumMember("DASHED_VERTICAL", 31, "dashVert", "Dashed Vertical"),
        XmlMappedEnumMember("DIAGONAL_BRICK", 40, "diagBrick", "Diagonal Brick"),
        XmlMappedEnumMember("DIAGONAL_CROSS", 54, "diagCross", "Diagonal Cross"),
        XmlMappedEnumMember("DIVOT", 46, "divot", "Pattern Divot"),
        XmlMappedEnumMember("DOTTED_DIAMOND", 24, "dotDmnd", "Dotted Diamond"),
        XmlMappedEnumMember("DOTTED_GRID", 45, "dotGrid", "Dotted Grid"),
        XmlMappedEnumMember("DOWNWARD_DIAGONAL", 52, "dnDiag", "Downward Diagonal"),
        XmlMappedEnumMember("HORIZONTAL", 49, "horz", "Horizontal"),
        XmlMappedEnumMember("HORIZONTAL_BRICK", 35, "horzBrick", "Horizontal Brick"),
        XmlMappedEnumMember(
            "LARGE_CHECKER_BOARD", 36, "lgCheck", "Large Checker Board"
        ),
        XmlMappedEnumMember("LARGE_CONFETTI", 33, "lgConfetti", "Large Confetti"),
        XmlMappedEnumMember("LARGE_GRID", 34, "lgGrid", "Large Grid"),
        XmlMappedEnumMember(
            "LIGHT_DOWNWARD_DIAGONAL", 21, "ltDnDiag", "Light Downward Diago" "nal"
        ),
        XmlMappedEnumMember("LIGHT_HORIZONTAL", 19, "ltHorz", "Light Horizontal"),
        XmlMappedEnumMember(
            "LIGHT_UPWARD_DIAGONAL", 22, "ltUpDiag", "Light Upward Diagonal"
        ),
        XmlMappedEnumMember("LIGHT_VERTICAL", 20, "ltVert", "Light Vertical"),
        XmlMappedEnumMember("NARROW_HORIZONTAL", 30, "narHorz", "Narrow Horizontal"),
        XmlMappedEnumMember("NARROW_VERTICAL", 29, "narVert", "Narrow Vertical"),
        XmlMappedEnumMember("OUTLINED_DIAMOND", 41, "openDmnd", "Outlined Diamond"),
        XmlMappedEnumMember("PERCENT_10", 2, "pct10", "10% of the foreground color."),
        XmlMappedEnumMember("PERCENT_20", 3, "pct20", "20% of the foreground color."),
        XmlMappedEnumMember("PERCENT_25", 4, "pct25", "25% of the foreground color."),
        XmlMappedEnumMember("PERCENT_30", 5, "pct30", "30% of the foreground color."),
        XmlMappedEnumMember("PERCENT_40", 6, "pct40", "40% of the foreground color."),
        XmlMappedEnumMember("PERCENT_5", 1, "pct5", "5% of the foreground color."),
        XmlMappedEnumMember("PERCENT_50", 7, "pct50", "50% of the foreground color."),
        XmlMappedEnumMember("PERCENT_60", 8, "pct60", "60% of the foreground color."),
        XmlMappedEnumMember("PERCENT_70", 9, "pct70", "70% of the foreground color."),
        XmlMappedEnumMember("PERCENT_75", 10, "pct75", "75% of the foreground color."),
        XmlMappedEnumMember("PERCENT_80", 11, "pct80", "80% of the foreground color."),
        XmlMappedEnumMember("PERCENT_90", 12, "pct90", "90% of the foreground color."),
        XmlMappedEnumMember("PLAID", 42, "plaid", "Plaid"),
        XmlMappedEnumMember("SHINGLE", 47, "shingle", "Shingle"),
        XmlMappedEnumMember(
            "SMALL_CHECKER_BOARD", 17, "smCheck", "Small Checker Board"
        ),
        XmlMappedEnumMember("SMALL_CONFETTI", 37, "smConfetti", "Small Confetti"),
        XmlMappedEnumMember("SMALL_GRID", 23, "smGrid", "Small Grid"),
        XmlMappedEnumMember("SOLID_DIAMOND", 39, "solidDmnd", "Solid Diamond"),
        XmlMappedEnumMember("SPHERE", 43, "sphere", "Sphere"),
        XmlMappedEnumMember("TRELLIS", 18, "trellis", "Trellis"),
        XmlMappedEnumMember("UPWARD_DIAGONAL", 53, "upDiag", "Upward Diagonal"),
        XmlMappedEnumMember("VERTICAL", 50, "vert", "Vertical"),
        XmlMappedEnumMember("WAVE", 48, "wave", "Wave"),
        XmlMappedEnumMember("WEAVE", 44, "weave", "Weave"),
        XmlMappedEnumMember(
            "WIDE_DOWNWARD_DIAGONAL", 25, "wdDnDiag", "Wide Downward Diagona" "l"
        ),
        XmlMappedEnumMember(
            "WIDE_UPWARD_DIAGONAL", 26, "wdUpDiag", "Wide Upward Diagonal"
        ),
        XmlMappedEnumMember("ZIG_ZAG", 38, "zigZag", "Zig Zag"),
        ReturnValueOnlyEnumMember("MIXED", -2, "Mixed pattern."),
    )

class MSO_PRESET_COLOR_INDEX(Enumeration):
    """
    Indicates the preset color using a string name
    """

    __ms_name__ = ""
    __url__ = ("")

    __members__ = (
        EnumMember("ALICEBLUE", "aliceBlue", "Alice Blue Preset Color"),
        EnumMember("ANTIQUEWHITE", "antiqueWhite", "Antique White Preset Color"),
        EnumMember("AQUA", "aqua", "Aqua Preset Color"),
        EnumMember("AQUAMARINE", "aquamarine", "Aquamarine Preset Color"),
        EnumMember("AZURE", "azure", "Azure Preset Color"),
        EnumMember("BEIGE", "beige", "Beige Preset Color"),
        EnumMember("BISQUE", "bisque", "Bisque Preset Color"),
        EnumMember("BLACK", "black", "Black Preset Color"),
        EnumMember("BLANCHEDALMOND", "blanchedAlmond", "Blanched Almond Preset Color"),
        EnumMember("BLUE", "blue", "Blue Preset Color"),
        EnumMember("BLUEVIOLET", "blueViolet", "Blue Violet Preset Color"),
        EnumMember("BROWN", "brown", "Brown Preset Color"),
        EnumMember("BURLYWOOD", "burlyWood", "Burly Wood Preset Color"),
        EnumMember("CADETBLUE", "cadetBlue", "Cadet Blue Preset Color"),
        EnumMember("CHARTREUSE", "chartreuse", "Chartreuse Preset Color"),
        EnumMember("CHOCOLATE", "chocolate", "Chocolate Preset Color"),
        EnumMember("CORAL", "coral", "Coral Preset Color"),
        EnumMember("CORNFLOWERBLUE", "cornflowerBlue", "Cornflower Blue Preset Color"),
        EnumMember("CORNSILK", "cornsilk", "Cornsilk Preset Color"),
        EnumMember("CRIMSON", "crimson", "Crimson Preset Color"),
        EnumMember("CYAN", "cyan", "Cyan Preset Color"),
        EnumMember("DKBLUE", "dkBlue", "Dark Blue Preset Color"),
        EnumMember("DKCYAN", "dkCyan", "Dark Cyan Preset Color"),
        EnumMember("DKGOLDENROD", "dkGoldenrod", "Dark Goldenrod Preset Color"),
        EnumMember("DKGRAY", "dkGray", "Dark Gray Preset Color"),
        EnumMember("DKGREEN", "dkGreen", "Dark Green Preset Color"),
        EnumMember("DKKHAKI", "dkKhaki", "Dark Khaki Preset Color"),
        EnumMember("DKMAGENTA", "dkMagenta", "Dark Magenta Preset Color"),
        EnumMember("DKOLIVEGREEN", "dkOliveGreen", "Dark Olive Green Preset Color"),
        EnumMember("DKORANGE", "dkOrange", "Dark Orange Preset Color"),
        EnumMember("DKORCHID", "dkOrchid", "Dark Orchid Preset Color"),
        EnumMember("DKRED", "dkRed", "Dark Red Preset Color"),
        EnumMember("DKSALMON", "dkSalmon", "Dark Salmon Preset Color"),
        EnumMember("DKSEAGREEN", "dkSeaGreen", "Dark Sea Green Preset Color"),
        EnumMember("DKSLATEBLUE", "dkSlateBlue", "Dark Slate Blue Preset Color"),
        EnumMember("DKSLATEGRAY", "dkSlateGray", "Dark Slate Gray Preset Color"),
        EnumMember("DKTURQUOISE", "dkTurquoise", "Dark Turquoise Preset Color"),
        EnumMember("DKVIOLET", "dkViolet", "Dark Violet Preset Color"),
        EnumMember("DEEPPINK", "deepPink", "Deep Pink Preset Color"),
        EnumMember("DEEPSKYBLUE", "deepSkyBlue", "Deep Sky Blue Preset Color"),
        EnumMember("DIMGRAY", "dimGray", "Dim Gray Preset Color"),
        EnumMember("DODGERBLUE", "dodgerBlue", "Dodger Blue Preset Color"),
        EnumMember("FIREBRICK", "firebrick", "Firebrick Preset Color"),
        EnumMember("FLORALWHITE", "floralWhite", "Floral White Preset Color"),
        EnumMember("FORESTGREEN", "forestGreen", "Forest Green Preset Color"),
        EnumMember("FUCHSIA", "fuchsia", "Fuchsia Preset Color"),
        EnumMember("GAINSBORO", "gainsboro", "Gainsboro Preset Color"),
        EnumMember("GHOSTWHITE", "ghostWhite", "Ghost White Preset Color"),
        EnumMember("GOLD", "gold", "Gold Preset Color"),
        EnumMember("GOLDENROD", "goldenrod", "Goldenrod Preset Color"),
        EnumMember("GRAY", "gray", "Gray Preset Color"),
        EnumMember("GREEN", "green", "Green Preset Color"),
        EnumMember("GREENYELLOW", "greenYellow", "Green Yellow Preset Color"),
        EnumMember("HONEYDEW", "honeydew", "Honeydew Preset Color"),
        EnumMember("HOTPINK", "hotPink", "Hot Pink Preset Color"),
        EnumMember("INDIANRED", "indianRed", "Indian Red Preset Color"),
        EnumMember("INDIGO", "indigo", "Indigo Preset Color"),
        EnumMember("IVORY", "ivory", "Ivory Preset Color"),
        EnumMember("KHAKI", "khaki", "Khaki Preset Color"),
        EnumMember("LAVENDER", "lavender", "Lavender Preset Color"),
        EnumMember("LAVENDERBLUSH", "lavenderBlush", "Lavender Blush Preset Color"),
        EnumMember("LAWNGREEN", "lawnGreen", "Lawn Green Preset Color"),
        EnumMember("LEMONCHIFFON", "lemonChiffon", "Lemon Chiffon Preset Color"),
        EnumMember("LTBLUE", "ltBlue", "Light Blue Preset Color"),
        EnumMember("LTCORAL", "ltCoral", "Light Coral Preset Color"),
        EnumMember("LTCYAN", "ltCyan", "Light Cyan Preset Color"),
        EnumMember("LTGOLDENRODYELLOW", "ltGoldenrodYellow", "Light Goldenrod Yellow Preset Color"),
        EnumMember("LTGRAY", "ltGray", "Light Gray Preset Color"),
        EnumMember("LTGREEN", "ltGreen", "Light Green Preset Color"),
        EnumMember("LTPINK", "ltPink", "Light Pink Preset Color"),
        EnumMember("LTSALMON", "ltSalmon", "Light Salmon Preset Color"),
        EnumMember("LTSEAGREEN", "ltSeaGreen", "Light Sea Green Preset Color"),
        EnumMember("LTSKYBLUE", "ltSkyBlue", "Light Sky Blue Preset Color"),
        EnumMember("LTSLATEGRAY", "ltSlateGray", "Light Slate Gray Preset Color"),
        EnumMember("LTSTEELBLUE", "ltSteelBlue", "Light Steel Blue Preset Color"),
        EnumMember("LTYELLOW", "ltYellow", "Light Yellow Preset Color"),
        EnumMember("LIME", "lime", "Lime Preset Color"),
        EnumMember("LIMEGREEN", "limeGreen", "Lime Green Preset Color"),
        EnumMember("LINEN", "linen", "Linen Preset Color"),
        EnumMember("MAGENTA", "magenta", "Magenta Preset Color"),
        EnumMember("MAROON", "maroon", "Maroon Preset Color"),
        EnumMember("MEDAQUAMARINE", "medAquamarine", "Medium Aquamarine Preset Color"),
        EnumMember("MEDBLUE", "medBlue", "Medium Blue Preset Color"),
        EnumMember("MEDORCHID", "medOrchid", "Medium Orchid Preset Color"),
        EnumMember("MEDPURPLE", "medPurple", "Medium Purple Preset Color"),
        EnumMember("MEDSEAGREEN", "medSeaGreen", "Medium Sea Green Preset Color"),
        EnumMember("MEDSLATEBLUE", "medSlateBlue", "Medium Slate Blue Preset Color"),
        EnumMember("MEDSPRINGGREEN", "medSpringGreen", "Medium Spring Green Preset Color"),
        EnumMember("MEDTURQUOISE", "medTurquoise", "Medium Turquoise Preset Color"),
        EnumMember("MEDVIOLETRED", "medVioletRed", "Medium Violet Red Preset Color"),
        EnumMember("MIDNIGHTBLUE", "midnightBlue", "Midnight Blue Preset Color"),
        EnumMember("MINTCREAM", "mintCream", "Mint Cream Preset Color"),
        EnumMember("MISTYROSE", "mistyRose", "Misty Rose Preset Color"),
        EnumMember("MOCCASIN", "moccasin", "Moccasin Preset Color"),
        EnumMember("NAVAJOWHITE", "navajoWhite", "Navajo White Preset Color"),
        EnumMember("NAVY", "navy", "Navy Preset Color"),
        EnumMember("OLDLACE", "oldLace", "Old Lace Preset Color"),
        EnumMember("OLIVE", "olive", "Olive Preset Color"),
        EnumMember("OLIVEDRAB", "oliveDrab", "Olive Drab Preset Color"),
        EnumMember("ORANGE", "orange", "Orange Preset Color"),
        EnumMember("ORANGERED", "orangeRed", "Orange Red Preset Color"),
        EnumMember("ORCHID", "orchid", "Orchid Preset Color"),
        EnumMember("PALEGOLDENROD", "paleGoldenrod", "Pale Goldenrod Preset Color"),
        EnumMember("PALEGREEN", "paleGreen", "Pale Green Preset Color"),
        EnumMember("PALETURQUOISE", "paleTurquoise", "Pale Turquoise Preset Color"),
        EnumMember("PALEVIOLETRED", "paleVioletRed", "Pale Violet Red Preset Color"),
        EnumMember("PAPAYAWHIP", "papayaWhip", "Papaya Whip Preset Color"),
        EnumMember("PEACHPUFF", "peachPuff", "Peach Puff Preset Color"),
        EnumMember("PERU", "peru", "Peru Preset Color"),
        EnumMember("PINK", "pink", "Pink Preset Color"),
        EnumMember("PLUM", "plum", "Plum Preset Color"),
        EnumMember("POWDERBLUE", "powderBlue", "Powder Blue Preset Color"),
        EnumMember("PURPLE", "purple", "Purple Preset Color"),
        EnumMember("RED", "red", "Red Preset Color"),
        EnumMember("ROSYBROWN", "rosyBrown", "Rosy Brown Preset Color"),
        EnumMember("ROYALBLUE", "royalBlue", "Royal Blue Preset Color"),
        EnumMember("SADDLEBROWN", "saddleBrown", "Saddle Brown Preset Color"),
        EnumMember("SALMON", "salmon", "Salmon Preset Color"),
        EnumMember("SANDYBROWN", "sandyBrown", "Sandy Brown Preset Color"),
        EnumMember("SEAGREEN", "seaGreen", "Sea Green Preset Color"),
        EnumMember("SEASHELL", "seaShell", "Sea Shell Preset Color"),
        EnumMember("SIENNA", "sienna", "Sienna Preset Color"),
        EnumMember("SILVER", "silver", "Silver Preset Color"),
        EnumMember("SKYBLUE", "skyBlue", "Sky Blue Preset Color"),
        EnumMember("SLATEBLUE", "slateBlue", "Slate Blue Preset Color"),
        EnumMember("SLATEGRAY", "slateGray", "Slate Gray Preset Color"),
        EnumMember("SNOW", "snow", "Snow Preset Color"),
        EnumMember("SPRINGGREEN", "springGreen", "Spring Green Preset Color"),
        EnumMember("STEELBLUE", "steelBlue", "Steel Blue Preset Color"),
        EnumMember("TAN", "tan", "Tan Preset Color"),
        EnumMember("TEAL", "teal", "Teal Preset Color"),
        EnumMember("THISTLE", "thistle", "Thistle Preset Color"),
        EnumMember("TOMATO", "tomato", "Tomato Preset Color"),
        EnumMember("TURQUOISE", "turquoise", "Turquoise Preset Color"),
        EnumMember("VIOLET", "violet", "Violet Preset Color"),
        EnumMember("WHEAT", "wheat", "Wheat Preset Color"),
        EnumMember("WHITE", "white", "White Preset Color"),
        EnumMember("WHITESMOKE", "whiteSmoke", "White Smoke Preset Color"),
        EnumMember("YELLOW", "yellow", "Yellow Preset Color"),
        EnumMember("YELLOWGREEN", "yellowGreen", "Yellow Green Preset Color")
    )

@alias("MSO_THEME_COLOR")
class MSO_THEME_COLOR_INDEX(XmlEnumeration):
    """
    Indicates the Office theme color, one of those shown in the color gallery
    on the formatting ribbon.

    Alias: ``MSO_THEME_COLOR``

    Example::

        from pptx.enum.dml import MSO_THEME_COLOR

        shape.fill.solid()
        shape.fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1
    """

    __ms_name__ = "MsoThemeColorIndex"

    __url__ = (
        "http://msdn.microsoft.com/en-us/library/office/ff860782(v=office.15" ").aspx"
    )

    __members__ = (
        EnumMember("NOT_THEME_COLOR", 0, "Indicates the color is not a theme color."),
        XmlMappedEnumMember(
            "ACCENT_1", 5, "accent1", "Specifies the Accent 1 theme color."
        ),
        XmlMappedEnumMember(
            "ACCENT_2", 6, "accent2", "Specifies the Accent 2 theme color."
        ),
        XmlMappedEnumMember(
            "ACCENT_3", 7, "accent3", "Specifies the Accent 3 theme color."
        ),
        XmlMappedEnumMember(
            "ACCENT_4", 8, "accent4", "Specifies the Accent 4 theme color."
        ),
        XmlMappedEnumMember(
            "ACCENT_5", 9, "accent5", "Specifies the Accent 5 theme color."
        ),
        XmlMappedEnumMember(
            "ACCENT_6", 10, "accent6", "Specifies the Accent 6 theme color."
        ),
        XmlMappedEnumMember(
            "BACKGROUND_1", 14, "bg1", "Specifies the Background 1 theme " "color."
        ),
        XmlMappedEnumMember(
            "BACKGROUND_2", 16, "bg2", "Specifies the Background 2 theme " "color."
        ),
        XmlMappedEnumMember("DARK_1", 1, "dk1", "Specifies the Dark 1 theme color."),
        XmlMappedEnumMember("DARK_2", 3, "dk2", "Specifies the Dark 2 theme color."),
        XmlMappedEnumMember(
            "FOLLOWED_HYPERLINK",
            12,
            "folHlink",
            "Specifies the theme color" " for a clicked hyperlink.",
        ),
        XmlMappedEnumMember(
            "HYPERLINK", 11, "hlink", "Specifies the theme color for a hyper" "link."
        ),
        XmlMappedEnumMember("LIGHT_1", 2, "lt1", "Specifies the Light 1 theme color."),
        XmlMappedEnumMember("LIGHT_2", 4, "lt2", "Specifies the Light 2 theme color."),
        XmlMappedEnumMember("TEXT_1", 13, "tx1", "Specifies the Text 1 theme color."),
        XmlMappedEnumMember("TEXT_2", 15, "tx2", "Specifies the Text 2 theme color."),
        ReturnValueOnlyEnumMember(
            "MIXED",
            -2,
            "Indicates multiple theme colors are used, such as " "in a group shape.",
        ),
    )
