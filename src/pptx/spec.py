# encoding: utf-8

"""Mappings from the ISO/IEC 29500 spec.

Some of these are inferred from PowerPoint application behavior
"""

from pptx.enum.shapes import MSO_SHAPE


GRAPHIC_DATA_URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
GRAPHIC_DATA_URI_OLEOBJ = "http://schemas.openxmlformats.org/presentationml/2006/ole"
GRAPHIC_DATA_URI_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"


# ============================================================================
# AutoShape type specs
# ============================================================================

autoshape_types = {
    MSO_SHAPE.ACTION_BUTTON_BACK_OR_PREVIOUS: {
        "basename": "Action Button: Back or Previous",
        "avLst": (),
    },
    MSO_SHAPE.ACTION_BUTTON_BEGINNING: {
        "basename": "Action Button: Beginning",
        "avLst": (),
    },
    MSO_SHAPE.ACTION_BUTTON_CUSTOM: {"basename": "Action Button: Custom", "avLst": ()},
    MSO_SHAPE.ACTION_BUTTON_DOCUMENT: {
        "basename": "Action Button: Document",
        "avLst": (),
    },
    MSO_SHAPE.ACTION_BUTTON_END: {"basename": "Action Button: End", "avLst": ()},
    MSO_SHAPE.ACTION_BUTTON_FORWARD_OR_NEXT: {
        "basename": "Action Button: Forward or Next",
        "avLst": (),
    },
    MSO_SHAPE.ACTION_BUTTON_HELP: {"basename": "Action Button: Help", "avLst": ()},
    MSO_SHAPE.ACTION_BUTTON_HOME: {"basename": "Action Button: Home", "avLst": ()},
    MSO_SHAPE.ACTION_BUTTON_INFORMATION: {
        "basename": "Action Button: Information",
        "avLst": (),
    },
    MSO_SHAPE.ACTION_BUTTON_MOVIE: {"basename": "Action Button: Movie", "avLst": ()},
    MSO_SHAPE.ACTION_BUTTON_RETURN: {"basename": "Action Button: Return", "avLst": ()},
    MSO_SHAPE.ACTION_BUTTON_SOUND: {"basename": "Action Button: Sound", "avLst": ()},
    MSO_SHAPE.ARC: {"basename": "Arc", "avLst": (("adj1", 16200000), ("adj2", 0))},
    MSO_SHAPE.BALLOON: {
        "basename": "Rounded Rectangular Callout",
        "avLst": (("adj1", -20833), ("adj2", 62500), ("adj3", 16667)),
    },
    MSO_SHAPE.BENT_ARROW: {
        "basename": "Bent Arrow",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 43750)),
    },
    MSO_SHAPE.BENT_UP_ARROW: {
        "basename": "Bent-Up Arrow",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000)),
    },
    MSO_SHAPE.BEVEL: {"basename": "Bevel", "avLst": (("adj", 12500),)},
    MSO_SHAPE.BLOCK_ARC: {
        "basename": "Block Arc",
        "avLst": (("adj1", 10800000), ("adj2", 0), ("adj3", 25000)),
    },
    MSO_SHAPE.CAN: {"basename": "Can", "avLst": (("adj", 25000),)},
    MSO_SHAPE.CHART_PLUS: {"basename": "Chart Plus", "avLst": ()},
    MSO_SHAPE.CHART_STAR: {"basename": "Chart Star", "avLst": ()},
    MSO_SHAPE.CHART_X: {"basename": "Chart X", "avLst": ()},
    MSO_SHAPE.CHEVRON: {"basename": "Chevron", "avLst": (("adj", 50000),)},
    MSO_SHAPE.CHORD: {
        "basename": "Chord",
        "avLst": (("adj1", 2700000), ("adj2", 16200000)),
    },
    MSO_SHAPE.CIRCULAR_ARROW: {
        "basename": "Circular Arrow",
        "avLst": (
            ("adj1", 12500),
            ("adj2", 1142319),
            ("adj3", 20457681),
            ("adj4", 10800000),
            ("adj5", 12500),
        ),
    },
    MSO_SHAPE.CLOUD: {"basename": "Cloud", "avLst": ()},
    MSO_SHAPE.CLOUD_CALLOUT: {
        "basename": "Cloud Callout",
        "avLst": (("adj1", -20833), ("adj2", 62500)),
    },
    MSO_SHAPE.CORNER: {
        "basename": "Corner",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.CORNER_TABS: {"basename": "Corner Tabs", "avLst": ()},
    MSO_SHAPE.CROSS: {"basename": "Cross", "avLst": (("adj", 25000),)},
    MSO_SHAPE.CUBE: {"basename": "Cube", "avLst": (("adj", 25000),)},
    MSO_SHAPE.CURVED_DOWN_ARROW: {
        "basename": "Curved Down Arrow",
        "avLst": (("adj1", 25000), ("adj2", 50000), ("adj3", 25000)),
    },
    MSO_SHAPE.CURVED_DOWN_RIBBON: {
        "basename": "Curved Down Ribbon",
        "avLst": (("adj1", 25000), ("adj2", 50000), ("adj3", 12500)),
    },
    MSO_SHAPE.CURVED_LEFT_ARROW: {
        "basename": "Curved Left Arrow",
        "avLst": (("adj1", 25000), ("adj2", 50000), ("adj3", 25000)),
    },
    MSO_SHAPE.CURVED_RIGHT_ARROW: {
        "basename": "Curved Right Arrow",
        "avLst": (("adj1", 25000), ("adj2", 50000), ("adj3", 25000)),
    },
    MSO_SHAPE.CURVED_UP_ARROW: {
        "basename": "Curved Up Arrow",
        "avLst": (("adj1", 25000), ("adj2", 50000), ("adj3", 25000)),
    },
    MSO_SHAPE.CURVED_UP_RIBBON: {
        "basename": "Curved Up Ribbon",
        "avLst": (("adj1", 25000), ("adj2", 50000), ("adj3", 12500)),
    },
    MSO_SHAPE.DECAGON: {"basename": "Decagon", "avLst": (("vf", 105146),)},
    MSO_SHAPE.DIAGONAL_STRIPE: {
        "basename": "Diagonal Stripe",
        "avLst": (("adj", 50000),),
    },
    MSO_SHAPE.DIAMOND: {"basename": "Diamond", "avLst": ()},
    MSO_SHAPE.DODECAGON: {"basename": "Dodecagon", "avLst": ()},
    MSO_SHAPE.DONUT: {"basename": "Donut", "avLst": (("adj", 25000),)},
    MSO_SHAPE.DOUBLE_BRACE: {"basename": "Double Brace", "avLst": (("adj", 8333),)},
    MSO_SHAPE.DOUBLE_BRACKET: {
        "basename": "Double Bracket",
        "avLst": (("adj", 16667),),
    },
    MSO_SHAPE.DOUBLE_WAVE: {
        "basename": "Double Wave",
        "avLst": (("adj1", 6250), ("adj2", 0)),
    },
    MSO_SHAPE.DOWN_ARROW: {
        "basename": "Down Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.DOWN_ARROW_CALLOUT: {
        "basename": "Down Arrow Callout",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977)),
    },
    MSO_SHAPE.DOWN_RIBBON: {
        "basename": "Down Ribbon",
        "avLst": (("adj1", 16667), ("adj2", 50000)),
    },
    MSO_SHAPE.EXPLOSION1: {"basename": "Explosion", "avLst": ()},
    MSO_SHAPE.EXPLOSION2: {"basename": "Explosion", "avLst": ()},
    MSO_SHAPE.FLOWCHART_ALTERNATE_PROCESS: {
        "basename": "Alternate process",
        "avLst": (),
    },
    MSO_SHAPE.FLOWCHART_CARD: {"basename": "Card", "avLst": ()},
    MSO_SHAPE.FLOWCHART_COLLATE: {"basename": "Collate", "avLst": ()},
    MSO_SHAPE.FLOWCHART_CONNECTOR: {"basename": "Connector", "avLst": ()},
    MSO_SHAPE.FLOWCHART_DATA: {"basename": "Data", "avLst": ()},
    MSO_SHAPE.FLOWCHART_DECISION: {"basename": "Decision", "avLst": ()},
    MSO_SHAPE.FLOWCHART_DELAY: {"basename": "Delay", "avLst": ()},
    MSO_SHAPE.FLOWCHART_DIRECT_ACCESS_STORAGE: {
        "basename": "Direct Access Storage",
        "avLst": (),
    },
    MSO_SHAPE.FLOWCHART_DISPLAY: {"basename": "Display", "avLst": ()},
    MSO_SHAPE.FLOWCHART_DOCUMENT: {"basename": "Document", "avLst": ()},
    MSO_SHAPE.FLOWCHART_EXTRACT: {"basename": "Extract", "avLst": ()},
    MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE: {"basename": "Internal Storage", "avLst": ()},
    MSO_SHAPE.FLOWCHART_MAGNETIC_DISK: {"basename": "Magnetic Disk", "avLst": ()},
    MSO_SHAPE.FLOWCHART_MANUAL_INPUT: {"basename": "Manual Input", "avLst": ()},
    MSO_SHAPE.FLOWCHART_MANUAL_OPERATION: {"basename": "Manual Operation", "avLst": ()},
    MSO_SHAPE.FLOWCHART_MERGE: {"basename": "Merge", "avLst": ()},
    MSO_SHAPE.FLOWCHART_MULTIDOCUMENT: {"basename": "Multidocument", "avLst": ()},
    MSO_SHAPE.FLOWCHART_OFFLINE_STORAGE: {"basename": "Offline Storage", "avLst": ()},
    MSO_SHAPE.FLOWCHART_OFFPAGE_CONNECTOR: {
        "basename": "Off-page Connector",
        "avLst": (),
    },
    MSO_SHAPE.FLOWCHART_OR: {"basename": "Or", "avLst": ()},
    MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS: {
        "basename": "Predefined Process",
        "avLst": (),
    },
    MSO_SHAPE.FLOWCHART_PREPARATION: {"basename": "Preparation", "avLst": ()},
    MSO_SHAPE.FLOWCHART_PROCESS: {"basename": "Process", "avLst": ()},
    MSO_SHAPE.FLOWCHART_PUNCHED_TAPE: {"basename": "Punched Tape", "avLst": ()},
    MSO_SHAPE.FLOWCHART_SEQUENTIAL_ACCESS_STORAGE: {
        "basename": "Sequential Access Storage",
        "avLst": (),
    },
    MSO_SHAPE.FLOWCHART_SORT: {"basename": "Sort", "avLst": ()},
    MSO_SHAPE.FLOWCHART_STORED_DATA: {"basename": "Stored Data", "avLst": ()},
    MSO_SHAPE.FLOWCHART_SUMMING_JUNCTION: {"basename": "Summing Junction", "avLst": ()},
    MSO_SHAPE.FLOWCHART_TERMINATOR: {"basename": "Terminator", "avLst": ()},
    MSO_SHAPE.FOLDED_CORNER: {"basename": "Folded Corner", "avLst": ()},
    MSO_SHAPE.FRAME: {"basename": "Frame", "avLst": (("adj1", 12500),)},
    MSO_SHAPE.FUNNEL: {"basename": "Funnel", "avLst": ()},
    MSO_SHAPE.GEAR_6: {
        "basename": "Gear 6",
        "avLst": (("adj1", 15000), ("adj2", 3526)),
    },
    MSO_SHAPE.GEAR_9: {
        "basename": "Gear 9",
        "avLst": (("adj1", 10000), ("adj2", 1763)),
    },
    MSO_SHAPE.HALF_FRAME: {
        "basename": "Half Frame",
        "avLst": (("adj1", 33333), ("adj2", 33333)),
    },
    MSO_SHAPE.HEART: {"basename": "Heart", "avLst": ()},
    MSO_SHAPE.HEPTAGON: {
        "basename": "Heptagon",
        "avLst": (("hf", 102572), ("vf", 105210)),
    },
    MSO_SHAPE.HEXAGON: {
        "basename": "Hexagon",
        "avLst": (("adj", 25000), ("vf", 115470)),
    },
    MSO_SHAPE.HORIZONTAL_SCROLL: {
        "basename": "Horizontal Scroll",
        "avLst": (("adj", 12500),),
    },
    MSO_SHAPE.ISOSCELES_TRIANGLE: {
        "basename": "Isosceles Triangle",
        "avLst": (("adj", 50000),),
    },
    MSO_SHAPE.LEFT_ARROW: {
        "basename": "Left Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.LEFT_ARROW_CALLOUT: {
        "basename": "Left Arrow Callout",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977)),
    },
    MSO_SHAPE.LEFT_BRACE: {
        "basename": "Left Brace",
        "avLst": (("adj1", 8333), ("adj2", 50000)),
    },
    MSO_SHAPE.LEFT_BRACKET: {"basename": "Left Bracket", "avLst": (("adj", 8333),)},
    MSO_SHAPE.LEFT_CIRCULAR_ARROW: {
        "basename": "Left Circular Arrow",
        "avLst": (
            ("adj1", 12500),
            ("adj2", -1142319),
            ("adj3", 1142319),
            ("adj4", 10800000),
            ("adj5", 12500),
        ),
    },
    MSO_SHAPE.LEFT_RIGHT_ARROW: {
        "basename": "Left-Right Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.LEFT_RIGHT_ARROW_CALLOUT: {
        "basename": "Left-Right Arrow Callout",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 48123)),
    },
    MSO_SHAPE.LEFT_RIGHT_CIRCULAR_ARROW: {
        "basename": "Left Right Circular Arrow",
        "avLst": (
            ("adj1", 12500),
            ("adj2", 1142319),
            ("adj3", 20457681),
            ("adj4", 11942319),
            ("adj5", 12500),
        ),
    },
    MSO_SHAPE.LEFT_RIGHT_RIBBON: {
        "basename": "Left Right Ribbon",
        "avLst": (("adj1", 50000), ("adj2", 50000), ("adj3", 16667)),
    },
    MSO_SHAPE.LEFT_RIGHT_UP_ARROW: {
        "basename": "Left-Right-Up Arrow",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000)),
    },
    MSO_SHAPE.LEFT_UP_ARROW: {
        "basename": "Left-Up Arrow",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000)),
    },
    MSO_SHAPE.LIGHTNING_BOLT: {"basename": "Lightning Bolt", "avLst": ()},
    MSO_SHAPE.LINE_CALLOUT_1: {
        "basename": "Line Callout 1",
        "avLst": (("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333)),
    },
    MSO_SHAPE.LINE_CALLOUT_1_ACCENT_BAR: {
        "basename": "Line Callout 1 (Accent Bar)",
        "avLst": (("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333)),
    },
    MSO_SHAPE.LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR: {
        "basename": "Line Callout 1 (Border and Accent Bar)",
        "avLst": (("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333)),
    },
    MSO_SHAPE.LINE_CALLOUT_1_NO_BORDER: {
        "basename": "Line Callout 1 (No Border)",
        "avLst": (("adj1", 18750), ("adj2", -8333), ("adj3", 112500), ("adj4", -38333)),
    },
    MSO_SHAPE.LINE_CALLOUT_2: {
        "basename": "Line Callout 2",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_2_ACCENT_BAR: {
        "basename": "Line Callout 2 (Accent Bar)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR: {
        "basename": "Line Callout 2 (Border and Accent Bar)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_2_NO_BORDER: {
        "basename": "Line Callout 2 (No Border)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 112500),
            ("adj6", -46667),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_3: {
        "basename": "Line Callout 3",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_3_ACCENT_BAR: {
        "basename": "Line Callout 3 (Accent Bar)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR: {
        "basename": "Line Callout 3 (Border and Accent Bar)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_3_NO_BORDER: {
        "basename": "Line Callout 3 (No Border)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_4: {
        "basename": "Line Callout 3",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_4_ACCENT_BAR: {
        "basename": "Line Callout 3 (Accent Bar)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR: {
        "basename": "Line Callout 3 (Border and Accent Bar)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_CALLOUT_4_NO_BORDER: {
        "basename": "Line Callout 3 (No Border)",
        "avLst": (
            ("adj1", 18750),
            ("adj2", -8333),
            ("adj3", 18750),
            ("adj4", -16667),
            ("adj5", 100000),
            ("adj6", -16667),
            ("adj7", 112963),
            ("adj8", -8333),
        ),
    },
    MSO_SHAPE.LINE_INVERSE: {"basename": "Straight Connector", "avLst": ()},
    MSO_SHAPE.MATH_DIVIDE: {
        "basename": "Division",
        "avLst": (("adj1", 23520), ("adj2", 5880), ("adj3", 11760)),
    },
    MSO_SHAPE.MATH_EQUAL: {
        "basename": "Equal",
        "avLst": (("adj1", 23520), ("adj2", 11760)),
    },
    MSO_SHAPE.MATH_MINUS: {"basename": "Minus", "avLst": (("adj1", 23520),)},
    MSO_SHAPE.MATH_MULTIPLY: {"basename": "Multiply", "avLst": (("adj1", 23520),)},
    MSO_SHAPE.MATH_NOT_EQUAL: {
        "basename": "Not Equal",
        "avLst": (("adj1", 23520), ("adj2", 6600000), ("adj3", 11760)),
    },
    MSO_SHAPE.MATH_PLUS: {"basename": "Plus", "avLst": (("adj1", 23520),)},
    MSO_SHAPE.MOON: {"basename": "Moon", "avLst": (("adj", 50000),)},
    MSO_SHAPE.NON_ISOSCELES_TRAPEZOID: {
        "basename": "Non-isosceles Trapezoid",
        "avLst": (("adj1", 25000), ("adj2", 25000)),
    },
    MSO_SHAPE.NOTCHED_RIGHT_ARROW: {
        "basename": "Notched Right Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.NO_SYMBOL: {"basename": '"No" Symbol', "avLst": (("adj", 18750),)},
    MSO_SHAPE.OCTAGON: {"basename": "Octagon", "avLst": (("adj", 29289),)},
    MSO_SHAPE.OVAL: {"basename": "Oval", "avLst": ()},
    MSO_SHAPE.OVAL_CALLOUT: {
        "basename": "Oval Callout",
        "avLst": (("adj1", -20833), ("adj2", 62500)),
    },
    MSO_SHAPE.PARALLELOGRAM: {"basename": "Parallelogram", "avLst": (("adj", 25000),)},
    MSO_SHAPE.PENTAGON: {"basename": "Pentagon", "avLst": (("adj", 50000),)},
    MSO_SHAPE.PIE: {"basename": "Pie", "avLst": (("adj1", 0), ("adj2", 16200000))},
    MSO_SHAPE.PIE_WEDGE: {"basename": "Pie", "avLst": ()},
    MSO_SHAPE.PLAQUE: {"basename": "Plaque", "avLst": (("adj", 16667),)},
    MSO_SHAPE.PLAQUE_TABS: {"basename": "Plaque Tabs", "avLst": ()},
    MSO_SHAPE.QUAD_ARROW: {
        "basename": "Quad Arrow",
        "avLst": (("adj1", 22500), ("adj2", 22500), ("adj3", 22500)),
    },
    MSO_SHAPE.QUAD_ARROW_CALLOUT: {
        "basename": "Quad Arrow Callout",
        "avLst": (("adj1", 18515), ("adj2", 18515), ("adj3", 18515), ("adj4", 48123)),
    },
    MSO_SHAPE.RECTANGLE: {"basename": "Rectangle", "avLst": ()},
    MSO_SHAPE.RECTANGULAR_CALLOUT: {
        "basename": "Rectangular Callout",
        "avLst": (("adj1", -20833), ("adj2", 62500)),
    },
    MSO_SHAPE.REGULAR_PENTAGON: {
        "basename": "Regular Pentagon",
        "avLst": (("hf", 105146), ("vf", 110557)),
    },
    MSO_SHAPE.RIGHT_ARROW: {
        "basename": "Right Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.RIGHT_ARROW_CALLOUT: {
        "basename": "Right Arrow Callout",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977)),
    },
    MSO_SHAPE.RIGHT_BRACE: {
        "basename": "Right Brace",
        "avLst": (("adj1", 8333), ("adj2", 50000)),
    },
    MSO_SHAPE.RIGHT_BRACKET: {"basename": "Right Bracket", "avLst": (("adj", 8333),)},
    MSO_SHAPE.RIGHT_TRIANGLE: {"basename": "Right Triangle", "avLst": ()},
    MSO_SHAPE.ROUNDED_RECTANGLE: {
        "basename": "Rounded Rectangle",
        "avLst": (("adj", 16667),),
    },
    MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT: {
        "basename": "Rounded Rectangular Callout",
        "avLst": (("adj1", -20833), ("adj2", 62500), ("adj3", 16667)),
    },
    MSO_SHAPE.ROUND_1_RECTANGLE: {
        "basename": "Round Single Corner Rectangle",
        "avLst": (("adj", 16667),),
    },
    MSO_SHAPE.ROUND_2_DIAG_RECTANGLE: {
        "basename": "Round Diagonal Corner Rectangle",
        "avLst": (("adj1", 16667), ("adj2", 0)),
    },
    MSO_SHAPE.ROUND_2_SAME_RECTANGLE: {
        "basename": "Round Same Side Corner Rectangle",
        "avLst": (("adj1", 16667), ("adj2", 0)),
    },
    MSO_SHAPE.SMILEY_FACE: {"basename": "Smiley Face", "avLst": (("adj", 4653),)},
    MSO_SHAPE.SNIP_1_RECTANGLE: {
        "basename": "Snip Single Corner Rectangle",
        "avLst": (("adj", 16667),),
    },
    MSO_SHAPE.SNIP_2_DIAG_RECTANGLE: {
        "basename": "Snip Diagonal Corner Rectangle",
        "avLst": (("adj1", 0), ("adj2", 16667)),
    },
    MSO_SHAPE.SNIP_2_SAME_RECTANGLE: {
        "basename": "Snip Same Side Corner Rectangle",
        "avLst": (("adj1", 16667), ("adj2", 0)),
    },
    MSO_SHAPE.SNIP_ROUND_RECTANGLE: {
        "basename": "Snip and Round Single Corner Rectangle",
        "avLst": (("adj1", 16667), ("adj2", 16667)),
    },
    MSO_SHAPE.SQUARE_TABS: {"basename": "Square Tabs", "avLst": ()},
    MSO_SHAPE.STAR_10_POINT: {
        "basename": "10-Point Star",
        "avLst": (("adj", 42533), ("hf", 105146)),
    },
    MSO_SHAPE.STAR_12_POINT: {"basename": "12-Point Star", "avLst": (("adj", 37500),)},
    MSO_SHAPE.STAR_16_POINT: {"basename": "16-Point Star", "avLst": (("adj", 37500),)},
    MSO_SHAPE.STAR_24_POINT: {"basename": "24-Point Star", "avLst": (("adj", 37500),)},
    MSO_SHAPE.STAR_32_POINT: {"basename": "32-Point Star", "avLst": (("adj", 37500),)},
    MSO_SHAPE.STAR_4_POINT: {"basename": "4-Point Star", "avLst": (("adj", 12500),)},
    MSO_SHAPE.STAR_5_POINT: {
        "basename": "5-Point Star",
        "avLst": (("adj", 19098), ("hf", 105146), ("vf", 110557)),
    },
    MSO_SHAPE.STAR_6_POINT: {
        "basename": "6-Point Star",
        "avLst": (("adj", 28868), ("hf", 115470)),
    },
    MSO_SHAPE.STAR_7_POINT: {
        "basename": "7-Point Star",
        "avLst": (("adj", 34601), ("hf", 102572), ("vf", 105210)),
    },
    MSO_SHAPE.STAR_8_POINT: {"basename": "8-Point Star", "avLst": (("adj", 37500),)},
    MSO_SHAPE.STRIPED_RIGHT_ARROW: {
        "basename": "Striped Right Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.SUN: {"basename": "Sun", "avLst": (("adj", 25000),)},
    MSO_SHAPE.SWOOSH_ARROW: {
        "basename": "Swoosh Arrow",
        "avLst": (("adj1", 25000), ("adj2", 16667)),
    },
    MSO_SHAPE.TEAR: {"basename": "Teardrop", "avLst": (("adj", 100000),)},
    MSO_SHAPE.TRAPEZOID: {"basename": "Trapezoid", "avLst": (("adj", 25000),)},
    MSO_SHAPE.UP_ARROW: {
        "basename": "Up Arrow",
        "avLst": (("adj1", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.UP_ARROW_CALLOUT: {
        "basename": "Up Arrow Callout",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 64977)),
    },
    MSO_SHAPE.UP_DOWN_ARROW: {
        "basename": "Up-Down Arrow",
        "avLst": (("adj1", 50000), ("adj1", 50000), ("adj2", 50000), ("adj2", 50000)),
    },
    MSO_SHAPE.UP_DOWN_ARROW_CALLOUT: {
        "basename": "Up-Down Arrow Callout",
        "avLst": (("adj1", 25000), ("adj2", 25000), ("adj3", 25000), ("adj4", 48123)),
    },
    MSO_SHAPE.UP_RIBBON: {
        "basename": "Up Ribbon",
        "avLst": (("adj1", 16667), ("adj2", 50000)),
    },
    MSO_SHAPE.U_TURN_ARROW: {
        "basename": "U-Turn Arrow",
        "avLst": (
            ("adj1", 25000),
            ("adj2", 25000),
            ("adj3", 25000),
            ("adj4", 43750),
            ("adj5", 75000),
        ),
    },
    MSO_SHAPE.VERTICAL_SCROLL: {
        "basename": "Vertical Scroll",
        "avLst": (("adj", 12500),),
    },
    MSO_SHAPE.WAVE: {"basename": "Wave", "avLst": (("adj1", 12500), ("adj2", 0))},
}
