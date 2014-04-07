# encoding: utf-8

"""
Constant values from the ISO/IEC 29500 spec that are needed for XML
generation and packaging, and a utility function or two for accessing some of
them.
"""

from __future__ import absolute_import

from pptx.enum.shapes import MSO_SHAPE


# ============================================================================
# AutoShape type specs
# ============================================================================

autoshape_types = {
    MSO_SHAPE.ACTION_BUTTON_BACK_OR_PREVIOUS: {
        'basename': 'Action Button: Back or Previous',
        'prst':     'actionButtonBackPrevious',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_BEGINNING: {
        'basename': 'Action Button: Beginning',
        'prst':     'actionButtonBeginning',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_CUSTOM: {
        'basename': 'Action Button: Custom',
        'prst':     'actionButtonBlank',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_DOCUMENT: {
        'basename': 'Action Button: Document',
        'prst':     'actionButtonDocument',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_END: {
        'basename': 'Action Button: End',
        'prst':     'actionButtonEnd',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_FORWARD_OR_NEXT: {
        'basename': 'Action Button: Forward or Next',
        'prst':     'actionButtonForwardNext',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_HELP: {
        'basename': 'Action Button: Help',
        'prst':     'actionButtonHelp',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_HOME: {
        'basename': 'Action Button: Home',
        'prst':     'actionButtonHome',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_INFORMATION: {
        'basename': 'Action Button: Information',
        'prst':     'actionButtonInformation',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_MOVIE: {
        'basename': 'Action Button: Movie',
        'prst':     'actionButtonMovie',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_RETURN: {
        'basename': 'Action Button: Return',
        'prst':     'actionButtonReturn',
        'avLst':    ()
    },
    MSO_SHAPE.ACTION_BUTTON_SOUND: {
        'basename': 'Action Button: Sound',
        'prst':     'actionButtonSound',
        'avLst':    ()
    },
    MSO_SHAPE.ARC: {
        'basename': 'Arc',
        'prst':     'arc',
        'avLst':    (
            ('adj1', 16200000),
            ('adj2', 0),
        )
    },
    MSO_SHAPE.BALLOON: {
        'basename': 'Rounded Rectangular Callout',
        'prst':     'wedgeRoundRectCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
            ('adj3', 16667),
        )
    },
    MSO_SHAPE.BENT_ARROW: {
        'basename': 'Bent Arrow',
        'prst':     'bentArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 43750),
        )
    },
    MSO_SHAPE.BENT_UP_ARROW: {
        'basename': 'Bent-Up Arrow',
        'prst':     'bentUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.BEVEL: {
        'basename': 'Bevel',
        'prst':     'bevel',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MSO_SHAPE.BLOCK_ARC: {
        'basename': 'Block Arc',
        'prst':     'blockArc',
        'avLst':    (
            ('adj1', 10800000),
            ('adj2', 0),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.CAN: {
        'basename': 'Can',
        'prst':     'can',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.CHART_PLUS: {
        'basename': 'Chart Plus',
        'prst':     'chartPlus',
        'avLst':    ()
    },
    MSO_SHAPE.CHART_STAR: {
        'basename': 'Chart Star',
        'prst':     'chartStar',
        'avLst':    ()
    },
    MSO_SHAPE.CHART_X: {
        'basename': 'Chart X',
        'prst':     'chartX',
        'avLst':    ()
    },
    MSO_SHAPE.CHEVRON: {
        'basename': 'Chevron',
        'prst':     'chevron',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MSO_SHAPE.CHORD: {
        'basename': 'Chord',
        'prst':     'chord',
        'avLst':    (
            ('adj1', 2700000),
            ('adj2', 16200000),
        )
    },
    MSO_SHAPE.CIRCULAR_ARROW: {
        'basename': 'Circular Arrow',
        'prst':     'circularArrow',
        'avLst':    (
            ('adj1', 12500),
            ('adj2', 1142319),
            ('adj3', 20457681),
            ('adj4', 10800000),
            ('adj5', 12500),
        )
    },
    MSO_SHAPE.CLOUD: {
        'basename': 'Cloud',
        'prst':     'cloud',
        'avLst':    ()
    },
    MSO_SHAPE.CLOUD_CALLOUT: {
        'basename': 'Cloud Callout',
        'prst':     'cloudCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
        )
    },
    MSO_SHAPE.CORNER: {
        'basename': 'Corner',
        'prst':     'corner',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.CORNER_TABS: {
        'basename': 'Corner Tabs',
        'prst':     'cornerTabs',
        'avLst':    ()
    },
    MSO_SHAPE.CROSS: {
        'basename': 'Cross',
        'prst':     'plus',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.CUBE: {
        'basename': 'Cube',
        'prst':     'cube',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.CURVED_DOWN_ARROW: {
        'basename': 'Curved Down Arrow',
        'prst':     'curvedDownArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.CURVED_DOWN_RIBBON: {
        'basename': 'Curved Down Ribbon',
        'prst':     'ellipseRibbon',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 12500),
        )
    },
    MSO_SHAPE.CURVED_LEFT_ARROW: {
        'basename': 'Curved Left Arrow',
        'prst':     'curvedLeftArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.CURVED_RIGHT_ARROW: {
        'basename': 'Curved Right Arrow',
        'prst':     'curvedRightArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.CURVED_UP_ARROW: {
        'basename': 'Curved Up Arrow',
        'prst':     'curvedUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.CURVED_UP_RIBBON: {
        'basename': 'Curved Up Ribbon',
        'prst':     'ellipseRibbon2',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 12500),
        )
    },
    MSO_SHAPE.DECAGON: {
        'basename': 'Decagon',
        'prst':     'decagon',
        'avLst':    (
            ('vf', 105146),
        )
    },
    MSO_SHAPE.DIAGONAL_STRIPE: {
        'basename': 'Diagonal Stripe',
        'prst':     'diagStripe',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MSO_SHAPE.DIAMOND: {
        'basename': 'Diamond',
        'prst':     'diamond',
        'avLst':    ()
    },
    MSO_SHAPE.DODECAGON: {
        'basename': 'Dodecagon',
        'prst':     'dodecagon',
        'avLst':    ()
    },
    MSO_SHAPE.DONUT: {
        'basename': 'Donut',
        'prst':     'donut',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.DOUBLE_BRACE: {
        'basename': 'Double Brace',
        'prst':     'bracePair',
        'avLst':    (
            ('adj', 8333),
        )
    },
    MSO_SHAPE.DOUBLE_BRACKET: {
        'basename': 'Double Bracket',
        'prst':     'bracketPair',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MSO_SHAPE.DOUBLE_WAVE: {
        'basename': 'Double Wave',
        'prst':     'doubleWave',
        'avLst':    (
            ('adj1', 6250),
            ('adj2', 0),
        )
    },
    MSO_SHAPE.DOWN_ARROW: {
        'basename': 'Down Arrow',
        'prst':     'downArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.DOWN_ARROW_CALLOUT: {
        'basename': 'Down Arrow Callout',
        'prst':     'downArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MSO_SHAPE.DOWN_RIBBON: {
        'basename': 'Down Ribbon',
        'prst':     'ribbon',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.EXPLOSION1: {
        'basename': 'Explosion',
        'prst':     'irregularSeal1',
        'avLst':    ()
    },
    MSO_SHAPE.EXPLOSION2: {
        'basename': 'Explosion',
        'prst':     'irregularSeal2',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_ALTERNATE_PROCESS: {
        'basename': 'Alternate process',
        'prst':     'flowChartAlternateProcess',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_CARD: {
        'basename': 'Card',
        'prst':     'flowChartPunchedCard',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_COLLATE: {
        'basename': 'Collate',
        'prst':     'flowChartCollate',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_CONNECTOR: {
        'basename': 'Connector',
        'prst':     'flowChartConnector',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_DATA: {
        'basename': 'Data',
        'prst':     'flowChartInputOutput',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_DECISION: {
        'basename': 'Decision',
        'prst':     'flowChartDecision',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_DELAY: {
        'basename': 'Delay',
        'prst':     'flowChartDelay',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_DIRECT_ACCESS_STORAGE: {
        'basename': 'Direct Access Storage',
        'prst':     'flowChartMagneticDrum',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_DISPLAY: {
        'basename': 'Display',
        'prst':     'flowChartDisplay',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_DOCUMENT: {
        'basename': 'Document',
        'prst':     'flowChartDocument',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_EXTRACT: {
        'basename': 'Extract',
        'prst':     'flowChartExtract',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE: {
        'basename': 'Internal Storage',
        'prst':     'flowChartInternalStorage',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_MAGNETIC_DISK: {
        'basename': 'Magnetic Disk',
        'prst':     'flowChartMagneticDisk',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_MANUAL_INPUT: {
        'basename': 'Manual Input',
        'prst':     'flowChartManualInput',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_MANUAL_OPERATION: {
        'basename': 'Manual Operation',
        'prst':     'flowChartManualOperation',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_MERGE: {
        'basename': 'Merge',
        'prst':     'flowChartMerge',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_MULTIDOCUMENT: {
        'basename': 'Multidocument',
        'prst':     'flowChartMultidocument',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_OFFLINE_STORAGE: {
        'basename': 'Offline Storage',
        'prst':     'flowChartOfflineStorage',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_OFFPAGE_CONNECTOR: {
        'basename': 'Off-page Connector',
        'prst':     'flowChartOffpageConnector',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_OR: {
        'basename': 'Or',
        'prst':     'flowChartOr',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS: {
        'basename': 'Predefined Process',
        'prst':     'flowChartPredefinedProcess',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_PREPARATION: {
        'basename': 'Preparation',
        'prst':     'flowChartPreparation',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_PROCESS: {
        'basename': 'Process',
        'prst':     'flowChartProcess',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_PUNCHED_TAPE: {
        'basename': 'Punched Tape',
        'prst':     'flowChartPunchedTape',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_SEQUENTIAL_ACCESS_STORAGE: {
        'basename': 'Sequential Access Storage',
        'prst':     'flowChartMagneticTape',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_SORT: {
        'basename': 'Sort',
        'prst':     'flowChartSort',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_STORED_DATA: {
        'basename': 'Stored Data',
        'prst':     'flowChartOnlineStorage',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_SUMMING_JUNCTION: {
        'basename': 'Summing Junction',
        'prst':     'flowChartSummingJunction',
        'avLst':    ()
    },
    MSO_SHAPE.FLOWCHART_TERMINATOR: {
        'basename': 'Terminator',
        'prst':     'flowChartTerminator',
        'avLst':    ()
    },
    MSO_SHAPE.FOLDED_CORNER: {
        'basename': 'Folded Corner',
        'prst':     'foldedCorner',
        'avLst':    ()
    },
    MSO_SHAPE.FRAME: {
        'basename': 'Frame',
        'prst':     'frame',
        'avLst':    (
            ('adj1', 12500),
        )
    },
    MSO_SHAPE.FUNNEL: {
        'basename': 'Funnel',
        'prst':     'funnel',
        'avLst':    ()
    },
    MSO_SHAPE.GEAR_6: {
        'basename': 'Gear 6',
        'prst':     'gear6',
        'avLst':    (
            ('adj1', 15000),
            ('adj2', 3526),
        )
    },
    MSO_SHAPE.GEAR_9: {
        'basename': 'Gear 9',
        'prst':     'gear9',
        'avLst':    (
            ('adj1', 10000),
            ('adj2', 1763),
        )
    },
    MSO_SHAPE.HALF_FRAME: {
        'basename': 'Half Frame',
        'prst':     'halfFrame',
        'avLst':    (
            ('adj1', 33333),
            ('adj2', 33333),
        )
    },
    MSO_SHAPE.HEART: {
        'basename': 'Heart',
        'prst':     'heart',
        'avLst':    ()
    },
    MSO_SHAPE.HEPTAGON: {
        'basename': 'Heptagon',
        'prst':     'heptagon',
        'avLst':    (
            ('hf', 102572),
            ('vf', 105210),
        )
    },
    MSO_SHAPE.HEXAGON: {
        'basename': 'Hexagon',
        'prst':     'hexagon',
        'avLst':    (
            ('adj', 25000),
            ('vf', 115470),
        )
    },
    MSO_SHAPE.HORIZONTAL_SCROLL: {
        'basename': 'Horizontal Scroll',
        'prst':     'horizontalScroll',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MSO_SHAPE.ISOSCELES_TRIANGLE: {
        'basename': 'Isosceles Triangle',
        'prst':     'triangle',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MSO_SHAPE.LEFT_ARROW: {
        'basename': 'Left Arrow',
        'prst':     'leftArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.LEFT_ARROW_CALLOUT: {
        'basename': 'Left Arrow Callout',
        'prst':     'leftArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MSO_SHAPE.LEFT_BRACE: {
        'basename': 'Left Brace',
        'prst':     'leftBrace',
        'avLst':    (
            ('adj1', 8333),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.LEFT_BRACKET: {
        'basename': 'Left Bracket',
        'prst':     'leftBracket',
        'avLst':    (
            ('adj', 8333),
        )
    },
    MSO_SHAPE.LEFT_CIRCULAR_ARROW: {
        'basename': 'Left Circular Arrow',
        'prst':     'leftCircularArrow',
        'avLst':    (
            ('adj1', 12500),
            ('adj2', -1142319),
            ('adj3', 1142319),
            ('adj4', 10800000),
            ('adj5', 12500),
        )
    },
    MSO_SHAPE.LEFT_RIGHT_ARROW: {
        'basename': 'Left-Right Arrow',
        'prst':     'leftRightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.LEFT_RIGHT_ARROW_CALLOUT: {
        'basename': 'Left-Right Arrow Callout',
        'prst':     'leftRightArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 48123),
        )
    },
    MSO_SHAPE.LEFT_RIGHT_CIRCULAR_ARROW: {
        'basename': 'Left Right Circular Arrow',
        'prst':     'leftRightCircularArrow',
        'avLst':    (
            ('adj1', 12500),
            ('adj2', 1142319),
            ('adj3', 20457681),
            ('adj4', 11942319),
            ('adj5', 12500),
        )
    },
    MSO_SHAPE.LEFT_RIGHT_RIBBON: {
        'basename': 'Left Right Ribbon',
        'prst':     'leftRightRibbon',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
            ('adj3', 16667),
        )
    },
    MSO_SHAPE.LEFT_RIGHT_UP_ARROW: {
        'basename': 'Left-Right-Up Arrow',
        'prst':     'leftRightUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.LEFT_UP_ARROW: {
        'basename': 'Left-Up Arrow',
        'prst':     'leftUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
        )
    },
    MSO_SHAPE.LIGHTNING_BOLT: {
        'basename': 'Lightning Bolt',
        'prst':     'lightningBolt',
        'avLst':    ()
    },
    MSO_SHAPE.LINE_CALLOUT_1: {
        'basename': 'Line Callout 1',
        'prst':     'borderCallout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_1_ACCENT_BAR: {
        'basename': 'Line Callout 1 (Accent Bar)',
        'prst':     'accentCallout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 1 (Border and Accent Bar)',
        'prst':     'accentBorderCallout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_1_NO_BORDER: {
        'basename': 'Line Callout 1 (No Border)',
        'prst':     'callout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_2: {
        'basename': 'Line Callout 2',
        'prst':     'borderCallout2',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 112500),
            ('adj6', -46667),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_2_ACCENT_BAR: {
        'basename': 'Line Callout 2 (Accent Bar)',
        'prst':     'accentCallout2',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 112500),
            ('adj6', -46667),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 2 (Border and Accent Bar)',
        'prst':     'accentBorderCallout2',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 112500),
            ('adj6', -46667),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_2_NO_BORDER: {
        'basename': 'Line Callout 2 (No Border)',
        'prst':     'callout2',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 112500),
            ('adj6', -46667),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_3: {
        'basename': 'Line Callout 3',
        'prst':     'borderCallout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_3_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Accent Bar)',
        'prst':     'accentCallout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Border and Accent Bar)',
        'prst':     'accentBorderCallout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_3_NO_BORDER: {
        'basename': 'Line Callout 3 (No Border)',
        'prst':     'callout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_4: {
        'basename': 'Line Callout 3',
        'prst':     'borderCallout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_4_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Accent Bar)',
        'prst':     'accentCallout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Border and Accent Bar)',
        'prst':     'accentBorderCallout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_CALLOUT_4_NO_BORDER: {
        'basename': 'Line Callout 3 (No Border)',
        'prst':     'callout3',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 18750),
            ('adj4', -16667),
            ('adj5', 100000),
            ('adj6', -16667),
            ('adj7', 112963),
            ('adj8', -8333),
        )
    },
    MSO_SHAPE.LINE_INVERSE: {
        'basename': 'Straight Connector',
        'prst':     'lineInv',
        'avLst':    ()
    },
    MSO_SHAPE.MATH_DIVIDE: {
        'basename': 'Division',
        'prst':     'mathDivide',
        'avLst':    (
            ('adj1', 23520),
            ('adj2', 5880),
            ('adj3', 11760),
        )
    },
    MSO_SHAPE.MATH_EQUAL: {
        'basename': 'Equal',
        'prst':     'mathEqual',
        'avLst':    (
            ('adj1', 23520),
            ('adj2', 11760),
        )
    },
    MSO_SHAPE.MATH_MINUS: {
        'basename': 'Minus',
        'prst':     'mathMinus',
        'avLst':    (
            ('adj1', 23520),
        )
    },
    MSO_SHAPE.MATH_MULTIPLY: {
        'basename': 'Multiply',
        'prst':     'mathMultiply',
        'avLst':    (
            ('adj1', 23520),
        )
    },
    MSO_SHAPE.MATH_NOT_EQUAL: {
        'basename': 'Not Equal',
        'prst':     'mathNotEqual',
        'avLst':    (
            ('adj1', 23520),
            ('adj2', 6600000),
            ('adj3', 11760),
        )
    },
    MSO_SHAPE.MATH_PLUS: {
        'basename': 'Plus',
        'prst':     'mathPlus',
        'avLst':    (
            ('adj1', 23520),
        )
    },
    MSO_SHAPE.MOON: {
        'basename': 'Moon',
        'prst':     'moon',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MSO_SHAPE.NON_ISOSCELES_TRAPEZOID: {
        'basename': 'Non-isosceles Trapezoid',
        'prst':     'nonIsoscelesTrapezoid',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
        )
    },
    MSO_SHAPE.NOTCHED_RIGHT_ARROW: {
        'basename': 'Notched Right Arrow',
        'prst':     'notchedRightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.NO_SYMBOL: {
        'basename': '"No" symbol',
        'prst':     'noSmoking',
        'avLst':    (
            ('adj', 18750),
        )
    },
    MSO_SHAPE.OCTAGON: {
        'basename': 'Octagon',
        'prst':     'octagon',
        'avLst':    (
            ('adj', 29289),
        )
    },
    MSO_SHAPE.OVAL: {
        'basename': 'Oval',
        'prst':     'ellipse',
        'avLst':    ()
    },
    MSO_SHAPE.OVAL_CALLOUT: {
        'basename': 'Oval Callout',
        'prst':     'wedgeEllipseCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
        )
    },
    MSO_SHAPE.PARALLELOGRAM: {
        'basename': 'Parallelogram',
        'prst':     'parallelogram',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.PENTAGON: {
        'basename': 'Pentagon',
        'prst':     'homePlate',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MSO_SHAPE.PIE: {
        'basename': 'Pie',
        'prst':     'pie',
        'avLst':    (
            ('adj1', 0),
            ('adj2', 16200000),
        )
    },
    MSO_SHAPE.PIE_WEDGE: {
        'basename': 'Pie',
        'prst':     'pieWedge',
        'avLst':    ()
    },
    MSO_SHAPE.PLAQUE: {
        'basename': 'Plaque',
        'prst':     'plaque',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MSO_SHAPE.PLAQUE_TABS: {
        'basename': 'Plaque Tabs',
        'prst':     'plaqueTabs',
        'avLst':    ()
    },
    MSO_SHAPE.QUAD_ARROW: {
        'basename': 'Quad Arrow',
        'prst':     'quadArrow',
        'avLst':    (
            ('adj1', 22500),
            ('adj2', 22500),
            ('adj3', 22500),
        )
    },
    MSO_SHAPE.QUAD_ARROW_CALLOUT: {
        'basename': 'Quad Arrow Callout',
        'prst':     'quadArrowCallout',
        'avLst':    (
            ('adj1', 18515),
            ('adj2', 18515),
            ('adj3', 18515),
            ('adj4', 48123),
        )
    },
    MSO_SHAPE.RECTANGLE: {
        'basename': 'Rectangle',
        'prst':     'rect',
        'avLst':    ()
    },
    MSO_SHAPE.RECTANGULAR_CALLOUT: {
        'basename': 'Rectangular Callout',
        'prst':     'wedgeRectCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
        )
    },
    MSO_SHAPE.REGULAR_PENTAGON: {
        'basename': 'Regular Pentagon',
        'prst':     'pentagon',
        'avLst':    (
            ('hf', 105146),
            ('vf', 110557),
        )
    },
    MSO_SHAPE.RIGHT_ARROW: {
        'basename': 'Right Arrow',
        'prst':     'rightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.RIGHT_ARROW_CALLOUT: {
        'basename': 'Right Arrow Callout',
        'prst':     'rightArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MSO_SHAPE.RIGHT_BRACE: {
        'basename': 'Right Brace',
        'prst':     'rightBrace',
        'avLst':    (
            ('adj1', 8333),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.RIGHT_BRACKET: {
        'basename': 'Right Bracket',
        'prst':     'rightBracket',
        'avLst':    (
            ('adj', 8333),
        )
    },
    MSO_SHAPE.RIGHT_TRIANGLE: {
        'basename': 'Right Triangle',
        'prst':     'rtTriangle',
        'avLst':    ()
    },
    MSO_SHAPE.ROUNDED_RECTANGLE: {
        'basename': 'Rounded Rectangle',
        'prst':     'roundRect',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT: {
        'basename': 'Rounded Rectangular Callout',
        'prst':     'wedgeRoundRectCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
            ('adj3', 16667),
        )
    },
    MSO_SHAPE.ROUND_1_RECTANGLE: {
        'basename': 'Round Single Corner Rectangle',
        'prst':     'round1Rect',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MSO_SHAPE.ROUND_2_DIAG_RECTANGLE: {
        'basename': 'Round Diagonal Corner Rectangle',
        'prst':     'round2DiagRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 0),
        )
    },
    MSO_SHAPE.ROUND_2_SAME_RECTANGLE: {
        'basename': 'Round Same Side Corner Rectangle',
        'prst':     'round2SameRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 0),
        )
    },
    MSO_SHAPE.SMILEY_FACE: {
        'basename': 'Smiley Face',
        'prst':     'smileyFace',
        'avLst':    (
            ('adj', 4653),
        )
    },
    MSO_SHAPE.SNIP_1_RECTANGLE: {
        'basename': 'Snip Single Corner Rectangle',
        'prst':     'snip1Rect',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MSO_SHAPE.SNIP_2_DIAG_RECTANGLE: {
        'basename': 'Snip Diagonal Corner Rectangle',
        'prst':     'snip2DiagRect',
        'avLst':    (
            ('adj1', 0),
            ('adj2', 16667),
        )
    },
    MSO_SHAPE.SNIP_2_SAME_RECTANGLE: {
        'basename': 'Snip Same Side Corner Rectangle',
        'prst':     'snip2SameRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 0),
        )
    },
    MSO_SHAPE.SNIP_ROUND_RECTANGLE: {
        'basename': 'Snip and Round Single Corner Rectangle',
        'prst':     'snipRoundRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 16667),
        )
    },
    MSO_SHAPE.SQUARE_TABS: {
        'basename': 'Square Tabs',
        'prst':     'squareTabs',
        'avLst':    ()
    },
    MSO_SHAPE.STAR_10_POINT: {
        'basename': '10-Point Star',
        'prst':     'star10',
        'avLst':    (
            ('adj', 42533),
            ('hf', 105146),
        )
    },
    MSO_SHAPE.STAR_12_POINT: {
        'basename': '12-Point Star',
        'prst':     'star12',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MSO_SHAPE.STAR_16_POINT: {
        'basename': '16-Point Star',
        'prst':     'star16',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MSO_SHAPE.STAR_24_POINT: {
        'basename': '24-Point Star',
        'prst':     'star24',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MSO_SHAPE.STAR_32_POINT: {
        'basename': '32-Point Star',
        'prst':     'star32',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MSO_SHAPE.STAR_4_POINT: {
        'basename': '4-Point Star',
        'prst':     'star4',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MSO_SHAPE.STAR_5_POINT: {
        'basename': '5-Point Star',
        'prst':     'star5',
        'avLst':    (
            ('adj', 19098),
            ('hf', 105146),
            ('vf', 110557),
        )
    },
    MSO_SHAPE.STAR_6_POINT: {
        'basename': '6-Point Star',
        'prst':     'star6',
        'avLst':    (
            ('adj', 28868),
            ('hf', 115470),
        )
    },
    MSO_SHAPE.STAR_7_POINT: {
        'basename': '7-Point Star',
        'prst':     'star7',
        'avLst':    (
            ('adj', 34601),
            ('hf', 102572),
            ('vf', 105210),
        )
    },
    MSO_SHAPE.STAR_8_POINT: {
        'basename': '8-Point Star',
        'prst':     'star8',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MSO_SHAPE.STRIPED_RIGHT_ARROW: {
        'basename': 'Striped Right Arrow',
        'prst':     'stripedRightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.SUN: {
        'basename': 'Sun',
        'prst':     'sun',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.SWOOSH_ARROW: {
        'basename': 'Swoosh Arrow',
        'prst':     'swooshArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 16667),
        )
    },
    MSO_SHAPE.TEAR: {
        'basename': 'Teardrop',
        'prst':     'teardrop',
        'avLst':    (
            ('adj', 100000),
        )
    },
    MSO_SHAPE.TRAPEZOID: {
        'basename': 'Trapezoid',
        'prst':     'trapezoid',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MSO_SHAPE.UP_ARROW: {
        'basename': 'Up Arrow',
        'prst':     'upArrow',
        'avLst':    ()
    },
    MSO_SHAPE.UP_ARROW_CALLOUT: {
        'basename': 'Up Arrow Callout',
        'prst':     'upArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MSO_SHAPE.UP_DOWN_ARROW: {
        'basename': 'Up-Down Arrow',
        'prst':     'upDownArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj1', 50000),
            ('adj2', 50000),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.UP_DOWN_ARROW_CALLOUT: {
        'basename': 'Up-Down Arrow Callout',
        'prst':     'upDownArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 48123),
        )
    },
    MSO_SHAPE.UP_RIBBON: {
        'basename': 'Up Ribbon',
        'prst':     'ribbon2',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 50000),
        )
    },
    MSO_SHAPE.U_TURN_ARROW: {
        'basename': 'U-Turn Arrow',
        'prst':     'uturnArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 43750),
            ('adj5', 75000),
        )
    },
    MSO_SHAPE.VERTICAL_SCROLL: {
        'basename': 'Vertical Scroll',
        'prst':     'verticalScroll',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MSO_SHAPE.WAVE: {
        'basename': 'Wave',
        'prst':     'wave',
        'avLst':    (
            ('adj1', 12500),
            ('adj2', 0),
        )
    },
}
