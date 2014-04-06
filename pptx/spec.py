# encoding: utf-8

"""
Constant values from the ISO/IEC 29500 spec that are needed for XML
generation and packaging, and a utility function or two for accessing some of
them.
"""

from __future__ import absolute_import

from pptx.constants import MSO_AUTO_SHAPE_TYPE as MAST
from pptx.opc.constants import CONTENT_TYPE as CT


# ============================================================================
# AutoShape type specs
# ============================================================================

autoshape_types = {
    MAST.ACTION_BUTTON_BACK_OR_PREVIOUS: {
        'basename': 'Action Button: Back or Previous',
        'prst':     'actionButtonBackPrevious',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_BEGINNING: {
        'basename': 'Action Button: Beginning',
        'prst':     'actionButtonBeginning',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_CUSTOM: {
        'basename': 'Action Button: Custom',
        'prst':     'actionButtonBlank',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_DOCUMENT: {
        'basename': 'Action Button: Document',
        'prst':     'actionButtonDocument',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_END: {
        'basename': 'Action Button: End',
        'prst':     'actionButtonEnd',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_FORWARD_OR_NEXT: {
        'basename': 'Action Button: Forward or Next',
        'prst':     'actionButtonForwardNext',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_HELP: {
        'basename': 'Action Button: Help',
        'prst':     'actionButtonHelp',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_HOME: {
        'basename': 'Action Button: Home',
        'prst':     'actionButtonHome',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_INFORMATION: {
        'basename': 'Action Button: Information',
        'prst':     'actionButtonInformation',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_MOVIE: {
        'basename': 'Action Button: Movie',
        'prst':     'actionButtonMovie',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_RETURN: {
        'basename': 'Action Button: Return',
        'prst':     'actionButtonReturn',
        'avLst':    ()
    },
    MAST.ACTION_BUTTON_SOUND: {
        'basename': 'Action Button: Sound',
        'prst':     'actionButtonSound',
        'avLst':    ()
    },
    MAST.ARC: {
        'basename': 'Arc',
        'prst':     'arc',
        'avLst':    (
            ('adj1', 16200000),
            ('adj2', 0),
        )
    },
    MAST.BALLOON: {
        'basename': 'Rounded Rectangular Callout',
        'prst':     'wedgeRoundRectCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
            ('adj3', 16667),
        )
    },
    MAST.BENT_ARROW: {
        'basename': 'Bent Arrow',
        'prst':     'bentArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 43750),
        )
    },
    MAST.BENT_UP_ARROW: {
        'basename': 'Bent-Up Arrow',
        'prst':     'bentUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
        )
    },
    MAST.BEVEL: {
        'basename': 'Bevel',
        'prst':     'bevel',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MAST.BLOCK_ARC: {
        'basename': 'Block Arc',
        'prst':     'blockArc',
        'avLst':    (
            ('adj1', 10800000),
            ('adj2', 0),
            ('adj3', 25000),
        )
    },
    MAST.CAN: {
        'basename': 'Can',
        'prst':     'can',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.CHART_PLUS: {
        'basename': 'Chart Plus',
        'prst':     'chartPlus',
        'avLst':    ()
    },
    MAST.CHART_STAR: {
        'basename': 'Chart Star',
        'prst':     'chartStar',
        'avLst':    ()
    },
    MAST.CHART_X: {
        'basename': 'Chart X',
        'prst':     'chartX',
        'avLst':    ()
    },
    MAST.CHEVRON: {
        'basename': 'Chevron',
        'prst':     'chevron',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MAST.CHORD: {
        'basename': 'Chord',
        'prst':     'chord',
        'avLst':    (
            ('adj1', 2700000),
            ('adj2', 16200000),
        )
    },
    MAST.CIRCULAR_ARROW: {
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
    MAST.CLOUD: {
        'basename': 'Cloud',
        'prst':     'cloud',
        'avLst':    ()
    },
    MAST.CLOUD_CALLOUT: {
        'basename': 'Cloud Callout',
        'prst':     'cloudCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
        )
    },
    MAST.CORNER: {
        'basename': 'Corner',
        'prst':     'corner',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.CORNER_TABS: {
        'basename': 'Corner Tabs',
        'prst':     'cornerTabs',
        'avLst':    ()
    },
    MAST.CROSS: {
        'basename': 'Cross',
        'prst':     'plus',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.CUBE: {
        'basename': 'Cube',
        'prst':     'cube',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.CURVED_DOWN_ARROW: {
        'basename': 'Curved Down Arrow',
        'prst':     'curvedDownArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MAST.CURVED_DOWN_RIBBON: {
        'basename': 'Curved Down Ribbon',
        'prst':     'ellipseRibbon',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 12500),
        )
    },
    MAST.CURVED_LEFT_ARROW: {
        'basename': 'Curved Left Arrow',
        'prst':     'curvedLeftArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MAST.CURVED_RIGHT_ARROW: {
        'basename': 'Curved Right Arrow',
        'prst':     'curvedRightArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MAST.CURVED_UP_ARROW: {
        'basename': 'Curved Up Arrow',
        'prst':     'curvedUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 25000),
        )
    },
    MAST.CURVED_UP_RIBBON: {
        'basename': 'Curved Up Ribbon',
        'prst':     'ellipseRibbon2',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 50000),
            ('adj3', 12500),
        )
    },
    MAST.DECAGON: {
        'basename': 'Decagon',
        'prst':     'decagon',
        'avLst':    (
            ('vf', 105146),
        )
    },
    MAST.DIAGONAL_STRIPE: {
        'basename': 'Diagonal Stripe',
        'prst':     'diagStripe',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MAST.DIAMOND: {
        'basename': 'Diamond',
        'prst':     'diamond',
        'avLst':    ()
    },
    MAST.DODECAGON: {
        'basename': 'Dodecagon',
        'prst':     'dodecagon',
        'avLst':    ()
    },
    MAST.DONUT: {
        'basename': 'Donut',
        'prst':     'donut',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.DOUBLE_BRACE: {
        'basename': 'Double Brace',
        'prst':     'bracePair',
        'avLst':    (
            ('adj', 8333),
        )
    },
    MAST.DOUBLE_BRACKET: {
        'basename': 'Double Bracket',
        'prst':     'bracketPair',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MAST.DOUBLE_WAVE: {
        'basename': 'Double Wave',
        'prst':     'doubleWave',
        'avLst':    (
            ('adj1', 6250),
            ('adj2', 0),
        )
    },
    MAST.DOWN_ARROW: {
        'basename': 'Down Arrow',
        'prst':     'downArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.DOWN_ARROW_CALLOUT: {
        'basename': 'Down Arrow Callout',
        'prst':     'downArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MAST.DOWN_RIBBON: {
        'basename': 'Down Ribbon',
        'prst':     'ribbon',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 50000),
        )
    },
    MAST.EXPLOSION1: {
        'basename': 'Explosion',
        'prst':     'irregularSeal1',
        'avLst':    ()
    },
    MAST.EXPLOSION2: {
        'basename': 'Explosion',
        'prst':     'irregularSeal2',
        'avLst':    ()
    },
    MAST.FLOWCHART_ALTERNATE_PROCESS: {
        'basename': 'Alternate process',
        'prst':     'flowChartAlternateProcess',
        'avLst':    ()
    },
    MAST.FLOWCHART_CARD: {
        'basename': 'Card',
        'prst':     'flowChartPunchedCard',
        'avLst':    ()
    },
    MAST.FLOWCHART_COLLATE: {
        'basename': 'Collate',
        'prst':     'flowChartCollate',
        'avLst':    ()
    },
    MAST.FLOWCHART_CONNECTOR: {
        'basename': 'Connector',
        'prst':     'flowChartConnector',
        'avLst':    ()
    },
    MAST.FLOWCHART_DATA: {
        'basename': 'Data',
        'prst':     'flowChartInputOutput',
        'avLst':    ()
    },
    MAST.FLOWCHART_DECISION: {
        'basename': 'Decision',
        'prst':     'flowChartDecision',
        'avLst':    ()
    },
    MAST.FLOWCHART_DELAY: {
        'basename': 'Delay',
        'prst':     'flowChartDelay',
        'avLst':    ()
    },
    MAST.FLOWCHART_DIRECT_ACCESS_STORAGE: {
        'basename': 'Direct Access Storage',
        'prst':     'flowChartMagneticDrum',
        'avLst':    ()
    },
    MAST.FLOWCHART_DISPLAY: {
        'basename': 'Display',
        'prst':     'flowChartDisplay',
        'avLst':    ()
    },
    MAST.FLOWCHART_DOCUMENT: {
        'basename': 'Document',
        'prst':     'flowChartDocument',
        'avLst':    ()
    },
    MAST.FLOWCHART_EXTRACT: {
        'basename': 'Extract',
        'prst':     'flowChartExtract',
        'avLst':    ()
    },
    MAST.FLOWCHART_INTERNAL_STORAGE: {
        'basename': 'Internal Storage',
        'prst':     'flowChartInternalStorage',
        'avLst':    ()
    },
    MAST.FLOWCHART_MAGNETIC_DISK: {
        'basename': 'Magnetic Disk',
        'prst':     'flowChartMagneticDisk',
        'avLst':    ()
    },
    MAST.FLOWCHART_MANUAL_INPUT: {
        'basename': 'Manual Input',
        'prst':     'flowChartManualInput',
        'avLst':    ()
    },
    MAST.FLOWCHART_MANUAL_OPERATION: {
        'basename': 'Manual Operation',
        'prst':     'flowChartManualOperation',
        'avLst':    ()
    },
    MAST.FLOWCHART_MERGE: {
        'basename': 'Merge',
        'prst':     'flowChartMerge',
        'avLst':    ()
    },
    MAST.FLOWCHART_MULTIDOCUMENT: {
        'basename': 'Multidocument',
        'prst':     'flowChartMultidocument',
        'avLst':    ()
    },
    MAST.FLOWCHART_OFFLINE_STORAGE: {
        'basename': 'Offline Storage',
        'prst':     'flowChartOfflineStorage',
        'avLst':    ()
    },
    MAST.FLOWCHART_OFFPAGE_CONNECTOR: {
        'basename': 'Off-page Connector',
        'prst':     'flowChartOffpageConnector',
        'avLst':    ()
    },
    MAST.FLOWCHART_OR: {
        'basename': 'Or',
        'prst':     'flowChartOr',
        'avLst':    ()
    },
    MAST.FLOWCHART_PREDEFINED_PROCESS: {
        'basename': 'Predefined Process',
        'prst':     'flowChartPredefinedProcess',
        'avLst':    ()
    },
    MAST.FLOWCHART_PREPARATION: {
        'basename': 'Preparation',
        'prst':     'flowChartPreparation',
        'avLst':    ()
    },
    MAST.FLOWCHART_PROCESS: {
        'basename': 'Process',
        'prst':     'flowChartProcess',
        'avLst':    ()
    },
    MAST.FLOWCHART_PUNCHED_TAPE: {
        'basename': 'Punched Tape',
        'prst':     'flowChartPunchedTape',
        'avLst':    ()
    },
    MAST.FLOWCHART_SEQUENTIAL_ACCESS_STORAGE: {
        'basename': 'Sequential Access Storage',
        'prst':     'flowChartMagneticTape',
        'avLst':    ()
    },
    MAST.FLOWCHART_SORT: {
        'basename': 'Sort',
        'prst':     'flowChartSort',
        'avLst':    ()
    },
    MAST.FLOWCHART_STORED_DATA: {
        'basename': 'Stored Data',
        'prst':     'flowChartOnlineStorage',
        'avLst':    ()
    },
    MAST.FLOWCHART_SUMMING_JUNCTION: {
        'basename': 'Summing Junction',
        'prst':     'flowChartSummingJunction',
        'avLst':    ()
    },
    MAST.FLOWCHART_TERMINATOR: {
        'basename': 'Terminator',
        'prst':     'flowChartTerminator',
        'avLst':    ()
    },
    MAST.FOLDED_CORNER: {
        'basename': 'Folded Corner',
        'prst':     'foldedCorner',
        'avLst':    ()
    },
    MAST.FRAME: {
        'basename': 'Frame',
        'prst':     'frame',
        'avLst':    (
            ('adj1', 12500),
        )
    },
    MAST.FUNNEL: {
        'basename': 'Funnel',
        'prst':     'funnel',
        'avLst':    ()
    },
    MAST.GEAR_6: {
        'basename': 'Gear 6',
        'prst':     'gear6',
        'avLst':    (
            ('adj1', 15000),
            ('adj2', 3526),
        )
    },
    MAST.GEAR_9: {
        'basename': 'Gear 9',
        'prst':     'gear9',
        'avLst':    (
            ('adj1', 10000),
            ('adj2', 1763),
        )
    },
    MAST.HALF_FRAME: {
        'basename': 'Half Frame',
        'prst':     'halfFrame',
        'avLst':    (
            ('adj1', 33333),
            ('adj2', 33333),
        )
    },
    MAST.HEART: {
        'basename': 'Heart',
        'prst':     'heart',
        'avLst':    ()
    },
    MAST.HEPTAGON: {
        'basename': 'Heptagon',
        'prst':     'heptagon',
        'avLst':    (
            ('hf', 102572),
            ('vf', 105210),
        )
    },
    MAST.HEXAGON: {
        'basename': 'Hexagon',
        'prst':     'hexagon',
        'avLst':    (
            ('adj', 25000),
            ('vf', 115470),
        )
    },
    MAST.HORIZONTAL_SCROLL: {
        'basename': 'Horizontal Scroll',
        'prst':     'horizontalScroll',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MAST.ISOSCELES_TRIANGLE: {
        'basename': 'Isosceles Triangle',
        'prst':     'triangle',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MAST.LEFT_ARROW: {
        'basename': 'Left Arrow',
        'prst':     'leftArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.LEFT_ARROW_CALLOUT: {
        'basename': 'Left Arrow Callout',
        'prst':     'leftArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MAST.LEFT_BRACE: {
        'basename': 'Left Brace',
        'prst':     'leftBrace',
        'avLst':    (
            ('adj1', 8333),
            ('adj2', 50000),
        )
    },
    MAST.LEFT_BRACKET: {
        'basename': 'Left Bracket',
        'prst':     'leftBracket',
        'avLst':    (
            ('adj', 8333),
        )
    },
    MAST.LEFT_CIRCULAR_ARROW: {
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
    MAST.LEFT_RIGHT_ARROW: {
        'basename': 'Left-Right Arrow',
        'prst':     'leftRightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.LEFT_RIGHT_ARROW_CALLOUT: {
        'basename': 'Left-Right Arrow Callout',
        'prst':     'leftRightArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 48123),
        )
    },
    MAST.LEFT_RIGHT_CIRCULAR_ARROW: {
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
    MAST.LEFT_RIGHT_RIBBON: {
        'basename': 'Left Right Ribbon',
        'prst':     'leftRightRibbon',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
            ('adj3', 16667),
        )
    },
    MAST.LEFT_RIGHT_UP_ARROW: {
        'basename': 'Left-Right-Up Arrow',
        'prst':     'leftRightUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
        )
    },
    MAST.LEFT_UP_ARROW: {
        'basename': 'Left-Up Arrow',
        'prst':     'leftUpArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
        )
    },
    MAST.LIGHTNING_BOLT: {
        'basename': 'Lightning Bolt',
        'prst':     'lightningBolt',
        'avLst':    ()
    },
    MAST.LINE_CALLOUT_1: {
        'basename': 'Line Callout 1',
        'prst':     'borderCallout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MAST.LINE_CALLOUT_1_ACCENT_BAR: {
        'basename': 'Line Callout 1 (Accent Bar)',
        'prst':     'accentCallout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MAST.LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 1 (Border and Accent Bar)',
        'prst':     'accentBorderCallout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MAST.LINE_CALLOUT_1_NO_BORDER: {
        'basename': 'Line Callout 1 (No Border)',
        'prst':     'callout1',
        'avLst':    (
            ('adj1', 18750),
            ('adj2', -8333),
            ('adj3', 112500),
            ('adj4', -38333),
        )
    },
    MAST.LINE_CALLOUT_2: {
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
    MAST.LINE_CALLOUT_2_ACCENT_BAR: {
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
    MAST.LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR: {
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
    MAST.LINE_CALLOUT_2_NO_BORDER: {
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
    MAST.LINE_CALLOUT_3: {
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
    MAST.LINE_CALLOUT_3_ACCENT_BAR: {
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
    MAST.LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR: {
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
    MAST.LINE_CALLOUT_3_NO_BORDER: {
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
    MAST.LINE_CALLOUT_4: {
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
    MAST.LINE_CALLOUT_4_ACCENT_BAR: {
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
    MAST.LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR: {
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
    MAST.LINE_CALLOUT_4_NO_BORDER: {
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
    MAST.LINE_INVERSE: {
        'basename': 'Straight Connector',
        'prst':     'lineInv',
        'avLst':    ()
    },
    MAST.MATH_DIVIDE: {
        'basename': 'Division',
        'prst':     'mathDivide',
        'avLst':    (
            ('adj1', 23520),
            ('adj2', 5880),
            ('adj3', 11760),
        )
    },
    MAST.MATH_EQUAL: {
        'basename': 'Equal',
        'prst':     'mathEqual',
        'avLst':    (
            ('adj1', 23520),
            ('adj2', 11760),
        )
    },
    MAST.MATH_MINUS: {
        'basename': 'Minus',
        'prst':     'mathMinus',
        'avLst':    (
            ('adj1', 23520),
        )
    },
    MAST.MATH_MULTIPLY: {
        'basename': 'Multiply',
        'prst':     'mathMultiply',
        'avLst':    (
            ('adj1', 23520),
        )
    },
    MAST.MATH_NOT_EQUAL: {
        'basename': 'Not Equal',
        'prst':     'mathNotEqual',
        'avLst':    (
            ('adj1', 23520),
            ('adj2', 6600000),
            ('adj3', 11760),
        )
    },
    MAST.MATH_PLUS: {
        'basename': 'Plus',
        'prst':     'mathPlus',
        'avLst':    (
            ('adj1', 23520),
        )
    },
    MAST.MOON: {
        'basename': 'Moon',
        'prst':     'moon',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MAST.NON_ISOSCELES_TRAPEZOID: {
        'basename': 'Non-isosceles Trapezoid',
        'prst':     'nonIsoscelesTrapezoid',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
        )
    },
    MAST.NOTCHED_RIGHT_ARROW: {
        'basename': 'Notched Right Arrow',
        'prst':     'notchedRightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.NO_SYMBOL: {
        'basename': '"No" symbol',
        'prst':     'noSmoking',
        'avLst':    (
            ('adj', 18750),
        )
    },
    MAST.OCTAGON: {
        'basename': 'Octagon',
        'prst':     'octagon',
        'avLst':    (
            ('adj', 29289),
        )
    },
    MAST.OVAL: {
        'basename': 'Oval',
        'prst':     'ellipse',
        'avLst':    ()
    },
    MAST.OVAL_CALLOUT: {
        'basename': 'Oval Callout',
        'prst':     'wedgeEllipseCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
        )
    },
    MAST.PARALLELOGRAM: {
        'basename': 'Parallelogram',
        'prst':     'parallelogram',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.PENTAGON: {
        'basename': 'Pentagon',
        'prst':     'homePlate',
        'avLst':    (
            ('adj', 50000),
        )
    },
    MAST.PIE: {
        'basename': 'Pie',
        'prst':     'pie',
        'avLst':    (
            ('adj1', 0),
            ('adj2', 16200000),
        )
    },
    MAST.PIE_WEDGE: {
        'basename': 'Pie',
        'prst':     'pieWedge',
        'avLst':    ()
    },
    MAST.PLAQUE: {
        'basename': 'Plaque',
        'prst':     'plaque',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MAST.PLAQUE_TABS: {
        'basename': 'Plaque Tabs',
        'prst':     'plaqueTabs',
        'avLst':    ()
    },
    MAST.QUAD_ARROW: {
        'basename': 'Quad Arrow',
        'prst':     'quadArrow',
        'avLst':    (
            ('adj1', 22500),
            ('adj2', 22500),
            ('adj3', 22500),
        )
    },
    MAST.QUAD_ARROW_CALLOUT: {
        'basename': 'Quad Arrow Callout',
        'prst':     'quadArrowCallout',
        'avLst':    (
            ('adj1', 18515),
            ('adj2', 18515),
            ('adj3', 18515),
            ('adj4', 48123),
        )
    },
    MAST.RECTANGLE: {
        'basename': 'Rectangle',
        'prst':     'rect',
        'avLst':    ()
    },
    MAST.RECTANGULAR_CALLOUT: {
        'basename': 'Rectangular Callout',
        'prst':     'wedgeRectCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
        )
    },
    MAST.REGULAR_PENTAGON: {
        'basename': 'Regular Pentagon',
        'prst':     'pentagon',
        'avLst':    (
            ('hf', 105146),
            ('vf', 110557),
        )
    },
    MAST.RIGHT_ARROW: {
        'basename': 'Right Arrow',
        'prst':     'rightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.RIGHT_ARROW_CALLOUT: {
        'basename': 'Right Arrow Callout',
        'prst':     'rightArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MAST.RIGHT_BRACE: {
        'basename': 'Right Brace',
        'prst':     'rightBrace',
        'avLst':    (
            ('adj1', 8333),
            ('adj2', 50000),
        )
    },
    MAST.RIGHT_BRACKET: {
        'basename': 'Right Bracket',
        'prst':     'rightBracket',
        'avLst':    (
            ('adj', 8333),
        )
    },
    MAST.RIGHT_TRIANGLE: {
        'basename': 'Right Triangle',
        'prst':     'rtTriangle',
        'avLst':    ()
    },
    MAST.ROUNDED_RECTANGLE: {
        'basename': 'Rounded Rectangle',
        'prst':     'roundRect',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MAST.ROUNDED_RECTANGULAR_CALLOUT: {
        'basename': 'Rounded Rectangular Callout',
        'prst':     'wedgeRoundRectCallout',
        'avLst':    (
            ('adj1', -20833),
            ('adj2', 62500),
            ('adj3', 16667),
        )
    },
    MAST.ROUND_1_RECTANGLE: {
        'basename': 'Round Single Corner Rectangle',
        'prst':     'round1Rect',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MAST.ROUND_2_DIAG_RECTANGLE: {
        'basename': 'Round Diagonal Corner Rectangle',
        'prst':     'round2DiagRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 0),
        )
    },
    MAST.ROUND_2_SAME_RECTANGLE: {
        'basename': 'Round Same Side Corner Rectangle',
        'prst':     'round2SameRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 0),
        )
    },
    MAST.SMILEY_FACE: {
        'basename': 'Smiley Face',
        'prst':     'smileyFace',
        'avLst':    (
            ('adj', 4653),
        )
    },
    MAST.SNIP_1_RECTANGLE: {
        'basename': 'Snip Single Corner Rectangle',
        'prst':     'snip1Rect',
        'avLst':    (
            ('adj', 16667),
        )
    },
    MAST.SNIP_2_DIAG_RECTANGLE: {
        'basename': 'Snip Diagonal Corner Rectangle',
        'prst':     'snip2DiagRect',
        'avLst':    (
            ('adj1', 0),
            ('adj2', 16667),
        )
    },
    MAST.SNIP_2_SAME_RECTANGLE: {
        'basename': 'Snip Same Side Corner Rectangle',
        'prst':     'snip2SameRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 0),
        )
    },
    MAST.SNIP_ROUND_RECTANGLE: {
        'basename': 'Snip and Round Single Corner Rectangle',
        'prst':     'snipRoundRect',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 16667),
        )
    },
    MAST.SQUARE_TABS: {
        'basename': 'Square Tabs',
        'prst':     'squareTabs',
        'avLst':    ()
    },
    MAST.STAR_10_POINT: {
        'basename': '10-Point Star',
        'prst':     'star10',
        'avLst':    (
            ('adj', 42533),
            ('hf', 105146),
        )
    },
    MAST.STAR_12_POINT: {
        'basename': '12-Point Star',
        'prst':     'star12',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MAST.STAR_16_POINT: {
        'basename': '16-Point Star',
        'prst':     'star16',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MAST.STAR_24_POINT: {
        'basename': '24-Point Star',
        'prst':     'star24',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MAST.STAR_32_POINT: {
        'basename': '32-Point Star',
        'prst':     'star32',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MAST.STAR_4_POINT: {
        'basename': '4-Point Star',
        'prst':     'star4',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MAST.STAR_5_POINT: {
        'basename': '5-Point Star',
        'prst':     'star5',
        'avLst':    (
            ('adj', 19098),
            ('hf', 105146),
            ('vf', 110557),
        )
    },
    MAST.STAR_6_POINT: {
        'basename': '6-Point Star',
        'prst':     'star6',
        'avLst':    (
            ('adj', 28868),
            ('hf', 115470),
        )
    },
    MAST.STAR_7_POINT: {
        'basename': '7-Point Star',
        'prst':     'star7',
        'avLst':    (
            ('adj', 34601),
            ('hf', 102572),
            ('vf', 105210),
        )
    },
    MAST.STAR_8_POINT: {
        'basename': '8-Point Star',
        'prst':     'star8',
        'avLst':    (
            ('adj', 37500),
        )
    },
    MAST.STRIPED_RIGHT_ARROW: {
        'basename': 'Striped Right Arrow',
        'prst':     'stripedRightArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj2', 50000),
        )
    },
    MAST.SUN: {
        'basename': 'Sun',
        'prst':     'sun',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.SWOOSH_ARROW: {
        'basename': 'Swoosh Arrow',
        'prst':     'swooshArrow',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 16667),
        )
    },
    MAST.TEAR: {
        'basename': 'Teardrop',
        'prst':     'teardrop',
        'avLst':    (
            ('adj', 100000),
        )
    },
    MAST.TRAPEZOID: {
        'basename': 'Trapezoid',
        'prst':     'trapezoid',
        'avLst':    (
            ('adj', 25000),
        )
    },
    MAST.UP_ARROW: {
        'basename': 'Up Arrow',
        'prst':     'upArrow',
        'avLst':    ()
    },
    MAST.UP_ARROW_CALLOUT: {
        'basename': 'Up Arrow Callout',
        'prst':     'upArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 64977),
        )
    },
    MAST.UP_DOWN_ARROW: {
        'basename': 'Up-Down Arrow',
        'prst':     'upDownArrow',
        'avLst':    (
            ('adj1', 50000),
            ('adj1', 50000),
            ('adj2', 50000),
            ('adj2', 50000),
        )
    },
    MAST.UP_DOWN_ARROW_CALLOUT: {
        'basename': 'Up-Down Arrow Callout',
        'prst':     'upDownArrowCallout',
        'avLst':    (
            ('adj1', 25000),
            ('adj2', 25000),
            ('adj3', 25000),
            ('adj4', 48123),
        )
    },
    MAST.UP_RIBBON: {
        'basename': 'Up Ribbon',
        'prst':     'ribbon2',
        'avLst':    (
            ('adj1', 16667),
            ('adj2', 50000),
        )
    },
    MAST.U_TURN_ARROW: {
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
    MAST.VERTICAL_SCROLL: {
        'basename': 'Vertical Scroll',
        'prst':     'verticalScroll',
        'avLst':    (
            ('adj', 12500),
        )
    },
    MAST.WAVE: {
        'basename': 'Wave',
        'prst':     'wave',
        'avLst':    (
            ('adj1', 12500),
            ('adj2', 0),
        )
    },
}


# ============================================================================
# Placeholder constants
# ============================================================================

# valid values for <p:ph> type attribute (ST_PlaceholderType)
# -----------------------------------------------------------
PH_TYPE_TITLE = 'title'
PH_TYPE_BODY = 'body'
PH_TYPE_CTRTITLE = 'ctrTitle'
PH_TYPE_SUBTITLE = 'subTitle'
PH_TYPE_DT = 'dt'
PH_TYPE_SLDNUM = 'sldNum'
PH_TYPE_FTR = 'ftr'
PH_TYPE_HDR = 'hdr'
PH_TYPE_OBJ = 'obj'
PH_TYPE_CHART = 'chart'
PH_TYPE_TBL = 'tbl'
PH_TYPE_CLIPART = 'clipArt'
PH_TYPE_DGM = 'dgm'
PH_TYPE_MEDIA = 'media'
PH_TYPE_SLDIMG = 'sldImg'
PH_TYPE_PIC = 'pic'

# valid values for <p:ph> orient attribute
# ----------------------------------------
PH_ORIENT_HORZ = 'horz'
PH_ORIENT_VERT = 'vert'

# valid values for <p:ph> sz (size) attribute
# -------------------------------------------
PH_SZ_FULL = 'full'
PH_SZ_HALF = 'half'
PH_SZ_QUARTER = 'quarter'

# mapping of type to rootname (part of auto-generated placeholder shape name)
slide_ph_basenames = {
    PH_TYPE_TITLE:    'Title',
    # this next one is named 'Notes Placeholder' in a notes master
    PH_TYPE_BODY:     'Text Placeholder',
    PH_TYPE_CTRTITLE: 'Title',
    PH_TYPE_SUBTITLE: 'Subtitle',
    PH_TYPE_DT:       'Date Placeholder',
    PH_TYPE_SLDNUM:   'Slide Number Placeholder',
    PH_TYPE_FTR:      'Footer Placeholder',
    PH_TYPE_HDR:      'Header Placeholder',
    PH_TYPE_OBJ:      'Content Placeholder',
    PH_TYPE_CHART:    'Chart Placeholder',
    PH_TYPE_TBL:      'Table Placeholder',
    PH_TYPE_CLIPART:  'ClipArt Placeholder',
    PH_TYPE_DGM:      'SmartArt Placeholder',
    PH_TYPE_MEDIA:    'Media Placeholder',
    PH_TYPE_SLDIMG:   'Slide Image Placeholder',
    PH_TYPE_PIC:      'Picture Placeholder'
}


# ============================================================================
# default_content_types
# ============================================================================
# Default file extension to MIME type mapping. This is used as a reference for
# adding <Default> elements to [Content_Types].xml for parts like media files.
#
# BACKLOG: I've seen .wmv elements in the media folder of at least one
# presentation, might need to add an entry for that and perhaps other rich
# media PowerPoint allows to be embedded (e.g. audio, movie, object, ...).
# ============================================================================

default_content_types = {
    '.bin':     CT.PML_PRINTER_SETTINGS,
    '.bmp':     'image/bmp',
    '.emf':     'image/x-emf',
    '.fntdata': 'application/x-fontdata',
    '.gif':     'image/gif',
    '.jpe':     'image/jpeg',
    '.jpeg':    'image/jpeg',
    '.jpg':     'image/jpeg',
    '.png':     'image/png',
    '.rels':    'application/vnd.openxmlformats-package.relationships+xml',
    '.tif':     'image/tiff',
    '.tiff':    'image/tiff',
    '.wdp':     'image/vnd.ms-photo',
    '.wmf':     'image/x-wmf',
    '.xlsx':    CT.SML_SHEET,
    '.xml':     'application/xml'
}
