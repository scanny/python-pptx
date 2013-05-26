# -*- coding: utf-8 -*-
#
# spec.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""
Constant values from the ECMA-376 spec that are needed for XML generation and
packaging, and a utility function or two for accessing some of them.
"""

from constants import MSO, PP, TEXT_ALIGN_TYPE as TAT


class ParagraphAlignment(object):
    """
    Mappings between ``PpParagraphAlignment`` values used in the API and
    ``ST_TextAlignType`` values used in the XML. ``PpParagraphAlignment``
    values are like ``PP.ALIGN_CENTER``.
    """
    _mapping = {
        PP.ALIGN_CENTER:          TAT.CENTER,
        PP.ALIGN_DISTRIBUTE:      TAT.DISTRIBUTE,
        PP.ALIGN_JUSTIFY:         TAT.JUSTIFY,
        PP.ALIGN_JUSTIFY_LOW:     TAT.JUSTIFY_LOW,
        PP.ALIGN_LEFT:            TAT.LEFT,
        PP.ALIGN_RIGHT:           TAT.RIGHT,
        PP.ALIGN_THAI_DISTRIBUTE: TAT.THAI_DISTRIBUTE
    }

    @classmethod
    def to_text_align_type(cls, alignment):
        """
        Map a paragraph alignment value (e.g. ``PP.ALIGN_CENTER`` to an
        ``ST_TextAlignType`` value (e.g. ``TAT.CENTER`` or ``'ctr'``).
        """
        try:
            text_align_type = cls._mapping[alignment]
        except KeyError:
            tmpl = "no ST_TextAlignType value for alignment '%s'"
            raise KeyError(tmpl % alignment)
        return text_align_type


# ============================================================================
# AutoShape type specs
# ============================================================================

autoshape_types = {
    MSO.SHAPE_10_POINT_STAR: {
        'basename': '10-Point Star',
        'prst':     'star10',
        'desc':     '10-Point Star'
    },
    MSO.SHAPE_12_POINT_STAR: {
        'basename': '12-Point Star',
        'prst':     'star12',
        'desc':     '12-Point Star'
    },
    MSO.SHAPE_16_POINT_STAR: {
        'basename': '16-Point Star',
        'prst':     'star16',
        'desc':     '16-point star'
    },
    MSO.SHAPE_24_POINT_STAR: {
        'basename': '24-Point Star',
        'prst':     'star24',
        'desc':     '24-point star'
    },
    MSO.SHAPE_32_POINT_STAR: {
        'basename': '32-Point Star',
        'prst':     'star32',
        'desc':     '32-point star'
    },
    MSO.SHAPE_4_POINT_STAR: {
        'basename': '4-Point Star',
        'prst':     'star4',
        'desc':     '4-point star'
    },
    MSO.SHAPE_5_POINT_STAR: {
        'basename': '5-Point Star',
        'prst':     'star5',
        'desc':     '5-point star'
    },
    MSO.SHAPE_6_POINT_STAR: {
        'basename': '6-Point Star',
        'prst':     'star6',
        'desc':     '6-Point Star'
    },
    MSO.SHAPE_7_POINT_STAR: {
        'basename': '7-Point Star',
        'prst':     'star7',
        'desc':     '7-Point Star'
    },
    MSO.SHAPE_8_POINT_STAR: {
        'basename': '8-Point Star',
        'prst':     'star8',
        'desc':     '8-point star'
    },
    MSO.SHAPE_ACTION_BUTTON_BACK_OR_PREVIOUS: {
        'basename': 'Action Button: Back or Previous',
        'prst':     'actionButtonBackPrevious',
        'desc':     ('Back or Previous button. Supports mouse-click and mouse'
                     '-over actions')
    },
    MSO.SHAPE_ACTION_BUTTON_BEGINNING: {
        'basename': 'Action Button: Beginning',
        'prst':     'actionButtonBeginning',
        'desc':     ('Beginning button. Supports mouse-click and mouse-over a'
                     'ctions')
    },
    MSO.SHAPE_ACTION_BUTTON_CUSTOM: {
        'basename': 'Action Button: Custom',
        'prst':     'actionButtonBlank',
        'desc':     ('Button with no default picture or text. Supports mouse-'
                     'click and mouse-over actions')
    },
    MSO.SHAPE_ACTION_BUTTON_DOCUMENT: {
        'basename': 'Action Button: Document',
        'prst':     'actionButtonDocument',
        'desc':     ('Document button. Supports mouse-click and mouse-over ac'
                     'tions')
    },
    MSO.SHAPE_ACTION_BUTTON_END: {
        'basename': 'Action Button: End',
        'prst':     'actionButtonEnd',
        'desc':     'End button. Supports mouse-click and mouse-over actions'
    },
    MSO.SHAPE_ACTION_BUTTON_FORWARD_OR_NEXT: {
        'basename': 'Action Button: Forward or Next',
        'prst':     'actionButtonForwardNext',
        'desc':     ('Forward or Next button. Supports mouse-click and mouse-'
                     'over actions')
    },
    MSO.SHAPE_ACTION_BUTTON_HELP: {
        'basename': 'Action Button: Help',
        'prst':     'actionButtonHelp',
        'desc':     'Help button. Supports mouse-click and mouse-over actions'
    },
    MSO.SHAPE_ACTION_BUTTON_HOME: {
        'basename': 'Action Button: Home',
        'prst':     'actionButtonHome',
        'desc':     'Home button. Supports mouse-click and mouse-over actions'
    },
    MSO.SHAPE_ACTION_BUTTON_INFORMATION: {
        'basename': 'Action Button: Information',
        'prst':     'actionButtonInformation',
        'desc':     ('Information button. Supports mouse-click and mouse-over'
                     ' actions')
    },
    MSO.SHAPE_ACTION_BUTTON_MOVIE: {
        'basename': 'Action Button: Movie',
        'prst':     'actionButtonMovie',
        'desc':     'Movie button. Supports mouse-click and mouse-over actions'
    },
    MSO.SHAPE_ACTION_BUTTON_RETURN: {
        'basename': 'Action Button: Return',
        'prst':     'actionButtonReturn',
        'desc':     ('Return button. Supports mouse-click and mouse-over acti'
                     'ons')
    },
    MSO.SHAPE_ACTION_BUTTON_SOUND: {
        'basename': 'Action Button: Sound',
        'prst':     'actionButtonSound',
        'desc':     'Sound button. Supports mouse-click and mouse-over actions'
    },
    MSO.SHAPE_ARC: {
        'basename': 'Arc',
        'prst':     'arc',
        'desc':     'Arc'
    },
    MSO.SHAPE_BALLOON: {
        'basename': 'Rounded Rectangular Callout',
        'prst':     'wedgeRoundRectCallout',
        'desc':     'Rounded Rectangular Callout'
    },
    MSO.SHAPE_BENT_ARROW: {
        'basename': 'Bent Arrow',
        'prst':     'bentArrow',
        'desc':     'Block arrow that follows a curved 90-degree angle'
    },
    MSO.SHAPE_BENT_UP_ARROW: {
        'basename': 'Bent-Up Arrow',
        'prst':     'bentUpArrow',
        'desc':     ('Block arrow that follows a sharp 90-degree angle. Point'
                     's up by default')
    },
    MSO.SHAPE_BEVEL: {
        'basename': 'Bevel',
        'prst':     'bevel',
        'desc':     'Bevel'
    },
    MSO.SHAPE_BLOCK_ARC: {
        'basename': 'Block Arc',
        'prst':     'blockArc',
        'desc':     'Block arc'
    },
    MSO.SHAPE_CAN: {
        'basename': 'Can',
        'prst':     'can',
        'desc':     'Can'
    },
    MSO.SHAPE_CHART_PLUS: {
        'basename': 'Chart Plus',
        'prst':     'chartPlus',
        'desc':     'Chart Plus'
    },
    MSO.SHAPE_CHART_STAR: {
        'basename': 'Chart Star',
        'prst':     'chartStar',
        'desc':     'Chart Star'
    },
    MSO.SHAPE_CHART_X: {
        'basename': 'Chart X',
        'prst':     'chartX',
        'desc':     'Chart X'
    },
    MSO.SHAPE_CHEVRON: {
        'basename': 'Chevron',
        'prst':     'chevron',
        'desc':     'Chevron'
    },
    MSO.SHAPE_CHORD: {
        'basename': 'Chord',
        'prst':     'chord',
        'desc':     'Geometric chord shape'
    },
    MSO.SHAPE_CIRCULAR_ARROW: {
        'basename': 'Circular Arrow',
        'prst':     'circularArrow',
        'desc':     'Block arrow that follows a curved 180-degree angle'
    },
    MSO.SHAPE_CLOUD: {
        'basename': 'Cloud',
        'prst':     'cloud',
        'desc':     'Cloud'
    },
    MSO.SHAPE_CLOUD_CALLOUT: {
        'basename': 'Cloud Callout',
        'prst':     'cloudCallout',
        'desc':     'Cloud callout'
    },
    MSO.SHAPE_CORNER: {
        'basename': 'Corner',
        'prst':     'corner',
        'desc':     'Corner'
    },
    MSO.SHAPE_CORNER_TABS: {
        'basename': 'Corner Tabs',
        'prst':     'cornerTabs',
        'desc':     'Corner Tabs'
    },
    MSO.SHAPE_CROSS: {
        'basename': 'Cross',
        'prst':     'plus',
        'desc':     'Cross'
    },
    MSO.SHAPE_CUBE: {
        'basename': 'Cube',
        'prst':     'cube',
        'desc':     'Cube'
    },
    MSO.SHAPE_CURVED_DOWN_ARROW: {
        'basename': 'Curved Down Arrow',
        'prst':     'curvedDownArrow',
        'desc':     'Block arrow that curves down'
    },
    MSO.SHAPE_CURVED_DOWN_RIBBON: {
        'basename': 'Curved Down Ribbon',
        'prst':     'ellipseRibbon',
        'desc':     'Ribbon banner that curves down'
    },
    MSO.SHAPE_CURVED_LEFT_ARROW: {
        'basename': 'Curved Left Arrow',
        'prst':     'curvedLeftArrow',
        'desc':     'Block arrow that curves left'
    },
    MSO.SHAPE_CURVED_RIGHT_ARROW: {
        'basename': 'Curved Right Arrow',
        'prst':     'curvedRightArrow',
        'desc':     'Block arrow that curves right'
    },
    MSO.SHAPE_CURVED_UP_ARROW: {
        'basename': 'Curved Up Arrow',
        'prst':     'curvedUpArrow',
        'desc':     'Block arrow that curves up'
    },
    MSO.SHAPE_CURVED_UP_RIBBON: {
        'basename': 'Curved Up Ribbon',
        'prst':     'ellipseRibbon2',
        'desc':     'Ribbon banner that curves up'
    },
    MSO.SHAPE_DECAGON: {
        'basename': 'Decagon',
        'prst':     'decagon',
        'desc':     'Decagon'
    },
    MSO.SHAPE_DIAGONAL_STRIPE: {
        'basename': 'Diagonal Stripe',
        'prst':     'diagStripe',
        'desc':     'Diagonal Stripe'
    },
    MSO.SHAPE_DIAMOND: {
        'basename': 'Diamond',
        'prst':     'diamond',
        'desc':     'Diamond'
    },
    MSO.SHAPE_DODECAGON: {
        'basename': 'Dodecagon',
        'prst':     'dodecagon',
        'desc':     'Dodecagon'
    },
    MSO.SHAPE_DONUT: {
        'basename': 'Donut',
        'prst':     'donut',
        'desc':     'Donut'
    },
    MSO.SHAPE_DOUBLE_BRACE: {
        'basename': 'Double Brace',
        'prst':     'bracePair',
        'desc':     'Double brace'
    },
    MSO.SHAPE_DOUBLE_BRACKET: {
        'basename': 'Double Bracket',
        'prst':     'bracketPair',
        'desc':     'Double bracket'
    },
    MSO.SHAPE_DOUBLE_WAVE: {
        'basename': 'Double Wave',
        'prst':     'doubleWave',
        'desc':     'Double wave'
    },
    MSO.SHAPE_DOWN_ARROW: {
        'basename': 'Down Arrow',
        'prst':     'downArrow',
        'desc':     'Block arrow that points down'
    },
    MSO.SHAPE_DOWN_ARROW_CALLOUT: {
        'basename': 'Down Arrow Callout',
        'prst':     'downArrowCallout',
        'desc':     'Callout with arrow that points down'
    },
    MSO.SHAPE_DOWN_RIBBON: {
        'basename': 'Down Ribbon',
        'prst':     'ribbon',
        'desc':     'Ribbon banner with center area below ribbon ends'
    },
    MSO.SHAPE_EXPLOSION1: {
        'basename': 'Explosion',
        'prst':     'irregularSeal1',
        'desc':     'Explosion'
    },
    MSO.SHAPE_EXPLOSION2: {
        'basename': 'Explosion',
        'prst':     'irregularSeal2',
        'desc':     'Explosion'
    },
    MSO.SHAPE_FLOWCHART_ALTERNATE_PROCESS: {
        'basename': 'Alternate process',
        'prst':     'flowChartAlternateProcess',
        'desc':     'Alternate process flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_CARD: {
        'basename': 'Card',
        'prst':     'flowChartPunchedCard',
        'desc':     'Card flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_COLLATE: {
        'basename': 'Collate',
        'prst':     'flowChartCollate',
        'desc':     'Collate flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_CONNECTOR: {
        'basename': 'Connector',
        'prst':     'flowChartConnector',
        'desc':     'Connector flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_DATA: {
        'basename': 'Data',
        'prst':     'flowChartInputOutput',
        'desc':     'Data flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_DECISION: {
        'basename': 'Decision',
        'prst':     'flowChartDecision',
        'desc':     'Decision flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_DELAY: {
        'basename': 'Delay',
        'prst':     'flowChartDelay',
        'desc':     'Delay flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_DIRECT_ACCESS_STORAGE: {
        'basename': 'Direct Access Storage',
        'prst':     'flowChartMagneticDrum',
        'desc':     'Direct access storage flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_DISPLAY: {
        'basename': 'Display',
        'prst':     'flowChartDisplay',
        'desc':     'Display flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_DOCUMENT: {
        'basename': 'Document',
        'prst':     'flowChartDocument',
        'desc':     'Document flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_EXTRACT: {
        'basename': 'Extract',
        'prst':     'flowChartExtract',
        'desc':     'Extract flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_INTERNAL_STORAGE: {
        'basename': 'Internal Storage',
        'prst':     'flowChartInternalStorage',
        'desc':     'Internal storage flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_MAGNETIC_DISK: {
        'basename': 'Magnetic Disk',
        'prst':     'flowChartMagneticDisk',
        'desc':     'Magnetic disk flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_MANUAL_INPUT: {
        'basename': 'Manual Input',
        'prst':     'flowChartManualInput',
        'desc':     'Manual input flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_MANUAL_OPERATION: {
        'basename': 'Manual Operation',
        'prst':     'flowChartManualOperation',
        'desc':     'Manual operation flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_MERGE: {
        'basename': 'Merge',
        'prst':     'flowChartMerge',
        'desc':     'Merge flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_MULTIDOCUMENT: {
        'basename': 'Multidocument',
        'prst':     'flowChartMultidocument',
        'desc':     'Multi-document flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_OFFLINE_STORAGE: {
        'basename': 'Offline Storage',
        'prst':     'flowChartOfflineStorage',
        'desc':     'Offline Storage'
    },
    MSO.SHAPE_FLOWCHART_OFFPAGE_CONNECTOR: {
        'basename': 'Off-page Connector',
        'prst':     'flowChartOffpageConnector',
        'desc':     'Off-page connector flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_OR: {
        'basename': 'Or',
        'prst':     'flowChartOr',
        'desc':     '"Or" flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_PREDEFINED_PROCESS: {
        'basename': 'Predefined Process',
        'prst':     'flowChartPredefinedProcess',
        'desc':     'Predefined process flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_PREPARATION: {
        'basename': 'Preparation',
        'prst':     'flowChartPreparation',
        'desc':     'Preparation flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_PROCESS: {
        'basename': 'Process',
        'prst':     'flowChartProcess',
        'desc':     'Process flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_PUNCHED_TAPE: {
        'basename': 'Punched Tape',
        'prst':     'flowChartPunchedTape',
        'desc':     'Punched tape flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_SEQUENTIAL_ACCESS_STORAGE: {
        'basename': 'Sequential Access Storage',
        'prst':     'flowChartMagneticTape',
        'desc':     'Sequential access storage flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_SORT: {
        'basename': 'Sort',
        'prst':     'flowChartSort',
        'desc':     'Sort flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_STORED_DATA: {
        'basename': 'Stored Data',
        'prst':     'flowChartOnlineStorage',
        'desc':     'Stored data flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_SUMMING_JUNCTION: {
        'basename': 'Summing Junction',
        'prst':     'flowChartSummingJunction',
        'desc':     'Summing junction flowchart symbol'
    },
    MSO.SHAPE_FLOWCHART_TERMINATOR: {
        'basename': 'Terminator',
        'prst':     'flowChartTerminator',
        'desc':     'Terminator flowchart symbol'
    },
    MSO.SHAPE_FOLDED_CORNER: {
        'basename': 'Folded Corner',
        'prst':     'folderCorner',
        'desc':     'Folded corner'
    },
    MSO.SHAPE_FRAME: {
        'basename': 'Frame',
        'prst':     'frame',
        'desc':     'Frame'
    },
    MSO.SHAPE_FUNNEL: {
        'basename': 'Funnel',
        'prst':     'funnel',
        'desc':     'Funnel'
    },
    MSO.SHAPE_GEAR_6: {
        'basename': 'Gear 6',
        'prst':     'gear6',
        'desc':     'Gear 6'
    },
    MSO.SHAPE_GEAR_9: {
        'basename': 'Gear 9',
        'prst':     'gear9',
        'desc':     'Gear 9'
    },
    MSO.SHAPE_HALF_FRAME: {
        'basename': 'Half Frame',
        'prst':     'halfFrame',
        'desc':     'Half Frame'
    },
    MSO.SHAPE_HEART: {
        'basename': 'Heart',
        'prst':     'heart',
        'desc':     'Heart'
    },
    MSO.SHAPE_HEPTAGON: {
        'basename': 'Heptagon',
        'prst':     'heptagon',
        'desc':     'Heptagon'
    },
    MSO.SHAPE_HEXAGON: {
        'basename': 'Hexagon',
        'prst':     'hexagon',
        'desc':     'Hexagon'
    },
    MSO.SHAPE_HORIZONTAL_SCROLL: {
        'basename': 'Horizontal Scroll',
        'prst':     'horizontalScroll',
        'desc':     'Horizontal scroll'
    },
    MSO.SHAPE_ISOSCELES_TRIANGLE: {
        'basename': 'Isosceles Triangle',
        'prst':     'triangle',
        'desc':     'Isosceles triangle'
    },
    MSO.SHAPE_LEFT_ARROW: {
        'basename': 'Left Arrow',
        'prst':     'leftArrow',
        'desc':     'Block arrow that points left'
    },
    MSO.SHAPE_LEFT_ARROW_CALLOUT: {
        'basename': 'Left Arrow Callout',
        'prst':     'leftArrowCallout',
        'desc':     'Callout with arrow that points left'
    },
    MSO.SHAPE_LEFT_BRACE: {
        'basename': 'Left Brace',
        'prst':     'leftBrace',
        'desc':     'Left brace'
    },
    MSO.SHAPE_LEFT_BRACKET: {
        'basename': 'Left Bracket',
        'prst':     'leftBracket',
        'desc':     'Left bracket'
    },
    MSO.SHAPE_LEFT_CIRCULAR_ARROW: {
        'basename': 'Left Circular Arrow',
        'prst':     'leftCircularArrow',
        'desc':     'Left Circular Arrow'
    },
    MSO.SHAPE_LEFT_RIGHT_ARROW: {
        'basename': 'Left-Right Arrow',
        'prst':     'leftRightArrow',
        'desc':     ('Block arrow with arrowheads that point both left and ri'
                     'ght')
    },
    MSO.SHAPE_LEFT_RIGHT_ARROW_CALLOUT: {
        'basename': 'Left-Right Arrow Callout',
        'prst':     'leftRightArrowCallout',
        'desc':     'Callout with arrowheads that point both left and right'
    },
    MSO.SHAPE_LEFT_RIGHT_CIRCULAR_ARROW: {
        'basename': 'Left Right Circular Arrow',
        'prst':     'leftRightCircularArrow',
        'desc':     'Left Right Circular Arrow'
    },
    MSO.SHAPE_LEFT_RIGHT_RIBBON: {
        'basename': 'Left Right Ribbon',
        'prst':     'leftRightRibbon',
        'desc':     'Left Right Ribbon'
    },
    MSO.SHAPE_LEFT_RIGHT_UP_ARROW: {
        'basename': 'Left-Right-Up Arrow',
        'prst':     'leftRightUpArrow',
        'desc':     ('Block arrow with arrowheads that point left, right, and'
                     ' up')
    },
    MSO.SHAPE_LEFT_UP_ARROW: {
        'basename': 'Left-Up Arrow',
        'prst':     'leftUpArrow',
        'desc':     'Block arrow with arrowheads that point left and up'
    },
    MSO.SHAPE_LIGHTNING_BOLT: {
        'basename': 'Lightning Bolt',
        'prst':     'lightningBolt',
        'desc':     'Lightning bolt'
    },
    MSO.SHAPE_LINE_CALLOUT_1: {
        'basename': 'Line Callout 1',
        'prst':     'borderCallout1',
        'desc':     'Callout with border and horizontal callout line'
    },
    MSO.SHAPE_LINE_CALLOUT_1_ACCENT_BAR: {
        'basename': 'Line Callout 1 (Accent Bar)',
        'prst':     'accentCallout1',
        'desc':     'Callout with vertical accent bar'
    },
    MSO.SHAPE_LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 1 (Border and Accent Bar)',
        'prst':     'accentBorderCallout1',
        'desc':     'Callout with border and vertical accent bar'
    },
    MSO.SHAPE_LINE_CALLOUT_1_NO_BORDER: {
        'basename': 'Line Callout 1 (No Border)',
        'prst':     'callout1',
        'desc':     'Callout with horizontal line'
    },
    MSO.SHAPE_LINE_CALLOUT_2: {
        'basename': 'Line Callout 2',
        'prst':     'borderCallout2',
        'desc':     'Callout with diagonal straight line'
    },
    MSO.SHAPE_LINE_CALLOUT_2_ACCENT_BAR: {
        'basename': 'Line Callout 2 (Accent Bar)',
        'prst':     'accentCallout2',
        'desc':     'Callout with diagonal callout line and accent bar'
    },
    MSO.SHAPE_LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 2 (Border and Accent Bar)',
        'prst':     'accentBorderCallout2',
        'desc':     ('Callout with border, diagonal straight line, and accent'
                     ' bar')
    },
    MSO.SHAPE_LINE_CALLOUT_2_NO_BORDER: {
        'basename': 'Line Callout 2 (No Border)',
        'prst':     'callout2',
        'desc':     'Callout with no border and diagonal callout line'
    },
    MSO.SHAPE_LINE_CALLOUT_3: {
        'basename': 'Line Callout 3',
        'prst':     'borderCallout3',
        'desc':     'Callout with angled line'
    },
    MSO.SHAPE_LINE_CALLOUT_3_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Accent Bar)',
        'prst':     'accentCallout3',
        'desc':     'Callout with angled callout line and accent bar'
    },
    MSO.SHAPE_LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Border and Accent Bar)',
        'prst':     'accentBorderCallout3',
        'desc':     'Callout with border, angled callout line, and accent bar'
    },
    MSO.SHAPE_LINE_CALLOUT_3_NO_BORDER: {
        'basename': 'Line Callout 3 (No Border)',
        'prst':     'callout3',
        'desc':     'Callout with no border and angled callout line'
    },
    MSO.SHAPE_LINE_CALLOUT_4: {
        'basename': 'Line Callout 3',
        'prst':     'borderCallout3',
        'desc':     'Callout with callout line segments forming a U-shape.'
    },
    MSO.SHAPE_LINE_CALLOUT_4_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Accent Bar)',
        'prst':     'accentCallout3',
        'desc':     ('Callout with accent bar and callout line segments formi'
                     'ng a U-shape.')
    },
    MSO.SHAPE_LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR: {
        'basename': 'Line Callout 3 (Border and Accent Bar)',
        'prst':     'accentBorderCallout3',
        'desc':     ('Callout with border, accent bar, and callout line segme'
                     'nts forming a U-shape.')
    },
    MSO.SHAPE_LINE_CALLOUT_4_NO_BORDER: {
        'basename': 'Line Callout 3 (No Border)',
        'prst':     'callout3',
        'desc':     ('Callout with no border and callout line segments formin'
                     'g a U-shape.')
    },
    MSO.SHAPE_LINE_INVERSE: {
        'basename': 'Straight Connector',
        'prst':     'lineInv',
        'desc':     'Straight Connector'
    },
    MSO.SHAPE_MATH_DIVIDE: {
        'basename': 'Division',
        'prst':     'mathDivide',
        'desc':     'Division'
    },
    MSO.SHAPE_MATH_EQUAL: {
        'basename': 'Equal',
        'prst':     'mathEqual',
        'desc':     'Equal'
    },
    MSO.SHAPE_MATH_MINUS: {
        'basename': 'Minus',
        'prst':     'mathMinus',
        'desc':     'Minus'
    },
    MSO.SHAPE_MATH_MULTIPLY: {
        'basename': 'Multiply',
        'prst':     'mathMultiply',
        'desc':     'Multiply'
    },
    MSO.SHAPE_MATH_NOT_EQUAL: {
        'basename': 'Not Equal',
        'prst':     'mathNotEqual',
        'desc':     'Not Equal'
    },
    MSO.SHAPE_MATH_PLUS: {
        'basename': 'Plus',
        'prst':     'mathPlus',
        'desc':     'Plus'
    },
    MSO.SHAPE_MOON: {
        'basename': 'Moon',
        'prst':     'moon',
        'desc':     'Moon'
    },
    MSO.SHAPE_NO_SYMBOL: {
        'basename': '"No" symbol',
        'prst':     'noSmoking',
        'desc':     '"No" symbol'
    },
    MSO.SHAPE_NON_ISOSCELES_TRAPEZOID: {
        'basename': 'Non-isosceles Trapezoid',
        'prst':     'nonIsoscelesTrapezoid',
        'desc':     'Non-isosceles Trapezoid'
    },
    MSO.SHAPE_NOTCHED_RIGHT_ARROW: {
        'basename': 'Notched Right Arrow',
        'prst':     'notchedRightArrow',
        'desc':     'Notched block arrow that points right'
    },
    MSO.SHAPE_OCTAGON: {
        'basename': 'Octagon',
        'prst':     'octagon',
        'desc':     'Octagon'
    },
    MSO.SHAPE_OVAL: {
        'basename': 'Oval',
        'prst':     'ellipse',
        'desc':     'Oval'
    },
    MSO.SHAPE_OVAL_CALLOUT: {
        'basename': 'Oval Callout',
        'prst':     'wedgeEllipseCallout',
        'desc':     'Oval-shaped callout'
    },
    MSO.SHAPE_PARALLELOGRAM: {
        'basename': 'Parallelogram',
        'prst':     'parallelogram',
        'desc':     'Parallelogram'
    },
    MSO.SHAPE_PENTAGON: {
        'basename': 'Pentagon',
        'prst':     'homePlate',
        'desc':     'Pentagon'
    },
    MSO.SHAPE_PIE: {
        'basename': 'Pie',
        'prst':     'pie',
        'desc':     'Pie'
    },
    MSO.SHAPE_PIE_WEDGE: {
        'basename': 'Pie',
        'prst':     'pieWedge',
        'desc':     'Pie'
    },
    MSO.SHAPE_PLAQUE: {
        'basename': 'Plaque',
        'prst':     'plaque',
        'desc':     'Plaque'
    },
    MSO.SHAPE_PLAQUE_TABS: {
        'basename': 'Plaque Tabs',
        'prst':     'plaqueTabs',
        'desc':     'Plaque Tabs'
    },
    MSO.SHAPE_QUAD_ARROW: {
        'basename': 'Quad Arrow',
        'prst':     'quadArrow',
        'desc':     'Block arrows that point up, down, left, and right'
    },
    MSO.SHAPE_QUAD_ARROW_CALLOUT: {
        'basename': 'Quad Arrow Callout',
        'prst':     'quadArrowCallout',
        'desc':     'Callout with arrows that point up, down, left, and right'
    },
    MSO.SHAPE_RECTANGLE: {
        'basename': 'Rectangle',
        'prst':     'rect',
        'desc':     'Rectangle'
    },
    MSO.SHAPE_RECTANGULAR_CALLOUT: {
        'basename': 'Rectangular Callout',
        'prst':     'wedgeRectCallout',
        'desc':     'Rectangular callout'
    },
    MSO.SHAPE_REGULAR_PENTAGON: {
        'basename': 'Regular Pentagon',
        'prst':     'pentagon',
        'desc':     'Pentagon'
    },
    MSO.SHAPE_RIGHT_ARROW: {
        'basename': 'Right Arrow',
        'prst':     'rightArrow',
        'desc':     'Block arrow that points right'
    },
    MSO.SHAPE_RIGHT_ARROW_CALLOUT: {
        'basename': 'Right Arrow Callout',
        'prst':     'rightArrowCallout',
        'desc':     'Callout with arrow that points right'
    },
    MSO.SHAPE_RIGHT_BRACE: {
        'basename': 'Right Brace',
        'prst':     'rightBrace',
        'desc':     'Right brace'
    },
    MSO.SHAPE_RIGHT_BRACKET: {
        'basename': 'Right Bracket',
        'prst':     'rightBracket',
        'desc':     'Right bracket'
    },
    MSO.SHAPE_RIGHT_TRIANGLE: {
        'basename': 'Right Triangle',
        'prst':     'rtTriangle',
        'desc':     'Right triangle'
    },
    MSO.SHAPE_ROUND_1_RECTANGLE: {
        'basename': 'Round Single Corner Rectangle',
        'prst':     'round1Rect',
        'desc':     'Round Single Corner Rectangle'
    },
    MSO.SHAPE_ROUND_2_DIAG_RECTANGLE: {
        'basename': 'Round Diagonal Corner Rectangle',
        'prst':     'round2DiagRect',
        'desc':     'Round Diagonal Corner Rectangle'
    },
    MSO.SHAPE_ROUND_2_SAME_RECTANGLE: {
        'basename': 'Round Same Side Corner Rectangle',
        'prst':     'round2SameRect',
        'desc':     'Round Same Side Corner Rectangle'
    },
    MSO.SHAPE_ROUNDED_RECTANGLE: {
        'basename': 'Rounded Rectangle',
        'prst':     'roundRect',
        'desc':     'Rounded rectangle'
    },
    MSO.SHAPE_ROUNDED_RECTANGULAR_CALLOUT: {
        'basename': 'Rounded Rectangular Callout',
        'prst':     'wedgeRoundRectCallout',
        'desc':     'Rounded rectangle-shaped callout'
    },
    MSO.SHAPE_SMILEY_FACE: {
        'basename': 'Smiley Face',
        'prst':     'smileyFace',
        'desc':     'Smiley face'
    },
    MSO.SHAPE_SNIP_1_RECTANGLE: {
        'basename': 'Snip Single Corner Rectangle',
        'prst':     'snip1Rect',
        'desc':     'Snip Single Corner Rectangle'
    },
    MSO.SHAPE_SNIP_2_DIAG_RECTANGLE: {
        'basename': 'Snip Diagonal Corner Rectangle',
        'prst':     'snip2DiagRect',
        'desc':     'Snip Diagonal Corner Rectangle'
    },
    MSO.SHAPE_SNIP_2_SAME_RECTANGLE: {
        'basename': 'Snip Same Side Corner Rectangle',
        'prst':     'snip2SameRect',
        'desc':     'Snip Same Side Corner Rectangle'
    },
    MSO.SHAPE_SNIP_ROUND_RECTANGLE: {
        'basename': 'Snip and Round Single Corner Rectangle',
        'prst':     'snipRoundRect',
        'desc':     'Snip and Round Single Corner Rectangle'
    },
    MSO.SHAPE_SQUARE_TABS: {
        'basename': 'Square Tabs',
        'prst':     'squareTabs',
        'desc':     'Square Tabs'
    },
    MSO.SHAPE_STRIPED_RIGHT_ARROW: {
        'basename': 'Striped Right Arrow',
        'prst':     'stripedRightArrow',
        'desc':     'Block arrow that points right with stripes at the tail'
    },
    MSO.SHAPE_SUN: {
        'basename': 'Sun',
        'prst':     'sun',
        'desc':     'Sun'
    },
    MSO.SHAPE_SWOOSH_ARROW: {
        'basename': 'Swoosh Arrow',
        'prst':     'swooshArrow',
        'desc':     'Swoosh Arrow'
    },
    MSO.SHAPE_TEAR: {
        'basename': 'Teardrop',
        'prst':     'teardrop',
        'desc':     'Teardrop'
    },
    MSO.SHAPE_TRAPEZOID: {
        'basename': 'Trapezoid',
        'prst':     'trapezoid',
        'desc':     'Trapezoid'
    },
    MSO.SHAPE_U_TURN_ARROW: {
        'basename': 'U-Turn Arrow',
        'prst':     'uturnArrow',
        'desc':     'Block arrow forming a U shape'
    },
    MSO.SHAPE_UP_ARROW: {
        'basename': 'Up Arrow',
        'prst':     'upArrow',
        'desc':     'Block arrow that points up'
    },
    MSO.SHAPE_UP_ARROW_CALLOUT: {
        'basename': 'Up Arrow Callout',
        'prst':     'upArrowCallout',
        'desc':     'Callout with arrow that points up'
    },
    MSO.SHAPE_UP_DOWN_ARROW: {
        'basename': 'Up-Down Arrow',
        'prst':     'upDownArrow',
        'desc':     'Block arrow that points up and down'
    },
    MSO.SHAPE_UP_DOWN_ARROW_CALLOUT: {
        'basename': 'Up-Down Arrow Callout',
        'prst':     'upDownArrowCallout',
        'desc':     'Callout with arrows that point up and down'
    },
    MSO.SHAPE_UP_RIBBON: {
        'basename': 'Up Ribbon',
        'prst':     'ribbon2',
        'desc':     'Ribbon banner with center area above ribbon ends'
    },
    MSO.SHAPE_VERTICAL_SCROLL: {
        'basename': 'Vertical Scroll',
        'prst':     'verticalScroll',
        'desc':     'Vertical scroll'
    },
    MSO.SHAPE_WAVE: {
        'basename': 'Wave',
        'prst':     'wave',
        'desc':     'Wave'
    }
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
# PresentationML Part Type specs
# ============================================================================
# Keyed by content type
# Not yet included:
# * Font Part (font1.fntdata) 15.2.13
# * themeOverride : 'application/vnd.openxmlformats-officedocument.'
#                   'themeOverride+xml'
# * several others, especially DrawingML parts ...
#
# TODO: Also check out other shared parts in section 15.
# ============================================================================

PTS_CARDINALITY_SINGLETON = 'singleton'
PTS_CARDINALITY_TUPLE = 'tuple'
PTS_HASRELS_ALWAYS = 'always'
PTS_HASRELS_NEVER = 'never'
PTS_HASRELS_OPTIONAL = 'optional'

CT_CHART = (
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
CT_COMMENT_AUTHORS = (
    'application/vnd.openxmlformats-officedocument.presentationml.commentAuth'
    'ors+xml')
CT_COMMENTS = (
    'application/vnd.openxmlformats-officedocument.presentationml.comments+xm'
    'l')
CT_CORE_PROPS = (
    'application/vnd.openxmlformats-package.core-properties+xml')
CT_CUSTOM_PROPS = (
    'application/vnd.openxmlformats-officedocument.custom-properties+xml')
CT_EXCEL_XLSX = (
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
CT_EXTENDED_PROPS = (
    'application/vnd.openxmlformats-officedocument.extended-properties+xml')
CT_HANDOUT_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.handoutMast'
    'er+xml')
CT_NOTES_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.notesMaster'
    '+xml')
CT_NOTES_SLIDE = (
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+'
    'xml')
CT_PRES_PROPS = (
    'application/vnd.openxmlformats-officedocument.presentationml.presProps+x'
    'ml')
CT_PRESENTATION = (
    'application/vnd.openxmlformats-officedocument.presentationml.presentatio'
    'n.main+xml')
CT_PRINTER_SETTINGS = (
    'application/vnd.openxmlformats-officedocument.presentationml.printerSett'
    'ings')
CT_SLIDE = (
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
CT_SLIDE_LAYOUT = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideLayout'
    '+xml')
CT_SLIDE_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideMaster'
    '+xml')
CT_SLIDESHOW = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow.m'
    'ain+xml')
CT_TABLE_STYLES = (
    'application/vnd.openxmlformats-officedocument.presentationml.tableStyles'
    '+xml')
CT_TAGS = (
    'application/vnd.openxmlformats-officedocument.presentationml.tags+xml')
CT_TEMPLATE = (
    'application/vnd.openxmlformats-officedocument.presentationml.template.ma'
    'in+xml')
CT_THEME = (
    'application/vnd.openxmlformats-officedocument.theme+xml')
CT_VIEW_PROPS = (
    'application/vnd.openxmlformats-officedocument.presentationml.viewProps+x'
    'ml')
CT_WORKSHEET = (
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


RT_CHART = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/char'
    't')
RT_COMMENT_AUTHORS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'entAuthors')
RT_COMMENTS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'ents')
RT_CORE_PROPS = (
    'http://schemas.openxmlformats.org/officedocument/2006/relationships/meta'
    'data/core-properties')
RT_CUSTOM_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/cust'
    'omProperties')
RT_EXTENDED_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/exte'
    'ndedProperties')
RT_HANDOUT_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hand'
    'outMaster')
RT_IMAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/imag'
    'e')
RT_NOTES_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/note'
    'sMaster')
RT_NOTES_SLIDE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/note'
    'sSlide')
RT_OFFICE_DOCUMENT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument')
RT_PACKAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pack'
    'age')
RT_PRES_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pres'
    'Props')
RT_PRINTER_SETTINGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/prin'
    'terSettings')
RT_SLIDE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'e')
RT_SLIDE_LAYOUT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'eLayout')
RT_SLIDE_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'eMaster')
RT_SLIDESHOW = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument')
RT_TABLESTYLES = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tabl'
    'eStyles')
RT_TAGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags')
RT_TEMPLATE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument')
RT_THEME = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/them'
    'e')
RT_VIEWPROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/view'
    'Props')

pml_parttypes = {
    CT_COMMENT_AUTHORS: {  # ECMA-376-1 13.3.1
        'basename':    'commentAuthors',
        'ext':         '.xml',
        'name':        'Comment Authors Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_COMMENT_AUTHORS},
    CT_COMMENTS: {  # ECMA-376-1 13.3.2
        'basename':    'comment',
        'ext':         '.xml',
        'name':        'Comments Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/comments',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['slide'],
        'reltype':     RT_COMMENTS},
    CT_CORE_PROPS: {  # ECMA-376-1 15.2.12.1 ('Core' as in Dublin Core)
        'basename':    'core',
        'ext':         '.xml',
        'name':        'Core File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_CORE_PROPS},
    CT_CUSTOM_PROPS: {  # ECMA-376-1 15.2.12.2
        'basename':    'custom',
        'ext':         '.xml',
        'name':        'Custom File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_CUSTOM_PROPS},
    CT_EXTENDED_PROPS: {  # ECMA-376-1 15.2.12.3 (Extended File Properties)
        'basename':    'app',
        'ext':         '.xml',
        'name':        'Application-Defined File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_EXTENDED_PROPS},
    CT_HANDOUT_MASTER: {  # ECMA-376-1 13.3.3
        'basename':    'handoutMaster',
        'ext':         '.xml',
        'name':        'Handout Master Part',
        # actually can only be one according to spec, but behaves like part
        # collection (handoutMasters folder, handoutMaster1.xml, etc.)
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/handoutMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation'],
        'reltype':     RT_HANDOUT_MASTER},
    CT_NOTES_MASTER: {  # ECMA-376-1 13.3.4
        'basename':    'notesMaster',
        'ext':         '.xml',
        'name':        'Notes Master Part',
        # actually can only be one according to spec, but behaves like part
        # collection (notesMasters folder, notesMaster1.xml, etc.)
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/notesMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'notesSlide'],
        'reltype':     RT_NOTES_MASTER},
    CT_NOTES_SLIDE: {  # ECMA-376-1 13.3.5
        'basename':    'notesSlide',
        'ext':         '.xml',
        'name':        'Notes Slide Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/notesSlides',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['slide'],
        'reltype':     RT_NOTES_SLIDE},
    CT_PRESENTATION: {  # ECMA-376-1 13.3.6
        # one of three possible Content Type values for presentation part
        'basename':    'presentation',
        'ext':         '.xml',
        'name':        'Presentation Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['package'],
        'reltype':     RT_OFFICE_DOCUMENT},
    CT_PRES_PROPS: {  # ECMA-376-1 13.3.7
        'basename':    'presProps',
        'ext':         '.xml',
        'name':        'Presentation Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_PRES_PROPS},
    CT_PRINTER_SETTINGS: {  # ECMA-376-1 15.2.15
        'basename':    'printerSettings',
        'ext':         '.bin',
        'name':        'Printer Settings Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/printerSettings',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_PRINTER_SETTINGS},
    CT_SLIDE: {  # ECMA-376-1 13.3.8
        'basename':    'slide',
        'ext':         '.xml',
        'name':        'Slide Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/slides',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'notesSlide'],
        'reltype':     RT_SLIDE},
    CT_SLIDE_LAYOUT: {  # ECMA-376-1 13.3.9
        'basename':    'slideLayout',
        'ext':         '.xml',
        'name':        'Slide Layout Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    True,
        'baseURI':     '/ppt/slideLayouts',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['slide', 'slideMaster'],
        'reltype':     RT_SLIDE_LAYOUT},
    CT_SLIDE_MASTER: {  # ECMA-376-1 13.3.10
        'basename':    'slideMaster',
        'ext':         '.xml',
        'name':        'Slide Master Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    True,
        'baseURI':     '/ppt/slideMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'slideLayout'],
        'reltype':     RT_SLIDE_MASTER},
    CT_SLIDESHOW: {  # ECMA-376-1 13.3.6
        # one of three possible Content Type values for presentation part
        'basename':    'presentation',
        'ext':         '.xml',
        'name':        'Presentation Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['package'],
        'reltype':     RT_SLIDESHOW},
    CT_TABLE_STYLES: {  # ECMA-376-1 14.2.9
        'basename':    'tableStyles',
        'ext':         '.xml',
        'name':        'Table Styles Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_TABLESTYLES},
    CT_TAGS: {  # ECMA-376-1 13.3.12
        'basename':    'tag',
        'ext':         '.xml',
        'name':        'User-Defined Tags Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/tags',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation', 'slide'],
        'reltype':     RT_TAGS},
    CT_TEMPLATE: {  # ECMA-376-1 13.3.6
        # one of three possible Content Type values for presentation part
        'basename':    'presentation',
        'ext':         '.xml',
        'name':        'Presentation Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['package'],
        'reltype':     RT_TEMPLATE},
    CT_THEME: {  # ECMA-376-1 14.2.7
        'basename':    'theme',
        'ext':         '.xml',
        'name':        'Theme Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        # spec indicates theme part is optional, but I've never seen a .pptx
        # without one
        'required':    True,
        'baseURI':     '/ppt/theme',
        # can have _rels items, but only if theme contains one or more images
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['presentation', 'handoutMaster', 'notesMaster',
                        'slideMaster'],
        'reltype':     RT_THEME},
    CT_VIEW_PROPS: {  # ECMA-376-1 13.3.13
        'basename':    'viewProps',
        'ext':         '.xml',
        'name':        'View Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_VIEWPROPS},
    'image/gif': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.gif',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    'image/jpeg': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.jpeg',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    'image/png': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.png',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE},
    'image/x-emf': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.emf',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE}}


# ============================================================================
# default_content_types
# ============================================================================
# Default file extension to MIME type mapping. This is used as a reference for
# adding <Default> elements to [Content_Types].xml for parts like media files.
#     TODO: I've seen .wmv elements in the media folder of at least one
# presentation, might need to add an entry for that and perhaps other rich
# media PowerPoint allows to be embedded (e.g. audio, movie, object, ...).
# ============================================================================

default_content_types = {
    '.bin':     CT_PRINTER_SETTINGS,
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
    '.wmf':     'image/x-wmf',
    '.xlsx':    CT_EXCEL_XLSX,
    '.xml':     'application/xml'}


# ============================================================================
# nsmap
# ============================================================================
# namespace prefix to namespace name map
# ============================================================================

nsmap = {
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'cp':  ('http://schemas.openxmlformats.org/package/2006/metadata/core-pro'
            'perties'),
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/',
    'ep':  ('http://schemas.openxmlformats.org/officeDocument/2006/extended-p'
            'roperties'),
    'i':   RT_IMAGE,
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mo':  'http://schemas.microsoft.com/office/mac/office/2008/main',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'o':   'urn:schemas-microsoft-com:office:office',
    'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
    'pd':  ('http://schemas.openxmlformats.org/drawingml/2006/presentationDra'
            'wing'),
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
    'r':   ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips'),
    'sl':  ('http://schemas.openxmlformats.org/officeDocument/2006/relationsh'
            'ips/slideLayout'),
    'v':   'urn:schemas-microsoft-com:vml',
    've':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingD'
            'rawing'),
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}


def namespaces(*prefixes):
    """
    Return a dict containing the subset namespace prefix mappings specified by
    *prefixes*. Any number of namespace prefixes can be supplied, e.g.
    namespaces('a', 'r', 'p').
    """
    namespaces = {}
    for prefix in prefixes:
        namespaces[prefix] = nsmap[prefix]
    return namespaces


def qtag(tag):
    """
    Return a qualified name (QName) for an XML element or attribute in Clark
    notation, e.g. ``'{http://www.w3.org/1999/xhtml}body'`` instead of
    ``'html:body'``, by looking up the specified namespace prefix in the
    overall namespace map (nsmap) above. Google on "xml clark notation" for
    more on Clark notation. *tag* is a namespace-prefixed tagname, e.g.
    ``'p:cSld'``.
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{%s}%s' % (uri, tagroot)
