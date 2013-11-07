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

from constants import (
    MSO_AUTO_SHAPE_TYPE as MAST, MSO, PP, TEXT_ALIGN_TYPE as TAT,
    TEXT_ANCHORING_TYPE as TANC
)


class VerticalAnchor(object):
    """
    Mappings between ``MsoVerticalAnchor`` values used in the API and
    ``ST_TextAnchoringType`` values used in the XML. ``MsoVerticalAnchor``
    values are like ``MSO.ANCHOR_MIDDLE``.
    """
    _mapping = {
        MSO.ANCHOR_TOP:    TANC.TOP,
        MSO.ANCHOR_MIDDLE: TANC.MIDDLE,
        MSO.ANCHOR_BOTTOM: TANC.BOTTOM
    }

    @classmethod
    def from_text_anchoring_type(cls, text_anchoring_type):
        """
        Map an ``ST_TextAnchoringType`` value (e.g. ``TANC.TOP`` or
        ``'t'``) to an MsoVerticalAnchor value (e.g. ``MSO.ANCHOR_TOP``).
        Returns |None| if *text_anchoring_type* is |None|.
        """
        if text_anchoring_type is None:
            return None
        for vertical_anchor, tanc in cls._mapping.iteritems():
            if tanc == text_anchoring_type:
                return vertical_anchor
        tmpl = "no vertical anchor type for ST_TextAnchoringType '%s'"
        raise KeyError(tmpl % text_anchoring_type)

    @classmethod
    def to_text_anchoring_type(cls, vertical_anchor):
        """
        Map an ``MsoVerticalAnchor`` value (e.g. ``MSO.ANCHOR_MIDDLE``) to an
        ``ST_TextAnchoringType`` value (e.g. ``TANC.MIDDLE`` or ``'ctr'``).
        Returns None if *vertical_anchor* is None.
        """
        if vertical_anchor is None:
            return None
        try:
            text_anchoring_type = cls._mapping[vertical_anchor]
        except KeyError:
            tmpl = "no ST_TextAnchoringType value for vertical_anchor '%s'"
            raise KeyError(tmpl % vertical_anchor)
        return text_anchoring_type


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
    def from_text_align_type(cls, text_align_type):
        """
        Map an ``ST_TextAlignType`` value (e.g. ``TAT.CENTER`` or ``'ctr'``)
        to a paragraph alignment value (e.g. ``PP.ALIGN_CENTER``). Returns
        None if *text_align_type* is None.
        """
        if text_align_type is None:
            return None
        for alignment, tat in cls._mapping.iteritems():
            if tat == text_align_type:
                return alignment
        tmpl = "no ST_TextAlignType '%s'"
        raise KeyError(tmpl % text_align_type)

    @classmethod
    def to_text_align_type(cls, alignment):
        """
        Map a paragraph alignment value (e.g. ``PP.ALIGN_CENTER``) to an
        ``ST_TextAlignType`` value (e.g. ``TAT.CENTER`` or ``'ctr'``).
        Returns None if *alignment* is None.
        """
        if alignment is None:
            return None
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
# PresentationML Part Type specs
# ============================================================================
# Keyed by content type
# Not yet included:
# * Font Part (font1.fntdata) 15.2.13
# * themeOverride : 'application/vnd.openxmlformats-officedocument.'
#                   'themeOverride+xml'
# * several others, especially DrawingML parts ...
#
# BACKLOG: Also check out other shared parts in section 15.
# ============================================================================

PTS_CARDINALITY_SINGLETON = 'singleton'
PTS_CARDINALITY_TUPLE = 'tuple'
PTS_HASRELS_ALWAYS = 'always'
PTS_HASRELS_NEVER = 'never'
PTS_HASRELS_OPTIONAL = 'optional'

CT_CHART = (
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
)
CT_COMMENT_AUTHORS = (
    'application/vnd.openxmlformats-officedocument.presentationml.commentAuth'
    'ors+xml'
)
CT_COMMENTS = (
    'application/vnd.openxmlformats-officedocument.presentationml.comments+xm'
    'l'
)
CT_CORE_PROPS = (
    'application/vnd.openxmlformats-package.core-properties+xml'
)
CT_CUSTOM_PROPS = (
    'application/vnd.openxmlformats-officedocument.custom-properties+xml'
)
CT_CUSTOM_XML = (
    'application/xml'
)
CT_CUSTOM_XML_PROPS = (
    'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'
)
CT_DOCUMENT = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.'
    'main+xml'
)
CT_ENDNOTES = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+'
    'xml'
)
CT_EXTENDED_PROPS = (
    'application/vnd.openxmlformats-officedocument.extended-properties+xml'
)
CT_FONT_TABLE = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable'
    '+xml'
)
CT_FOOTER = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xm'
    'l'
)
CT_FOOTNOTES = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes'
    '+xml'
)
CT_GLOSSARY_DOCUMENT = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.'
    'glossary+xml'
)
CT_HANDOUT_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.handoutMast'
    'er+xml'
)
CT_HEADER = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xm'
    'l'
)
CT_NOTES_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.notesMaster'
    '+xml'
)
CT_NOTES_SLIDE = (
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+'
    'xml'
)
CT_NUMBERING = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering'
    '+xml'
)
CT_PRES_PROPS = (
    'application/vnd.openxmlformats-officedocument.presentationml.presProps+x'
    'ml'
)
CT_PRESENTATION = (
    'application/vnd.openxmlformats-officedocument.presentationml.presentatio'
    'n.main+xml'
)
CT_PRINTER_SETTINGS = (
    'application/vnd.openxmlformats-officedocument.presentationml.printerSett'
    'ings'
)
CT_SETTINGS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+'
    'xml'
)
CT_SLIDE = (
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
)
CT_SLIDE_LAYOUT = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideLayout'
    '+xml'
)
CT_SLIDE_MASTER = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideMaster'
    '+xml'
)
CT_SLIDESHOW = (
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow.m'
    'ain+xml'
)
CT_STYLES = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xm'
    'l'
)
CT_TABLE_STYLES = (
    'application/vnd.openxmlformats-officedocument.presentationml.tableStyles'
    '+xml'
)
CT_TAGS = (
    'application/vnd.openxmlformats-officedocument.presentationml.tags+xml'
)
CT_TEMPLATE = (
    'application/vnd.openxmlformats-officedocument.presentationml.template.ma'
    'in+xml'
)
CT_THEME = (
    'application/vnd.openxmlformats-officedocument.theme+xml'
)
CT_VIEW_PROPS = (
    'application/vnd.openxmlformats-officedocument.presentationml.viewProps+x'
    'ml'
)
CT_WEB_SETTINGS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettin'
    'gs+xml'
)
CT_WML_COMMENTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+'
    'xml'
)
CT_WORKSHEET = (
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)


RT_CHART = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/char'
    't'
)
RT_COMMENT_AUTHORS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'entAuthors'
)
RT_COMMENTS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'ents'
)
RT_CORE_PROPS = (
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/co'
    're-properties'
)
RT_CUSTOM_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/cust'
    'omProperties'
)
RT_CUSTOM_XML = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/cus'
    'tomXml'
)
RT_CUSTOM_XML_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/cus'
    'tomXmlProps'
)
RT_ENDNOTES = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endn'
    'otes'
)
RT_EXTENDED_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/exte'
    'ndedProperties'
)
RT_FONT_TABLE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'
    'Table'
)
RT_FOOTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/foot'
    'er'
)
RT_FOOTNOTES = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/foot'
    'notes'
)
RT_GLOSSARY_DOCUMENT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glos'
    'saryDocument'
)
RT_HANDOUT_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hand'
    'outMaster'
)
RT_HEADER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/head'
    'er'
)
RT_IMAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/imag'
    'e'
)
RT_NOTES_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/note'
    'sMaster'
)
RT_NOTES_SLIDE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/note'
    'sSlide'
)
RT_NUMBERING = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numb'
    'ering'
)
RT_OFFICE_DOCUMENT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument'
)
RT_PACKAGE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pack'
    'age'
)
RT_PRES_PROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pres'
    'Props'
)
RT_PRINTER_SETTINGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/prin'
    'terSettings'
)
RT_SETTINGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sett'
    'ings'
)
RT_SLIDE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'e'
)
RT_SLIDE_LAYOUT = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'eLayout'
)
RT_SLIDE_MASTER = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slid'
    'eMaster'
)
RT_SLIDESHOW = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument'
)
RT_STYLES = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styl'
    'es'
)
RT_TABLESTYLES = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tabl'
    'eStyles'
)
RT_TAGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'
)
RT_TEMPLATE = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/offi'
    'ceDocument'
)
RT_THEME = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/them'
    'e'
)
RT_VIEWPROPS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/view'
    'Props'
)
RT_WEB_SETTINGS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webS'
    'ettings'
)
RT_WML_COMMENTS = (
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comm'
    'ents'
)

pml_parttypes = {
    CT_CHART: {  # ISO/IEC 29500-1 14.2.1
        'basename':    'chart',
        'ext':         '.xml',
        'name':        'Chart Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/charts',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['presentation'],
        'reltype':     RT_CHART
    },
    CT_COMMENT_AUTHORS: {  # ECMA-376-1 13.3.1
        'basename':    'commentAuthors',
        'ext':         '.xml',
        'name':        'Comment Authors Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_COMMENT_AUTHORS
    },
    CT_COMMENTS: {  # ECMA-376-1 13.3.2
        'basename':    'comment',
        'ext':         '.xml',
        'name':        'Comments Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/comments',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['slide'],
        'reltype':     RT_COMMENTS
    },
    CT_CORE_PROPS: {  # ECMA-376-1 15.2.12.1 ('Core' as in Dublin Core)
        'basename':    'core',
        'ext':         '.xml',
        'name':        'Core File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_CORE_PROPS
    },
    CT_CUSTOM_PROPS: {  # ECMA-376-1 15.2.12.2
        'basename':    'custom',
        'ext':         '.xml',
        'name':        'Custom File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_CUSTOM_PROPS
    },
    CT_CUSTOM_XML: {  # ISO/IEC 29500-1 15.2.5
        'basename':    'item',
        'ext':         '.xml',
        'name':        'Custom XML Data Storage Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/customXML',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   [],
        'reltype':     RT_CUSTOM_XML
    },
    CT_CUSTOM_XML_PROPS: {  # ISO/IEC 29500-1 15.2.6
        'basename':    'itemProps',
        'ext':         '.xml',
        'name':        'Custom XML Data Storage Properties Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/customXML',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   [],
        'reltype':     RT_CUSTOM_XML_PROPS
    },
    CT_DOCUMENT: {  # ISO/IEC 29500-1 11.3.10
        'basename':    'document',
        'ext':         '.xml',
        'name':        'Main Document Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['package'],
        'reltype':     RT_OFFICE_DOCUMENT
    },
    CT_ENDNOTES: {  # ISO/IEC 29500-1 11.3.4
        'basename':    'endnotes',
        'ext':         '.xml',
        'name':        'Endnotes Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document', 'glossary'],
        'reltype':     RT_ENDNOTES
    },
    CT_EXTENDED_PROPS: {  # ECMA-376-1 15.2.12.3 (Extended File Properties)
        'basename':    'app',
        'ext':         '.xml',
        'name':        'Application-Defined File Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/docProps',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['package'],
        'reltype':     RT_EXTENDED_PROPS
    },
    CT_FONT_TABLE: {  # ISO/IEC 29500-1 11.3.5
        'basename':    'fontTable',
        'ext':         '.xml',
        'name':        'Font Table Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document'],
        'reltype':     RT_FONT_TABLE
    },
    CT_FOOTER: {  # ISO/IEC 29500-1 11.3.6
        'basename':    'footer',
        'ext':         '.xml',
        'name':        'Footer Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document', 'glossary'],
        'reltype':     RT_FOOTER
    },
    CT_FOOTNOTES: {  # ISO/IEC 29500-1 11.3.7
        'basename':    'footnotes',
        'ext':         '.xml',
        'name':        'Footnotes Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document'],
        'reltype':     RT_FOOTNOTES
    },
    CT_GLOSSARY_DOCUMENT: {  # ISO/IEC 29500-1 11.3.8
        'basename':    'document',
        'ext':         '.xml',
        'name':        'Glossary Document Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word/glossary',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document'],
        'reltype':     RT_GLOSSARY_DOCUMENT
    },
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
        'reltype':     RT_HANDOUT_MASTER
    },
    CT_HEADER: {  # ISO/IEC 29500-1 11.3.9
        'basename':    'header',
        'ext':         '.xml',
        'name':        'Header Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document', 'glossary'],
        'reltype':     RT_HEADER
    },
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
        'reltype':     RT_NOTES_MASTER
    },
    CT_NOTES_SLIDE: {  # ECMA-376-1 13.3.5
        'basename':    'notesSlide',
        'ext':         '.xml',
        'name':        'Notes Slide Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/notesSlides',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['slide'],
        'reltype':     RT_NOTES_SLIDE
    },
    CT_NUMBERING: {  # ISO/IEC 29500-1 11.3.11
        'basename':    'numbering',
        'ext':         '.xml',
        'name':        'Numbering Definitions Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document'],
        'reltype':     RT_NUMBERING
    },
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
        'reltype':     RT_OFFICE_DOCUMENT
    },
    CT_PRES_PROPS: {  # ECMA-376-1 13.3.7
        'basename':    'presProps',
        'ext':         '.xml',
        'name':        'Presentation Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    True,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_PRES_PROPS
    },
    CT_PRINTER_SETTINGS: {  # ECMA-376-1 15.2.15
        'basename':    'printerSettings',
        'ext':         '.bin',
        'name':        'Printer Settings Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/printerSettings',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_PRINTER_SETTINGS
    },
    CT_SETTINGS: {  # ISO/IEC 29500-1 11.3.3
        'basename':    'settings',
        'ext':         '.xml',
        'name':        'Document Settings Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document', 'glossary'],
        'reltype':     RT_SETTINGS
    },
    CT_SLIDE: {  # ECMA-376-1 13.3.8
        'basename':    'slide',
        'ext':         '.xml',
        'name':        'Slide Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/slides',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'notesSlide'],
        'reltype':     RT_SLIDE
    },
    CT_SLIDE_LAYOUT: {  # ECMA-376-1 13.3.9
        'basename':    'slideLayout',
        'ext':         '.xml',
        'name':        'Slide Layout Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    True,
        'baseURI':     '/ppt/slideLayouts',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['slide', 'slideMaster'],
        'reltype':     RT_SLIDE_LAYOUT
    },
    CT_SLIDE_MASTER: {  # ECMA-376-1 13.3.10
        'basename':    'slideMaster',
        'ext':         '.xml',
        'name':        'Slide Master Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    True,
        'baseURI':     '/ppt/slideMasters',
        'has_rels':    PTS_HASRELS_ALWAYS,
        'rels_from':   ['presentation', 'slideLayout'],
        'reltype':     RT_SLIDE_MASTER
    },
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
        'reltype':     RT_SLIDESHOW
    },
    CT_STYLES: {  # ISO/IEC 29500-1 11.3.12
        'basename':    'styles',
        'ext':         '.xml',
        'name':        'Style Definitions Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['document'],
        'reltype':     RT_STYLES
    },
    CT_TABLE_STYLES: {  # ECMA-376-1 14.2.9
        'basename':    'tableStyles',
        'ext':         '.xml',
        'name':        'Table Styles Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_TABLESTYLES
    },
    CT_TAGS: {  # ECMA-376-1 13.3.12
        'basename':    'tag',
        'ext':         '.xml',
        'name':        'User-Defined Tags Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/tags',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation', 'slide'],
        'reltype':     RT_TAGS
    },
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
        'reltype':     RT_TEMPLATE
    },
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
        'reltype':     RT_THEME
    },
    CT_VIEW_PROPS: {  # ECMA-376-1 13.3.13
        'basename':    'viewProps',
        'ext':         '.xml',
        'name':        'View Properties Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/ppt',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['presentation'],
        'reltype':     RT_VIEWPROPS
    },
    CT_WEB_SETTINGS: {  # ISO/IEC 29500-1 11.3.13
        'basename':    'webSettings',
        'ext':         '.xml',
        'name':        'Web Settings Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document', 'glossary'],
        'reltype':     RT_WEB_SETTINGS
    },
    CT_WML_COMMENTS: {  # ISO/IEC 29500-1 11.3.2
        'basename':    'comments',
        'ext':         '.xml',
        'name':        'Comments Part',
        'cardinality': PTS_CARDINALITY_SINGLETON,
        'required':    False,
        'baseURI':     '/word',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['document'],
        'reltype':     RT_WML_COMMENTS
    },
    CT_WORKSHEET: {  # Embedded worksheet package; ISO/IEC 29500-1 15.2.11
        'basename':    'embeddings',
        'ext':         '.xlsx',
        'name':        'Embedded Excel Spreadsheet',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/embeddings',
        'has_rels':    PTS_HASRELS_OPTIONAL,
        'rels_from':   ['chart', 'handoutMaster', 'notesSlide',
                        'notesMaster', 'slide', 'slideLayout',
                        'slideMaster'],
        'reltype':     RT_PACKAGE
    },
    'image/bmp': {  # ECMA-376-1 15.2.14
        'basename':    'image',
        'ext':         '.bmp',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster', 'slide',
                        'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE
    },
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
        'reltype':     RT_IMAGE
    },
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
        'reltype':     RT_IMAGE
    },
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
        'reltype':     RT_IMAGE
    },
    'image/tiff': {
        'basename':    'image',
        'ext':         '.tiff',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   [],
        'reltype':     RT_IMAGE
    },
    'image/vnd.ms-photo': {
        'basename':    'hdphoto',
        'ext':         '.wdp',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   [],
        'reltype':     RT_IMAGE
    },
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
        'reltype':     RT_IMAGE
    },
    'image/x-wmf': {
        'basename':    'image',
        'ext':         '.wmf',
        'name':        'Image Part',
        'cardinality': PTS_CARDINALITY_TUPLE,
        'required':    False,
        'baseURI':     '/ppt/media',
        'has_rels':    PTS_HASRELS_NEVER,
        'rels_from':   ['handoutMaster', 'notesSlide', 'notesMaster',
                        'slide', 'slideLayout', 'slideMaster'],
        'reltype':     RT_IMAGE,
    },
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
    '.bin':     CT_PRINTER_SETTINGS,
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
    '.xlsx':    CT_WORKSHEET,
    '.xml':     'application/xml'
}


# ============================================================================
# nsmap
# ============================================================================
# namespace prefix to namespace name map
# ============================================================================

nsmap = {
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c':   'http://schemas.openxmlformats.org/drawingml/2006/chart',
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
