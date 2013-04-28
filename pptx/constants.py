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


class MSO(object):
    """
    Constants corresponding to things like ``msoAnchorMiddle`` in the MS
    Office API.
    """
    # _TextFrame.vertical_anchor values
    ANCHOR_TOP = 1
    ANCHOR_MIDDLE = 3
    ANCHOR_BOTTOM = 4
