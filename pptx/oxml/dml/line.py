# encoding: utf-8

"""
lxml custom element classes for DrawingML line-related XML elements.
"""

from __future__ import absolute_import


class EG_LineDashProperties(object):

    __member_names__ = ('a:prstDash', 'a:custDash')


class EG_LineFillProperties(object):

    __member_names__ = (
        'a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill'
    )


class EG_LineJoinProperties(object):

    __member_names__ = ('a:round', 'a:bevel', 'a:miter')
