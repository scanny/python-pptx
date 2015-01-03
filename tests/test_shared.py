# encoding: utf-8

"""
Test suite for the docx.shared module
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import pytest

from pptx.shared import ElementProxy


class DescribeElementProxy(object):

    def it_raises_on_assign_to_undefined_attr(self):
        element_proxy = ElementProxy(None)
        with pytest.raises(AttributeError):
            element_proxy.foobar = 42
