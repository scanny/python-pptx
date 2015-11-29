# encoding: utf-8

"""
Title-related oxml objects and mixins.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..xmlchemy import BaseOxmlElement, ZeroOrOne


class TitleMixin(object):
    """
    Mixin class for chart titles.
    """
    @property
    def has_title(self):
        """
        True if this element has a title defined, False otherwise.
        """
        return self.title is not None

    @has_title.setter
    def has_title(self, bool_value):
        """
        Add, remove, or leave alone the ``<c:title>`` child element depending
        on current state and *bool_value*. If *bool_value* is |True| and no
        ``<c:title>`` element is present, a new default element is added.
        When |False|, any existing title element is removed.
        """
        if bool(bool_value) is False:
            self._remove_title()
        else:
            if self.title is None:
                self._add_title()


class CT_Title(BaseOxmlElement):
    """
    ``<c:title>`` element.
    """
    _tag_seq = (
        'c:tx', 'c:layout', 'c:overlay'
    )
    tx = ZeroOrOne('c:tx', successors=_tag_seq[1:])
    layout = ZeroOrOne('c:layout', successors=_tag_seq[2:])
    overlay = ZeroOrOne('c:overlay', successors=_tag_seq[3:])
    del _tag_seq

    @property
    def text_frame(self):
        return self.xpath('./c:tx/c:rich')[0]
