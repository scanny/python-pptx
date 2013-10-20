# encoding: utf-8

"""
Part objects, including _BasePart.
"""

import pptx.util as util

from pptx.util import Collection


class _Observable(object):
    """
    Simple observer pattern mixin. Limitations:

    * observers get all message types from subject (_Observable), subscription
      is on subject basis, not subject + event_type.

    * notifications are oriented toward "value has been updated", which seems
      like only one possible event, could also be something like "load has
      completed" or "preparing to load".
    """
    def __init__(self):
        super(_Observable, self).__init__()
        self._observers = []

    def _notify_observers(self, name, value):
        # value = getattr(self, name)
        for observer in self._observers:
            observer.notify(self, name, value)

    def add_observer(self, observer):
        """
        Begin notifying *observer* of events. *observer* must implement method
        ``notify(observed, name, new_value)``
        """
        if observer not in self._observers:
            self._observers.append(observer)


class _PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(_PartCollection, self).__init__()

    def _loadpart(self, part):
        """
        Insert a new part loaded from a package, such that list remains
        sorted in logical partname order (e.g. slide10.xml comes after
        slide9.xml).
        """
        new_partidx = util.Partname(part.partname).idx
        for idx, seq_part in enumerate(self._values):
            partidx = util.Partname(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)
