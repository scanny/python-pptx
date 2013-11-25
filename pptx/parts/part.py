# encoding: utf-8

"""
General-purpose Part-related objects
"""

from pptx.opc.packuri import PackURI
from pptx.util import Collection


class PartCollection(Collection):
    """
    Sequence of parts. Sensitive to partname index when ordering parts added
    via _loadpart(), e.g. ``/ppt/slide/slide2.xml`` appears before
    ``/ppt/slide/slide10.xml`` rather than after it as it does in a
    lexicographical sort.
    """
    def __init__(self):
        super(PartCollection, self).__init__()

    def add_part(self, part):
        """
        Insert a new part into the collection such that list remains sorted
        in logical partname order (e.g. slide10.xml comes after slide9.xml).
        """
        new_partidx = part.partname.idx
        for idx, seq_part in enumerate(self._values):
            partidx = PackURI(seq_part.partname).idx
            if partidx > new_partidx:
                self._values.insert(idx, part)
                return
        self._values.append(part)
