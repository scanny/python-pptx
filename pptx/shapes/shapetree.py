# encoding: utf-8

"""
The shape tree, the structure that holds a slide's shapes.
"""

from .autoshape import Shape
from ..oxml.ns import qn
from .picture import Picture
from .shape import BaseShape
from .table import Table


class BaseShapeTree(object):
    """
    Base class for a shape collection appearing in a slide-type object,
    include Slide, SlideLayout, and SlideMaster, providing common methods.
    """
    def __init__(self, slide):
        super(BaseShapeTree, self).__init__()
        self._slide = slide

    def __getitem__(self, idx):
        """
        Return shape at *idx* in sequence, e.g. ``shapes[2]``.
        """
        shape_elms = list(self._iter_member_elms())
        try:
            shape_elm = shape_elms[idx]
        except IndexError:
            raise IndexError('shape index out of range')
        return self._shape_factory(shape_elm)

    def __iter__(self):
        """
        Generate a reference to each shape in the collection, in sequence.
        """
        for shape_elm in self._iter_member_elms():
            yield self._shape_factory(shape_elm)

    def __len__(self):
        """
        Return count of shapes in this shape tree. A group shape contributes
        1 to the total, without regard to the number of shapes contained in
        the group.
        """
        shape_elms = list(self._iter_member_elms())
        return len(shape_elms)

    @property
    def part(self):
        """
        The package part containing this object, a _BaseSlide subclass in
        this case.
        """
        return self._slide

    @staticmethod
    def _is_member_elm(shape_elm):
        """
        Return true if *shape_elm* represents a member of this collection,
        False otherwise.
        """
        return True

    def _iter_member_elms(self):
        """
        Generate each child of the ``<p:spTree>`` element that corresponds to
        a shape, in the sequence they appear in the XML.
        """
        spTree = self._slide.spTree
        for shape_elm in spTree.iter_shape_elms():
            if self._is_member_elm(shape_elm):
                yield shape_elm

    @property
    def _next_shape_id(self):
        """
        Next available positive integer drawing object id in shape tree,
        starting from 1 and making use of any gaps in numbering. In practice,
        the minimum id is 2 because the spTree element is always assigned
        id="1".
        """
        id_str_lst = self._spTree.xpath('//@id')
        used_ids = [int(id_str) for id_str in id_str_lst if id_str.isdigit()]
        for n in range(1, len(used_ids)+2):
            if n not in used_ids:
                return n

    def _shape_factory(self, shape_elm):
        """
        Return an instance of the appropriate shape proxy class for
        *shape_elm*.
        """
        return BaseShapeFactory(shape_elm, self)

    @property
    def _spTree(self):
        """
        The ``<p:spTree>`` element underlying this shape tree object
        """
        return self._slide.spTree


def BaseShapeFactory(shape_elm, parent):
    """
    Return an instance of the appropriate shape proxy class for *shape_elm*.
    """
    tag_name = shape_elm.tag
    if tag_name == qn('p:sp'):
        return Shape(shape_elm, parent)
    if tag_name == qn('p:pic'):
        return Picture(shape_elm, parent)
    if tag_name == qn('p:graphicFrame'):
        if shape_elm.has_table:
            return Table(shape_elm, parent)
    return BaseShape(shape_elm, parent)
