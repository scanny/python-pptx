# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.tags` module."""

from __future__ import annotations

from unittest.mock import Mock

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.oxml.tags import CT_TagList
from pptx.parts.tags import TagsPart


class DescribeTagsPart:
    def it_constructs_a_new_empty_part_at_supplied_partname(self):
        package = Mock(name="package")
        partname = PackURI("/ppt/tags/tag1.xml")

        part = TagsPart.new(package, partname)

        assert isinstance(part, TagsPart)
        assert part.partname == partname
        assert part.content_type == CT.PML_TAGS
        assert isinstance(part._element, CT_TagList)
        assert len(part._element.tag_lst) == 0

    def it_exposes_the_tag_list(self):
        package = Mock(name="package")
        partname = PackURI("/ppt/tags/tag1.xml")
        part = TagsPart.new(package, partname)

        assert part.tag_list is part._element
