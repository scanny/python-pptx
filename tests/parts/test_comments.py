# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.comments` module."""

from __future__ import annotations

from unittest.mock import Mock

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.oxml.comments import CT_CommentAuthorList, CT_CommentList
from pptx.parts.comments import CommentAuthorsPart, CommentsPart


class DescribeCommentAuthorsPart:
    def it_constructs_a_new_empty_part(self):
        package = Mock(name="package")

        part = CommentAuthorsPart.new(package)

        assert isinstance(part, CommentAuthorsPart)
        assert part.partname == "/ppt/commentAuthors.xml"
        assert part.content_type == CT.PML_COMMENT_AUTHORS
        assert isinstance(part._element, CT_CommentAuthorList)
        assert len(part._element.cmAuthor_lst) == 0

    def it_exposes_the_author_list(self):
        package = Mock(name="package")
        part = CommentAuthorsPart.new(package)

        assert part.author_list is part._element


class DescribeCommentsPart:
    def it_constructs_a_new_empty_part_at_supplied_partname(self):
        package = Mock(name="package")
        partname = PackURI("/ppt/comments/comment1.xml")

        part = CommentsPart.new(package, partname)

        assert isinstance(part, CommentsPart)
        assert part.partname == partname
        assert part.content_type == CT.PML_COMMENTS
        assert isinstance(part._element, CT_CommentList)
        assert len(part._element.cm_lst) == 0

    def it_exposes_the_comment_list(self):
        package = Mock(name="package")
        partname = PackURI("/ppt/comments/comment1.xml")
        part = CommentsPart.new(package, partname)

        assert part.comment_list is part._element
