"""Unit-test suite for `pptx.parts.font` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import OpcPackage, PackURI
from pptx.parts.font import FontPart

from ..unitutil.mock import ANY, FixtureRequest, initializer_mock, instance_mock


class DescribeFontPart(object):
    """Unit-test suite for `pptx.parts.font.FontPart`."""

    def it_provides_a_new_classmethod_that_builds_a_part_from_a_blob(self, request: FixtureRequest):
        font_blob = b"fake-ttf-bytes"
        package_ = instance_mock(request, OpcPackage)
        partname_ = instance_mock(request, PackURI)
        package_.next_partname.return_value = partname_
        _init_ = initializer_mock(request, FontPart, autospec=True)

        font_part = FontPart.new(font_blob, package_)

        package_.next_partname.assert_called_once_with("/ppt/fonts/font%d.fntdata")
        _init_.assert_called_once_with(ANY, partname_, CT.X_FONTDATA, package_, font_blob)
        assert isinstance(font_part, FontPart)

    def it_exposes_its_content_type_and_partname_template(self):
        assert FontPart.content_type == CT.X_FONTDATA
        assert FontPart.partname_template == "/ppt/fonts/font%d.fntdata"
