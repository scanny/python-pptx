# encoding: utf-8

"""Unit-test suite for the `pptx.opc.packuri` module."""

import pytest

from pptx.opc.packuri import PackURI


class DescribePackURI(object):
    """Unit-test suite for the `pptx.opc.packuri.PackURI` objects."""

    def it_can_construct_from_relative_ref(self):
        pack_uri = PackURI.from_rel_ref(
            "/ppt/slides", "../slideLayouts/slideLayout1.xml"
        )
        assert pack_uri == "/ppt/slideLayouts/slideLayout1.xml"

    def it_should_raise_on_construct_with_bad_pack_uri_str(self):
        with pytest.raises(ValueError):
            PackURI("foobar")

    @pytest.mark.parametrize(
        "uri, expected_value",
        (
            ("/", "/"),
            ("/ppt/presentation.xml", "/ppt"),
            ("/ppt/slides/slide1.xml", "/ppt/slides"),
        ),
    )
    def it_knows_its_base_URI(self, uri, expected_value):
        assert PackURI(uri).baseURI == expected_value

    @pytest.mark.parametrize(
        "uri, expected_value",
        (
            ("/", ""),
            ("/ppt/presentation.xml", "xml"),
            ("/ppt/media/image.PnG", "PnG"),
        ),
    )
    def it_knows_its_extension(self, uri, expected_value):
        assert PackURI(uri).ext == expected_value

    @pytest.mark.parametrize(
        "uri, expected_value",
        (
            ("/", ""),
            ("/ppt/presentation.xml", "presentation.xml"),
            ("/ppt/media/image.png", "image.png"),
        ),
    )
    def it_knows_its_filename(self, uri, expected_value):
        assert PackURI(uri).filename == expected_value

    @pytest.mark.parametrize(
        "uri, expected_value",
        (
            ("/", None),
            ("/ppt/presentation.xml", None),
            ("/ppt/,foo,grob!.xml", None),
            ("/ppt/media/image42.png", 42),
        ),
    )
    def it_knows_the_filename_index(self, uri, expected_value):
        assert PackURI(uri).idx == expected_value

    @pytest.mark.parametrize(
        "uri, base_uri, expected_value",
        (
            ("/ppt/presentation.xml", "/", "ppt/presentation.xml"),
            (
                "/ppt/slideMasters/slideMaster1.xml",
                "/ppt",
                "slideMasters/slideMaster1.xml",
            ),
            (
                "/ppt/slideLayouts/slideLayout1.xml",
                "/ppt/slides",
                "../slideLayouts/slideLayout1.xml",
            ),
        ),
    )
    def it_can_compute_its_relative_reference(self, uri, base_uri, expected_value):
        assert PackURI(uri).relative_ref(base_uri) == expected_value

    @pytest.mark.parametrize(
        "uri, expected_value",
        (
            ("/", "/_rels/.rels"),
            ("/ppt/presentation.xml", "/ppt/_rels/presentation.xml.rels"),
            ("/ppt/slides/slide42.xml", "/ppt/slides/_rels/slide42.xml.rels"),
        ),
    )
    def it_knows_the_uri_of_its_rels_part(self, uri, expected_value):
        assert PackURI(uri).rels_uri == expected_value
