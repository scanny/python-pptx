# encoding: utf-8

"""
Test suite for pptx.package module
"""

from __future__ import absolute_import, print_function

import pytest

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, _Relationship
from pptx.opc.packuri import PackURI
from pptx.package import _ImageParts, Package
from pptx.parts.coreprops import CoreProperties
from pptx.parts.image import Image, ImagePart
from pptx.parts.presentation import PresentationPart


from .unitutil.file import absjoin
from .unitutil.mock import (
    class_mock, instance_mock, method_mock, property_mock
)


class DescribePackage(object):

    def it_loads_default_template_when_opened_with_no_path(self):
        prs = Package.open().presentation
        assert prs is not None
        slide_masters = prs.slide_masters
        assert slide_masters is not None
        assert len(slide_masters) == 1
        slide_layouts = slide_masters[0].slide_layouts
        assert slide_layouts is not None
        assert len(slide_layouts) == 11

    def it_provides_ref_to_package_presentation_part(self):
        pkg = Package.open()
        assert isinstance(pkg.presentation, PresentationPart)

    def it_provides_access_to_its_core_properties_part(self):
        pkg = Package.open()
        assert isinstance(pkg.core_properties, CoreProperties)

    def it_can_get_or_add_an_image_part(self, image_part_fixture):
        package, image_file, image_part_ = image_part_fixture

        image_part = package.get_or_add_image_part(image_file)

        package._image_parts.get_or_add_image_part.assert_called_once_with(
            image_file
        )
        assert image_part is image_part_

    def it_can_save_itself_to_a_pptx_file(self, temp_pptx_path):
        """
        Package.save produces a .pptx with plausible contents
        """
        # setup ------------------------
        pkg = Package.open()
        # exercise ---------------------
        pkg.save(temp_pptx_path)
        # verify -----------------------
        pkg = Package.open(temp_pptx_path)
        prs = pkg.presentation
        assert prs is not None
        slide_masters = prs.slide_masters
        assert slide_masters is not None
        assert len(slide_masters) == 1
        slide_layouts = slide_masters[0].slide_layouts
        assert slide_layouts is not None
        assert len(slide_layouts) == 11

    def it_knows_the_next_available_image_partname(self, next_fixture):
        package, ext, expected_value = next_fixture
        partname = package.next_image_partname(ext)
        assert partname == expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture
    def image_part_fixture(self, _image_parts_, image_part_):
        package = Package()
        image_file = 'foobar.png'
        package._image_parts.get_or_add_image_part.return_value = image_part_
        return package, image_file, image_part_

    @pytest.fixture(params=[
        ((3, 4, 2), 1),
        ((4, 2, 1), 3),
        ((2, 3, 1), 4),
    ])
    def next_fixture(self, request, iter_parts_):
        idxs, idx = request.param
        package = Package()
        package.iter_parts.return_value = self.i_image_parts(request, idxs)
        ext = 'foo'
        expected_value = '/ppt/media/image%d.%s' % (idx, ext)
        return package, ext, expected_value

    @pytest.fixture
    def temp_pptx_path(self, tmpdir):
        return absjoin(str(tmpdir), 'test-pptx.pptx')

    # fixture components ---------------------------------------------

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def _image_parts_(self, request):
        return property_mock(request, Package, '_image_parts')

    def i_image_parts(self, request, idxs):
        def part(idx):
            partname = PackURI('/ppt/media/image%d.png' % idx)
            return instance_mock(request, Part, partname=partname)
        return iter([part(idx) for idx in idxs])

    @pytest.fixture
    def iter_parts_(self, request):
        return property_mock(request, Package, 'iter_parts')


class Describe_ImageParts(object):

    def it_can_iterate_over_the_package_image_parts(self, iter_fixture):
        image_parts, expected_parts = iter_fixture
        assert list(image_parts) == expected_parts

    def it_can_get_a_matching_image_part(self, get_fixture):
        image_parts, image_file, Image_, image_, image_part_ = get_fixture

        image_part = image_parts.get_or_add_image_part(image_file)

        Image_.from_file.assert_called_once_with(image_file)
        image_parts._find_by_sha1.assert_called_once_with(image_.sha1)
        assert image_part is image_part_

    def it_can_add_an_image_part(self, add_fixture):
        image_parts, image_file, Image_, image_ = add_fixture[:4]
        ImagePart_, package_, image_part_ = add_fixture[4:]

        image_part = image_parts.get_or_add_image_part(image_file)

        Image_.from_file.assert_called_once_with(image_file)
        image_parts._find_by_sha1.assert_called_once_with(image_.sha1)
        ImagePart_.new.assert_called_once_with(package_, image_)
        assert image_part is image_part_

    def it_can_find_an_image_part_by_sha1_hash(self, find_fixture):
        image_parts, sha1, expected_value = find_fixture
        image_part = image_parts._find_by_sha1(sha1)
        assert image_part is expected_value

    # fixtures ---------------------------------------------

    @pytest.fixture
    def add_fixture(self, package_, Image_, image_, _find_by_sha1_,
                    ImagePart_, image_part_):
        image_parts = _ImageParts(package_)
        image_file = 'foobar.png'
        Image_.from_file.return_value = image_
        _find_by_sha1_.return_value = None
        ImagePart_.new.return_value = image_part_
        return (
            image_parts, image_file, Image_, image_, ImagePart_, package_,
            image_part_
        )

    @pytest.fixture(params=[True, False])
    def find_fixture(self, request, _iter_, image_part_):
        image_part_is_present = request.param
        image_parts = _ImageParts(None)
        _iter_.return_value = iter((image_part_,))
        sha1 = 'foobar'
        if image_part_is_present:
            image_part_.sha1 = 'foobar'
            expected_value = image_part_
        else:
            image_part_.sha1 = 'barfoo'
            expected_value = None
        return image_parts, sha1, expected_value

    @pytest.fixture
    def get_fixture(self, Image_, image_, image_part_, _find_by_sha1_):
        image_parts = _ImageParts(None)
        image_file = 'foobar.png'
        Image_.from_file.return_value = image_
        _find_by_sha1_.return_value = image_part_
        return image_parts, image_file, Image_, image_, image_part_

    @pytest.fixture
    def iter_fixture(self, request, package_):

        def rel(is_external, reltype):
            part = instance_mock(request, Part)
            return instance_mock(
                request, _Relationship, is_external=is_external,
                reltype=reltype, target_part=part
            )

        rels = (
            rel(True,  RT.IMAGE),
            rel(False, RT.SLIDE),
            rel(False, RT.IMAGE),
        )

        package_.iter_rels.return_value = iter(
            (rels[0], rels[1], rels[2], rels[2])
        )
        image_parts = _ImageParts(package_)
        expected_parts = [rels[2].target_part]
        return image_parts, expected_parts

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _find_by_sha1_(self, request):
        return method_mock(request, _ImageParts, '_find_by_sha1')

    @pytest.fixture
    def Image_(self, request):
        return class_mock(request, 'pptx.package.Image')

    @pytest.fixture
    def image_(self, request):
        return instance_mock(request, Image)

    @pytest.fixture
    def ImagePart_(self, request):
        return class_mock(request, 'pptx.package.ImagePart')

    @pytest.fixture
    def image_part_(self, request):
        return instance_mock(request, ImagePart)

    @pytest.fixture
    def _iter_(self, request):
        return method_mock(request, _ImageParts, '__iter__')

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)
