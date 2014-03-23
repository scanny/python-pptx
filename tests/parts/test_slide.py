# encoding: utf-8

"""
Test suite for pptx.parts.slide module
"""

from __future__ import absolute_import

import pytest

from lxml import objectify
from mock import ANY, call, MagicMock

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.opc.package import Part, _Relationship
from pptx.oxml.ns import _nsmap
from pptx.oxml.presentation import CT_SlideId, CT_SlideIdList
from pptx.oxml.shapetree import CT_GroupShape
from pptx.oxml.slide import CT_Slide
from pptx.package import Package
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import (
    BaseSlide, Slide, SlideCollection, _SlideShapeTree
)
from pptx.parts.slidelayout import SlideLayout
from pptx.shapes.shapetree import ShapeCollection

from ..oxml.unitdata.shape import an_spTree
from ..oxml.unitdata.slides import a_sld, a_cSld
from ..unitutil import (
    absjoin, class_mock, initializer_mock, instance_mock, loose_mock,
    method_mock, parse_xml_file, property_mock, serialize_xml, test_file_dir
)


def actual_xml(elm):
    objectify.deannotate(elm, cleanup_namespaces=True)
    return serialize_xml(elm, pretty_print=True)


def _sldLayout1():
    path = absjoin(test_file_dir, 'slideLayout1.xml')
    sldLayout = parse_xml_file(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=_nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


class DescribeBaseSlide(object):

    def it_provides_access_to_its_spTree_element_to_help_ShapeTree(
            self, slide):
        spTree = slide.spTree
        assert isinstance(spTree, CT_GroupShape)

    def it_knows_the_name_of_the_slide(self, base_slide):
        # setup ------------------------
        base_slide._element = _sldLayout1()
        # exercise ---------------------
        name = base_slide.name
        # verify -----------------------
        expected = 'Title Slide'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        assert actual == expected, msg

    def it_can_add_an_image_part_to_the_slide(self, base_slide_fixture):
        # fixture ----------------------
        base_slide, img_file_, image_, rId_ = base_slide_fixture
        # exercise ---------------------
        image, rId = base_slide._add_image(img_file_)
        # verify -----------------------
        base_slide._package._images.add_image.assert_called_once_with(
            img_file_)
        base_slide.relate_to.assert_called_once_with(image, RT.IMAGE)
        assert image is image_
        assert rId is rId_

    def it_knows_it_is_the_part_its_child_objects_belong_to(
            self, base_slide):
        assert base_slide.part is base_slide

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def base_slide_fixture(self, request, base_slide):
        # mock BaseSlide._package._images.add_image() train wreck
        img_file_ = loose_mock(request, name='img_file_')
        image_ = loose_mock(request, name='image_')
        pkg_ = loose_mock(request, name='_package', spec=Package)
        pkg_._images.add_image.return_value = image_
        base_slide._package = pkg_
        # mock BaseSlide.relate_to()
        rId_ = loose_mock(request, name='rId_')
        method_mock(request, BaseSlide, 'relate_to', return_value=rId_)
        return base_slide, img_file_, image_, rId_

    # fixture components ---------------------------------------------

    @pytest.fixture
    def base_slide(self):
        partname = PackURI('/foo/bar.xml')
        return BaseSlide(partname, None, None, None)

    @pytest.fixture
    def sld(self):
        sld_bldr = (
            a_sld().with_nsdecls().with_child(
                a_cSld().with_child(
                    an_spTree()))
        )
        return sld_bldr.element

    @pytest.fixture
    def slide(self, sld):
        return BaseSlide(None, None, sld, None)


class DescribeSlide(object):

    def it_provides_access_to_the_shapes_on_the_slide(self, shapes_fixture):
        slide, _SlideShapeTree_, slide_shape_tree_ = shapes_fixture
        shapes = slide.shapes_new
        _SlideShapeTree_.assert_called_once_with(slide)
        assert shapes is slide_shape_tree_

    def it_can_create_a_new_slide(self, new_fixture):
        slide_layout_, partname_, package_ = new_fixture[:3]
        Slide_init_, slide_elm_, shapes_, relate_to_ = new_fixture[3:]
        slide = Slide.new(slide_layout_, partname_, package_)
        Slide_init_.assert_called_once_with(
            partname_, CT.PML_SLIDE, slide_elm_, package_
        )
        shapes_._clone_layout_placeholders.assert_called_once_with(
            slide_layout_
        )
        relate_to_.assert_called_once_with(
            slide_layout_, RT.SLIDE_LAYOUT
        )
        assert isinstance(slide, Slide)

    def it_knows_the_slide_layout_it_inherits_from(self, layout_fixture):
        slide, slide_layout_ = layout_fixture
        slide_layout = slide.slide_layout
        slide.part_related_by.assert_called_once_with(RT.SLIDE_LAYOUT)
        assert slide_layout is slide_layout_

    def it_knows_the_minimal_element_xml_for_a_slide(self, slide):
        path = absjoin(test_file_dir, 'minimal_slide.xml')
        sld = slide._minimal_element()
        with open(path, 'r') as f:
            expected_xml = f.read()
        assert actual_xml(sld) == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def layout_fixture(self, slide_layout_, part_related_by_):
        slide = Slide(None, None, None, None)
        return slide, slide_layout_

    @pytest.fixture
    def shapes_fixture(self, _SlideShapeTree_, slide_shape_tree_):
        slide = Slide(None, None, None, None)
        return slide, _SlideShapeTree_, slide_shape_tree_

    @pytest.fixture
    def new_fixture(
            self, slide_layout_, partname_, package_, Slide_init_,
            _minimal_element_, slide_elm_, shapes_prop_, shapes_,
            relate_to_):
        return (
            slide_layout_, partname_, package_, Slide_init_, slide_elm_,
            shapes_, relate_to_
        )

    # fixture components -----------------------------------

    @pytest.fixture
    def _minimal_element_(self, request, slide_elm_):
        return method_mock(
            request, Slide, '_minimal_element', return_value=slide_elm_
        )

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request, slide_layout_):
        return method_mock(
            request, Slide, 'part_related_by',
            return_value=slide_layout_
        )

    @pytest.fixture
    def partname_(self, request):
        return instance_mock(request, PackURI)

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, Part, 'relate_to')

    @pytest.fixture
    def shapes_(self, request):
        return instance_mock(request, _SlideShapeTree)

    @pytest.fixture
    def shapes_prop_(self, request, shapes_):
        return property_mock(request, Slide, 'shapes', return_value=shapes_)

    @pytest.fixture
    def slide(self):
        return Slide(None, None, None, None)

    @pytest.fixture
    def _SlideShapeTree_(self, request, slide_shape_tree_):
        return class_mock(
            request, 'pptx.parts.slide._SlideShapeTree',
            return_value=slide_shape_tree_
        )

    @pytest.fixture
    def slide_elm_(self, request):
        return instance_mock(request, CT_Slide)

    @pytest.fixture
    def Slide_init_(self, request):
        return initializer_mock(request, Slide)

    @pytest.fixture
    def slide_layout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slide_shape_tree_(self, request):
        return instance_mock(request, _SlideShapeTree)


class DescribeSlideCollection(object):

    def it_supports_indexed_access(self, slides_with_slide_parts_, rIds_):
        slides, slide_, slide_2_ = slides_with_slide_parts_
        rId_, rId_2_ = rIds_
        # verify -----------------------
        assert slides[0] is slide_
        assert slides[1] is slide_2_
        slides._sldIdLst.__getitem__.assert_has_calls(
            [call(0), call(1)]
        )
        slides._prs.related_parts.__getitem__.assert_has_calls(
            [call(rId_), call(rId_2_)]
        )

    def it_raises_on_slide_index_out_of_range(self, slides):
        with pytest.raises(IndexError):
            slides[2]

    def it_can_iterate_over_the_slides(self, slides, slide_, slide_2_):
        assert [s for s in slides] == [slide_, slide_2_]

    def it_supports_len(self, slides):
        assert len(slides) == 2

    def it_can_add_a_new_slide(self, slides, slidelayout_, Slide_, slide_):
        slide = slides.add_slide(slidelayout_)
        Slide_.new.assert_called_once_with(
            slidelayout_, PackURI('/ppt/slides/slide3.xml'),
            slides._prs.package
        )
        slides._prs.relate_to.assert_called_once_with(slide_, RT.SLIDE)
        slides._sldIdLst.add_sldId.assert_called_once_with(ANY)
        assert slide is slide_

    def it_knows_the_next_available_slide_partname(
            self, slides_with_slide_parts_):
        slides = slides_with_slide_parts_[0]
        expected_partname = PackURI('/ppt/slides/slide3.xml')
        partname = slides._next_partname
        assert isinstance(partname, PackURI)
        assert partname == expected_partname

    def it_can_assign_partnames_to_the_slides(
            self, slides, slide_, slide_2_):
        slides.rename_slides()
        assert slide_.partname == '/ppt/slides/slide1.xml'
        assert slide_2_.partname == '/ppt/slides/slide2.xml'

    # fixtures -------------------------------------------------------
    #
    #   slides
    #   |
    #   +- ._sldIdLst = [sldId_, sldId_2_]
    #   |                |       |
    #   |                |       +- .rId = rId_2_
    #   |                |
    #   |                +- .rId = rId_
    #   +- ._prs
    #       |
    #       +- .related_parts = {rId_: slide_, rId_2_: slide_2_}
    #
    # ----------------------------------------------------------------

    @pytest.fixture
    def prs_(self, request, rel_, related_parts_):
        prs_ = instance_mock(request, PresentationPart)
        prs_.load_rel.return_value = rel_
        prs_.related_parts = related_parts_
        return prs_

    @pytest.fixture
    def rel_(self, request, rId_):
        return instance_mock(request, _Relationship, rId=rId_)

    @pytest.fixture
    def related_parts_(self, request, rIds_, slide_parts_):
        """
        Return pass-thru mock dict that both operates as a dict an records
        calls to __getitem__ for call asserts.
        """
        rId_, rId_2_ = rIds_
        slide_, slide_2_ = slide_parts_
        slide_rId_map = {rId_: slide_, rId_2_: slide_2_}

        def getitem(key):
            return slide_rId_map[key]

        related_parts_ = MagicMock()
        related_parts_.__getitem__.side_effect = getitem
        return related_parts_

    @pytest.fixture
    def rename_slides_(self, request):
        return method_mock(request, SlideCollection, 'rename_slides')

    @pytest.fixture
    def rId_(self, request):
        return 'rId1'

    @pytest.fixture
    def rId_2_(self, request):
        return 'rId2'

    @pytest.fixture
    def rIds_(self, request, rId_, rId_2_):
        return rId_, rId_2_

    @pytest.fixture
    def Slide_(self, request, slide_):
        Slide_ = class_mock(request, 'pptx.parts.slide.Slide')
        Slide_.new.return_value = slide_
        return Slide_

    @pytest.fixture
    def sldId_(self, request, rId_):
        return instance_mock(request, CT_SlideId, rId=rId_)

    @pytest.fixture
    def sldId_2_(self, request, rId_2_):
        return instance_mock(request, CT_SlideId, rId=rId_2_)

    @pytest.fixture
    def sldIdLst_(self, request, sldId_, sldId_2_):
        sldIdLst_ = instance_mock(request, CT_SlideIdList)
        sldIdLst_.__getitem__.side_effect = [sldId_, sldId_2_]
        sldIdLst_.__iter__.return_value = iter([sldId_, sldId_2_])
        sldIdLst_.__len__.return_value = 2
        return sldIdLst_

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_2_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_parts_(self, request, slide_, slide_2_):
        return slide_, slide_2_

    @pytest.fixture
    def slidelayout_(self, request):
        return instance_mock(request, SlideLayout)

    @pytest.fixture
    def slides(self, sldIdLst_, prs_):
        return SlideCollection(sldIdLst_, prs_)

    @pytest.fixture
    def slides_with_slide_parts_(self, sldIdLst_, prs_, slide_parts_):
        slide_, slide_2_ = slide_parts_
        slides = SlideCollection(sldIdLst_, prs_)
        return slides, slide_, slide_2_
