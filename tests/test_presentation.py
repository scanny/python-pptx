# encoding: utf-8

"""Test suite for pptx.presentation module."""

import gc
import os

from datetime import datetime, timedelta
from StringIO import StringIO

from hamcrest import (
    assert_that, equal_to, instance_of, is_, is_in, is_not, less_than
)
from mock import Mock, patch, PropertyMock

import pptx.presentation

from pptx.exceptions import InvalidPackageError
from pptx.opc_constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.oxml import CT_CoreProperties, oxml_fromstring, oxml_parse
from pptx.presentation import (
    _BasePart, _BaseSlide, _CoreProperties, _Image, _Package, _Part,
    Presentation, _Relationship, _RelationshipCollection, _Slide,
    _SlideCollection, _SlideLayout, _SlideMaster
)
from pptx.shapes.shapetree import _ShapeCollection
from pptx.spec import namespaces, qtag
from pptx.util import Px
from testing import TestCase

import logging
log = logging.getLogger('pptx.test.presentation')
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - '
                              '%(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)


# module globals -------------------------------------------------------------
def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
test_file_dir = absjoin(thisdir, 'test_files')

test_image_path = absjoin(test_file_dir, 'python-icon.jpeg')
test_bmp_path = absjoin(test_file_dir, 'python.bmp')
new_image_path = absjoin(test_file_dir, 'monty-truth.png')
test_pptx_path = absjoin(test_file_dir, 'test.pptx')
images_pptx_path = absjoin(test_file_dir, 'with_images.pptx')

nsmap = namespaces('a', 'r', 'p')


def _sldLayout1():
    path = os.path.join(thisdir, 'test_files/slideLayout1.xml')
    sldLayout = oxml_parse(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = _ShapeCollection(spTree)
    return shapes


class PartBuilder(object):
    """Builder class for test Parts"""
    def __init__(self):
        self.partname = '/ppt/slides/slide1.xml'

    def with_partname(self, partname):
        self.partname = partname
        return self

    def build(self):
        p = _BasePart()
        p.partname = self.partname
        return p


class RelationshipCollectionBuilder(object):
    """Builder class for test RelationshipCollections"""
    partname_tmpls = {RT.SLIDE_MASTER: '/ppt/slideMasters/slideMaster%d.xml',
                      RT.SLIDE: '/ppt/slides/slide%d.xml'}

    def __init__(self):
        self.relationships = []
        self.next_rel_num = 1
        self.next_partnums = {}
        self.reltype_ordering = None

    def with_ordering(self, *reltypes):
        self.reltype_ordering = tuple(reltypes)
        return self

    def with_tuple_targets(self, count, reltype):
        for i in range(count):
            rId = self.__next_rId
            partname = self.__next_tuple_partname(reltype)
            target = PartBuilder().with_partname(partname).build()
            rel = _Relationship(rId, reltype, target)
            self.relationships.append(rel)
        return self

    # def with_singleton_target(self, reltype):
    #     rId = self.__next_rId
    #     partname = self.__singleton_partname(reltype)
    #     target = PartBuilder().with_partname(partname).build()
    #     rel = _Relationship(rId, reltype, target)
    #     self.relationships.append(rel)
    #     return self
    #
    def __next_partnum(self, reltype):
        if reltype not in self.next_partnums:
            self.next_partnums[reltype] = 1
        partnum = self.next_partnums[reltype]
        self.next_partnums[reltype] = partnum + 1
        return partnum

    @property
    def __next_rId(self):
        rId = 'rId%d' % self.next_rel_num
        self.next_rel_num += 1
        return rId

    def __next_tuple_partname(self, reltype):
        partname_tmpl = self.partname_tmpls[reltype]
        partnum = self.__next_partnum(reltype)
        return partname_tmpl % partnum

    def build(self):
        rels = _RelationshipCollection()
        for rel in self.relationships:
            rels._additem(rel)
        if self.reltype_ordering:
            rels._reltype_ordering = self.reltype_ordering
        return rels


class Test_BaseSlide(TestCase):
    """Test _BaseSlide"""
    def setUp(self):
        self.base_slide = _BaseSlide()

    def test_name_value(self):
        """_BaseSlide.name value is correct"""
        # setup ------------------------
        self.base_slide._element = _sldLayout1()
        # exercise ---------------------
        name = self.base_slide.name
        # verify -----------------------
        expected = 'Title Slide'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_shapes_size_after__load(self):
        """_BaseSlide.shapes is expected size after _load()"""
        # setup ------------------------
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/slides/slide1.xml'
        with open(path, 'r') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = []
        part_dict = {}
        self.base_slide._load(pkgpart, part_dict)
        # exercise ---------------------
        shapes = self.base_slide.shapes
        # verify -----------------------
        self.assertLength(shapes, 9)

    @patch('pptx.presentation._BaseSlide._package', new_callable=PropertyMock)
    def test__add_image_collaboration(self, _package):
        """_BaseSlide._add_image() returns (image, rel) tuple"""
        # setup ------------------------
        base_slide = self.base_slide
        image = Mock(name='image')
        rel = Mock(name='rel')
        base_slide._package._images.add_image.return_value = image
        base_slide._add_relationship = Mock('_add_relationship')
        base_slide._add_relationship.return_value = rel
        file = test_image_path
        # exercise ---------------------
        retval_image, retval_rel = base_slide._add_image(file)
        # verify -----------------------
        base_slide._package._images.add_image.assert_called_once_with(file)
        base_slide._add_relationship.assert_called_once_with(RT.IMAGE, image)
        assert_that(retval_image, is_(image))
        assert_that(retval_rel, is_(rel))


class Test_CoreProperties(TestCase):
    """Test _CoreProperties"""
    def test_default_constructs_default_core_props(self):
        """_CoreProperties.default() returns new default core props part"""
        # exercise ---------------------
        core_props = _CoreProperties._default()
        # verify -----------------------
        assert_that(core_props, is_(instance_of(_CoreProperties)))
        assert_that(core_props._content_type, is_(CT.OPC_CORE_PROPERTIES))
        assert_that(core_props.partname, is_('/docProps/core.xml'))
        assert_that(core_props._element, is_(instance_of(CT_CoreProperties)))
        assert_that(core_props.title, is_('PowerPoint Presentation'))
        assert_that(core_props.last_modified_by, is_('python-pptx'))
        assert_that(core_props.revision, is_(1))
        # core_props.modified only stores time with seconds resolution, so
        # comparison needs to be a little loose (within two seconds)
        modified_timedelta = datetime.utcnow() - core_props.modified
        max_expected_timedelta = timedelta(seconds=2)
        assert_that(modified_timedelta, less_than(max_expected_timedelta))


class Test_Image(TestCase):
    """Test _Image"""
    def test_construction_from_file(self):
        """_Image(path) constructor produces correct attribute values"""
        # exercise ---------------------
        image = _Image(test_image_path)
        # verify -----------------------
        assert_that(image.ext, is_(equal_to('.jpeg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))
        assert_that(image._desc, is_(equal_to('python-icon.jpeg')))

    def test_construction_from_stream(self):
        """_Image(stream) construction produces correct attribute values"""
        # exercise ---------------------
        with open(test_image_path, 'rb') as f:
            stream = StringIO(f.read())
        image = _Image(stream)
        # verify -----------------------
        assert_that(image.ext, is_(equal_to('.jpg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))
        assert_that(image._desc, is_(equal_to('image.jpg')))

    def test_construction_from_file_raises_on_bad_path(self):
        """_Image(path) constructor raises on bad path"""
        # verify -----------------------
        with self.assertRaises(IOError):
            _Image('foobar27.png')

    def test__scale_calculates_correct_dimensions(self):
        """_Image._scale() calculates correct dimensions"""
        # setup ------------------------
        test_cases = (
            ((None, None), (Px(204), Px(204))),
            ((1000, None), (1000, 1000)),
            ((None, 3000), (3000, 3000)),
            ((3337, 9999), (3337, 9999)))
        image = _Image(test_image_path)
        # verify -----------------------
        for params, expected in test_cases:
            width, height = params
            assert_that(image._scale(width, height), is_(equal_to(expected)))

    def test__size_returns_image_native_pixel_dimensions(self):
        """_Image._size is width, height tuple of image pixel dimensions"""
        image = _Image(test_image_path)
        assert_that(image._size, is_(equal_to((204, 204))))

    def test___ext_from_image_stream_raises_on_incompatible_format(self):
        """_Image.__ext_from_image_stream() raises on incompatible format"""
        # verify -----------------------
        with self.assertRaises(ValueError):
            with open(test_bmp_path) as stream:
                _Image._Image__ext_from_image_stream(stream)

    def test___image_ext_content_type_known_type(self):
        """_Image.__image_ext_content_type() correct for known content type"""
        # exercise ---------------------
        content_type = _Image._Image__image_ext_content_type('.jpeg')
        # verify -----------------------
        expected = 'image/jpeg'
        actual = content_type
        msg = ("expected content type '%s', got '%s'" % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test___image_ext_content_type_raises_on_bad_ext(self):
        """_Image.__image_ext_content_type() raises on bad extension"""
        # verify -----------------------
        with self.assertRaises(TypeError):
            _Image._Image__image_ext_content_type('.xj7')

    def test___image_ext_content_type_raises_on_non_img_ext(self):
        """_Image.__image_ext_content_type() raises on non-image extension"""
        # verify -----------------------
        with self.assertRaises(TypeError):
            _Image._Image__image_ext_content_type('.xml')


class Test_ImageCollection(TestCase):
    """Test _ImageCollection"""
    def test_add_image_returns_matching_image(self):
        """_ImageCollection.add_image() returns existing image on match"""
        # setup ------------------------
        pkg = _Package(images_pptx_path)
        matching_idx = 4
        matching_image = pkg._images[matching_idx]
        # exercise ---------------------
        image = pkg._images.add_image(test_image_path)
        # verify -----------------------
        expected = matching_image
        actual = image
        msg = ("expected images[%d], got images[%d]"
               % (matching_idx, pkg._images.index(image)))
        self.assertEqual(expected, actual, msg)

    def test_add_image_adds_new_image(self):
        """_ImageCollection.add_image() adds new image on no match"""
        # setup ------------------------
        pkg = _Package(images_pptx_path)
        expected_partname = '/ppt/media/image8.png'
        expected_len = len(pkg._images) + 1
        expected_sha1 = '79769f1e202add2e963158b532e36c2c0f76a70c'
        # exercise ---------------------
        image = pkg._images.add_image(new_image_path)
        # verify -----------------------
        expected = (expected_partname, expected_len, expected_sha1)
        actual = (image.partname, len(pkg._images), image._sha1)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class Test_Package(TestCase):
    """Test _Package"""
    def setUp(self):
        self.test_pptx_path = absjoin(test_file_dir, 'test_python-pptx.pptx')
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)

    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)

    def test_construction_with_no_path_loads_default_template(self):
        """_Package() call with no path loads default template"""
        prs = _Package().presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))

    def test_instances_are_tracked(self):
        """_Package instances are tracked"""
        pkg = _Package()
        self.assertIn(pkg, _Package.instances())

    def test_instance_refs_are_garbage_collected(self):
        """_Package instance refs are garbage collected with old instances"""
        pkg = _Package()
        pkg1_repr = "%r" % pkg
        pkg = _Package()
        # pkg2_repr = "%r" % pkg
        gc.collect()
        reprs = [repr(pkg_inst) for pkg_inst in _Package.instances()]
        # log.debug("pkg1, pkg2, reprs: %s, %s, %s"
        #           % (pkg1_repr, pkg2_repr, reprs))
        assert_that(pkg1_repr, is_not(is_in(reprs)))

    def test_containing_returns_correct_pkg(self):
        """_Package.containing() returns right package instance"""
        # setup ------------------------
        pkg1 = _Package(test_pptx_path)
        pkg1.presentation  # does nothing, just needed to fake out pep8 warning
        pkg2 = _Package(test_pptx_path)
        slide = pkg2.presentation.slides[0]
        # exercise ---------------------
        found_pkg = _Package.containing(slide)
        # verify -----------------------
        expected = pkg2
        actual = found_pkg
        msg = "expected %r, got %r" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_containing_raises_on_no_pkg_contains_part(self):
        """_Package.containing(part) raises on no package contains part"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        pkg.presentation  # does nothing, just needed to fake out pep8 warning
        part = Mock(name='part')
        # verify -----------------------
        with self.assertRaises(KeyError):
            _Package.containing(part)

    def test_open_gathers_image_parts(self):
        """_Package open gathers image parts into image collection"""
        # exercise ---------------------
        pkg = _Package(images_pptx_path)
        # verify -----------------------
        expected = 7
        actual = len(pkg._Package__images)
        msg = "expected image count of %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_presentation_presentation_after_open(self):
        """_Package.presentation is instance of Presentation after open()"""
        # setup ------------------------
        cls = Presentation
        pkg = _Package()
        # exercise ---------------------
        obj = pkg.presentation
        # verify -----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)

    def test_it_should_have_core_props(self):
        """_Package should provide access to core document properties"""
        # setup ------------------------
        pkg = _Package()
        # verify -----------------------
        assert_that(pkg.core_properties, is_(instance_of(_CoreProperties)))

    def test_saved_file_has_plausible_contents(self):
        """_Package.save produces a .pptx with plausible contents"""
        # setup ------------------------
        pkg = _Package()
        # exercise ---------------------
        pkg.save(self.test_pptx_path)
        # verify -----------------------
        pkg = _Package(self.test_pptx_path)
        prs = pkg.presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))


class Test_Part(TestCase):
    """Test _Part"""
    def test_constructs_presentation_for_rt_officedocument(self):
        """_Part() returns Presentation for RT.OFFICE_DOCUMENT"""
        # setup ------------------------
        cls = Presentation
        # exercise ---------------------
        obj = _Part(RT.OFFICE_DOCUMENT, CT.PML_PRESENTATION_MAIN)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slide_for_rt_slide(self):
        """_Part() returns _Slide for RT.SLIDE"""
        # setup ------------------------
        cls = _Slide
        # exercise ---------------------
        obj = _Part(RT.SLIDE, CT.PML_SLIDE)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slidelayout_for_rt_slidelayout(self):
        """_Part() returns _SlideLayout for RT.SLIDE_LAYOUT"""
        # setup ------------------------
        cls = _SlideLayout
        # exercise ---------------------
        obj = _Part(RT.SLIDE_LAYOUT, CT.PML_SLIDE_LAYOUT)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slidemaster_for_rt_slidemaster(self):
        """_Part() returns _SlideMaster for RT.SLIDE_MASTER"""
        # setup ------------------------
        cls = _SlideMaster
        # exercise ---------------------
        obj = _Part(RT.SLIDE_MASTER, CT.PML_SLIDE_MASTER)
        # verify -----------------------
        self.assertIsInstance(obj, cls)

    def test_contructor_raises_on_invalid_prs_content_type(self):
        """_Part() raises on invalid presentation content type"""
        with self.assertRaises(InvalidPackageError):
            _Part(RT.OFFICE_DOCUMENT, CT.PML_SLIDE_MASTER)


class Test_Presentation(TestCase):
    """Test Presentation"""
    def setUp(self):
        self.prs = Presentation()

    def test__blob_rewrites_sldIdLst(self):
        """Presentation._blob rewrites sldIdLst"""
        # setup ------------------------
        rels = RelationshipCollectionBuilder()
        rels = rels.with_tuple_targets(2, RT.SLIDE_MASTER)
        rels = rels.with_tuple_targets(3, RT.SLIDE)
        rels = rels.with_ordering(RT.SLIDE_MASTER, RT.SLIDE)
        rels = rels.build()
        prs = Presentation()
        prs._relationships = rels
        prs.partname = '/ppt/presentation.xml'
        path = os.path.join(thisdir, 'test_files/presentation.xml')
        prs._element = oxml_parse(path).getroot()
        # exercise ---------------------
        blob = prs._blob
        # verify -----------------------
        presentation = oxml_fromstring(blob)
        sldIds = presentation.xpath('./p:sldIdLst/p:sldId', namespaces=nsmap)
        expected = ['rId3', 'rId4', 'rId5']
        actual = [sldId.get(qtag('r:id')) for sldId in sldIds]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_slidemasters_property_empty_on_construction(self):
        """Presentation.slidemasters property empty on construction"""
        # verify -----------------------
        self.assertIsSizedProperty(self.prs, 'slidemasters', 0)

    def test_slidemasters_correct_length_after_pkg_open(self):
        """Presentation.slidemasters correct length after load"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        prs = pkg.presentation
        # exercise ---------------------
        slidemasters = prs.slidemasters
        # verify -----------------------
        self.assertLength(slidemasters, 1)

    def test_slides_property_empty_on_construction(self):
        """Presentation.slides property empty on construction"""
        # verify -----------------------
        self.assertIsSizedProperty(self.prs, 'slides', 0)

    def test_slides_correct_length_after_pkg_open(self):
        """Presentation.slides correct length after load"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        prs = pkg.presentation
        # exercise ---------------------
        slides = prs.slides
        # verify -----------------------
        self.assertLength(slides, 1)


class Test_Slide(TestCase):
    """Test _Slide"""
    def setUp(self):
        self.sld = _Slide()

    def test_constructor_sets_correct_content_type(self):
        """_Slide constructor sets correct content type"""
        # exercise ---------------------
        content_type = self.sld._content_type
        # verify -----------------------
        expected = CT.PML_SLIDE
        actual = content_type
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_construction_adds_slide_layout_relationship(self):
        """_Slide(slidelayout) adds relationship slide->slidelayout"""
        # setup ------------------------
        slidelayout = _SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        # exercise ---------------------
        slide = _Slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(slide._relationships)
        msg = ("expected len(slide._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = slide._relationships[0]
        expected = ('rId1', RT.SLIDE_LAYOUT, slidelayout)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = "expected relationship\n%s\ngot\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__element_minimal_sld_on_construction(self):
        """_Slide._element is minimal sld on construction"""
        # setup ------------------------
        path = os.path.join(thisdir, 'test_files/minimal_slide.xml')
        # exercise ---------------------
        elm = self.sld._element
        # verify -----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        self.assertEqualLineByLine(expected_xml, elm)

    def test_slidelayout_property_none_on_construction(self):
        """_Slide.slidelayout property None on construction"""
        # verify -----------------------
        self.assertIsProperty(self.sld, 'slidelayout', None)

    def test__load_sets_slidelayout(self):
        """_Slide._load() sets slidelayout"""
        # setup ------------------------
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        slidelayout = Mock(name='slideLayout')
        slidelayout.partname = '/ppt/slideLayouts/slideLayout1.xml'
        rel = Mock(name='pptx.packaging._Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT.SLIDE_LAYOUT
        rel.target = slidelayout
        pkgpart = Mock(name='pptx.packaging.Part')
        with open(path, 'rb') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = [rel]
        part_dict = {slidelayout.partname: slidelayout}
        slide = self.sld._load(pkgpart, part_dict)
        # exercise ---------------------
        retval = slide.slidelayout
        # verify -----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test___minimal_element_xml(self):
        """_Slide.__minimal_element generates correct XML"""
        # setup ------------------------
        path = os.path.join(thisdir, 'test_files/minimal_slide.xml')
        # exercise ---------------------
        sld = self.sld._Slide__minimal_element
        # verify -----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        self.assertEqualLineByLine(expected_xml, sld)


class Test_SlideCollection(TestCase):
    """Test _SlideCollection"""
    def setUp(self):
        prs = Presentation()
        self.slides = _SlideCollection(prs)

    def test_add_slide_returns_slide(self):
        """_SlideCollection.add_slide() returns instance of _Slide"""
        # exercise ---------------------
        retval = self.slides.add_slide(None)
        # verify -----------------------
        self.assertIsInstance(retval, _Slide)

    def test_add_slide_sets_slidelayout(self):
        """
        _SlideCollection.add_slide() sets _Slide.slidelayout

        Kind of a throw-away test, but was helpful for initial debugging.
        """
        # setup ------------------------
        slidelayout = Mock(name='slideLayout')
        slidelayout.shapes = []
        slide = self.slides.add_slide(slidelayout)
        # exercise ---------------------
        retval = slide.slidelayout
        # verify -----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_slide_adds_slide_layout_relationship(self):
        """_SlideCollection.add_slide() adds relationship prs->slide"""
        # setup ------------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = _SlideLayout()
        slidelayout._shapes = []
        # exercise ---------------------
        slide = slides.add_slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(prs._relationships)
        msg = ("expected len(prs._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = prs._relationships[0]
        expected = ('rId1', RT.SLIDE, slide)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = ("expected relationship 1:, got 2:\n1: %s\n2: %s"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test_add_slide_sets_partname(self):
        """_SlideCollection.add_slide() sets partname of new slide"""
        # setup ------------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = _SlideLayout()
        slidelayout._shapes = []
        # exercise ---------------------
        slide = slides.add_slide(slidelayout)
        # verify -----------------------
        expected = '/ppt/slides/slide1.xml'
        actual = slide.partname
        msg = "expected partname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class Test_SlideLayout(TestCase):
    """Test _SlideLayout"""
    def setUp(self):
        self.slidelayout = _SlideLayout()

    def __loaded_slidelayout(self, prs_slidemaster=None):
        """
        Return _SlideLayout instance loaded using mocks. *prs_slidemaster* is
        an already-loaded model-side _SlideMaster instance (or mock, as
        appropriate to calling test).
        """
        # partname for related slideMaster
        sldmaster_partname = '/ppt/slideMasters/slideMaster1.xml'
        # path to test slideLayout XML
        slidelayout_path = absjoin(test_file_dir, 'slideLayout1.xml')
        # model-side slideMaster part
        if prs_slidemaster is None:
            prs_slidemaster = Mock(spec=_SlideMaster)
        # a part dict containing the already-loaded model-side slideMaster
        loaded_part_dict = {sldmaster_partname: prs_slidemaster}
        # a slideMaster package part for rel target
        pkg_slidemaster_part = Mock(spec=pptx.packaging.Part)
        pkg_slidemaster_part.partname = sldmaster_partname
        # a package-side relationship from slideLayout to its slideMaster
        rel = Mock(name='pptx.packaging._Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT.SLIDE_MASTER
        rel.target = pkg_slidemaster_part
        # the slideLayout package part to send to _load()
        pkg_slidelayout_part = Mock(spec=pptx.packaging.Part)
        pkg_slidelayout_part.relationships = [rel]
        with open(slidelayout_path, 'rb') as f:
            pkg_slidelayout_part.blob = f.read()
        # _load and return
        slidelayout = _SlideLayout()
        return slidelayout._load(pkg_slidelayout_part, loaded_part_dict)

    def test__load_sets_slidemaster(self):
        """_SlideLayout._load() sets slidemaster"""
        # setup ------------------------
        prs_slidemaster = Mock(spec=_SlideMaster)
        # exercise ---------------------
        loaded_slidelayout = self.__loaded_slidelayout(prs_slidemaster)
        # verify -----------------------
        expected = prs_slidemaster
        actual = loaded_slidelayout.slidemaster
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_slidemaster_is_readonly(self):
        """_SlideLayout.slidemaster is read-only"""
        # verify -----------------------
        self.assertIsReadOnly(self.slidelayout, 'slidemaster')

    def test_slidemaster_raises_on_ref_before_assigned(self):
        """_SlideLayout.slidemaster raises on referenced before assigned"""
        with self.assertRaises(AssertionError):
            self.slidelayout.slidemaster


class Test_SlideMaster(TestCase):
    """Test _SlideMaster"""
    def setUp(self):
        self.sldmaster = _SlideMaster()

    def test_slidelayouts_property_empty_on_construction(self):
        """_SlideMaster.slidelayouts property empty on construction"""
        # verify -----------------------
        self.assertIsSizedProperty(self.sldmaster, 'slidelayouts', 0)

    def test_slidelayouts_correct_length_after_open(self):
        """_SlideMaster.slidelayouts correct length after open"""
        # setup ------------------------
        pkg = _Package(test_pptx_path)
        slidemaster = pkg.presentation.slidemasters[0]
        # exercise ---------------------
        slidelayouts = slidemaster.slidelayouts
        # verify -----------------------
        self.assertLength(slidelayouts, 11)
