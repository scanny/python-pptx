# -*- coding: utf-8 -*-
#
# test_presentation.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.presentation module."""

import gc
import os
import re

from hamcrest import assert_that, is_, is_in, is_not, equal_to
from StringIO import StringIO

try:
    from PIL import Image as PILImage
except ImportError:
    import Image as PILImage

from mock import Mock, patch, PropertyMock

import pptx.presentation

from pptx.constants import MSO
from pptx.exceptions import InvalidPackageError

from pptx.oxml import _SubElement, oxml_fromstring, oxml_tostring, oxml_parse

from pptx.packaging import prettify_nsdecls

from pptx.presentation import (
    Package, Collection, _RelationshipCollection, _Relationship, Presentation,
    PartCollection, BasePart, Part, SlideCollection, BaseSlide, Slide,
    SlideLayout, SlideMaster, Image, ShapeCollection, BaseShape, Shape,
    Placeholder, TextFrame, _Font, Paragraph, Run, _to_unicode)

from pptx.spec import namespaces, qtag
from pptx.spec import (
    CT_PRESENTATION, CT_SLIDE, CT_SLIDELAYOUT, CT_SLIDEMASTER)
from pptx.spec import (
    RT_IMAGE, RT_OFFICEDOCUMENT, RT_PRESPROPS, RT_SLIDE, RT_SLIDELAYOUT,
    RT_SLIDEMASTER)
from pptx.spec import (
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL, PH_TYPE_TITLE, PH_ORIENT_HORZ,
    PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER)
from pptx.util import Inches, Px, Pt
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
nsprefix_decls = (
    ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xm'
    'lns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="h'
    'ttp://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def _empty_spTree():
    xml = ('<p:spTree xmlns:p="http://schemas.openxmlformats.org/'
           'presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.'
           'org/drawingml/2006/main"><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
           '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>')
    return oxml_fromstring(xml)


def _sldLayout1():
    path = os.path.join(thisdir, 'test_files/slideLayout1.xml')
    sldLayout = oxml_parse(path).getroot()
    return sldLayout


def _sldLayout1_shapes():
    sldLayout = _sldLayout1()
    spTree = sldLayout.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
    shapes = ShapeCollection(spTree)
    return shapes


def _strip_nsdecls(xml):
    ptrn_str = ' xmlns(:[a-z]+)="[a-zA-Z0-9:/.]+"'
    nsdecl_re = re.compile(ptrn_str)
    return nsdecl_re.sub('', xml)


def _txbox_xml():
    xml = (
        '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n<p:sp'
        ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main'
        '"\n      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/m'
        'ain">\n  <p:nvSpPr>\n    <p:cNvPr id="2" name="TextBox 1"/>\n    <p:'
        'cNvSpPr txBox="1"/>\n    <p:nvPr/>\n  </p:nvSpPr>\n  <p:spPr>\n    <'
        'a:xfrm>\n      <a:off x="914400" y="1828800"/>\n      <a:ext cx="137'
        '1600" cy="457200"/>\n    </a:xfrm>\n    <a:prstGeom prst="rect">\n  '
        '    <a:avLst/>\n    </a:prstGeom>\n    <a:noFill/>\n  </p:spPr>\n  <'
        'p:txBody>\n    <a:bodyPr wrap="none">\n      <a:spAutoFit/>\n    </a'
        ':bodyPr>\n    <a:lstStyle/>\n    <a:p/>\n  </p:txBody>\n</p:sp>')
    return xml


class PartBuilder(object):
    """Builder class for test Parts"""
    def __init__(self):
        self.partname = '/ppt/slides/slide1.xml'

    def with_partname(self, partname):
        self.partname = partname
        return self

    def build(self):
        p = BasePart()
        p.partname = self.partname
        return p


class RelationshipCollectionBuilder(object):
    """Builder class for test RelationshipCollections"""
    partname_tmpls = {RT_SLIDEMASTER: '/ppt/slideMasters/slideMaster%d.xml',
                      RT_SLIDE:       '/ppt/slides/slide%d.xml'}

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


class TestBasePart(TestCase):
    """Test BasePart"""
    def setUp(self):
        self.basepart = BasePart()
        self.cls = BasePart

    def test__add_relationship_adds_specified_relationship(self):
        """BasePart._add_relationship adds specified relationship"""
        # setup -----------------------
        reltype = RT_IMAGE
        target = Mock(name='image')
        # exercise --------------------
        rel = self.basepart._add_relationship(reltype, target)
        # verify ----------------------
        expected = ('rId1', reltype, target)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__blob_value_for_binary_part(self):
        """BasePart._blob value is correct for binary part"""
        # setup -----------------------
        blob = '0123456789'
        self.basepart._load_blob = blob
        self.basepart.partname = '/docProps/thumbnail.jpeg'
        # exercise --------------------
        retval = self.basepart._blob
        # verify ----------------------
        expected = blob
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__blob_value_for_xml_part(self):
        """BasePart._blob value is correct for XML part"""
        # setup -----------------------
        elm = oxml_fromstring('<root><elm1 attr="one"/></root>')
        self.basepart._element = elm
        self.basepart.partname = '/ppt/presentation.xml'
        # exercise --------------------
        retval = self.basepart._blob
        # verify ----------------------
        expected = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"\
                   '\n<root>\n  <elm1 attr="one"/>\n</root>\n'
        actual = retval
        msg = "expected: \n'%s'\n, got \n'%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__content_type_raises_on_accessed_before_assigned(self):
        """BasePart._content_type raises on access before assigned"""
        with self.assertRaises(ValueError):
            self.basepart._content_type

    def test__load_sets__element_for_xml_part(self):
        """BasePart._load() sets _element for xml part"""
        # setup -----------------------
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/presentation.xml'
        pkgpart.blob = '<root><elm1   attr="spam"/></root>'
        pkgpart.relationships = []
        part_dict = {}
        part = self.basepart._load(pkgpart, part_dict)
        # exercise --------------------
        elm = part._element
        # verify ----------------------
        expected = '<root><elm1 attr="spam"/></root>'
        actual = oxml_tostring(elm)
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_observable_on_partname(self):
        """BasePart observable on partname value change"""
        # setup -----------------------
        old_partname = '/ppt/slides/slide1.xml'
        new_partname = '/ppt/slides/slide2.xml'
        observer = Mock()
        self.basepart.partname = old_partname
        self.basepart.add_observer(observer)
        # exercise --------------------
        self.basepart.partname = new_partname
        # verify ----------------------
        observer.notify.assert_called_with(self.basepart, 'partname',
                                           new_partname)

    def test_partname_setter(self):
        """BasePart.partname setter stores passed value"""
        # setup -----------------------
        partname = '/ppt/presentation.xml'
        # exercise ----------------
        self.basepart.partname = partname
        # verify ------------------
        expected = partname
        actual = self.basepart.partname
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestBaseShape(TestCase):
    """Test BaseShape"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = oxml_parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:pic'
        pic = self.sld.xpath(xpath, namespaces=nsmap)[0]
        self.base_shape = BaseShape(pic)

    def test_has_textframe_value(self):
        """BaseShape.has_textframe value correct"""
        # setup -----------------------
        spTree = self.sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        shapes = ShapeCollection(spTree)
        indexes = []
        # exercise --------------------
        for idx, shape in enumerate(shapes):
            if shape.has_textframe:
                indexes.append(idx)
        # verify ----------------------
        expected = [0, 1, 3, 5, 6]
        actual = indexes
        msg = "expected txBody element in shapes %s, got %s"\
              % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_id_value(self):
        """BaseShape.id value is correct"""
        # exercise --------------------
        id = self.base_shape.id
        # verify ----------------------
        expected = 6
        actual = id
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_is_placeholder_true_for_placeholder(self):
        """BaseShape.is_placeholder True for placeholder shape"""
        # setup -----------------------
        xpath = './p:cSld/p:spTree/p:sp'
        sp = self.sld.xpath(xpath, namespaces=nsmap)[0]
        base_shape = BaseShape(sp)
        # verify ----------------------
        actual = base_shape.is_placeholder
        msg = "expected True, got %s" % (actual)
        self.assertTrue(actual, msg)

    def test_is_placeholder_false_for_non_placeholder(self):
        """BaseShape.is_placeholder False for non-placeholder shape"""
        # verify ----------------------
        actual = self.base_shape.is_placeholder
        msg = "expected False, got %s" % (actual)
        self.assertFalse(actual, msg)

    def test__is_title_true_for_title_placeholder(self):
        """BaseShape._is_title True for title placeholder shape"""
        # setup -----------------------
        xpath = './p:cSld/p:spTree/p:sp'
        title_placeholder_sp = self.sld.xpath(xpath, namespaces=nsmap)[0]
        base_shape = BaseShape(title_placeholder_sp)
        # verify ----------------------
        actual = base_shape._is_title
        msg = "expected True, got %s" % (actual)
        self.assertTrue(actual, msg)

    def test__is_title_false_for_no_ph_element(self):
        """BaseShape._is_title False on shape has no <p:ph> element"""
        # setup -----------------------
        self.base_shape._element = Mock(name='_element')
        self.base_shape._element.xpath.return_value = []
        # verify ----------------------
        assert_that(self.base_shape._is_title, is_(False))

    def test_name_value(self):
        """BaseShape.name value is correct"""
        # exercise --------------------
        name = self.base_shape.name
        # verify ----------------------
        expected = 'Picture 5'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_textframe_raises_on_no_textframe(self):
        """BaseShape.textframe raises on shape with no text frame"""
        with self.assertRaises(ValueError):
            self.base_shape.textframe

    def test_text_setter_structure_and_value(self):
        """assign to BaseShape.text yields single run para set to value"""
        # setup -----------------------
        test_text = 'python-pptx was here!!'
        xpath = './p:cSld/p:spTree/p:sp'
        textbox_sp = self.sld.xpath(xpath, namespaces=nsmap)[2]
        base_shape = BaseShape(textbox_sp)
        # exercise --------------------
        base_shape.text = test_text
        # verify paragraph count ------
        expected = 1
        actual = len(base_shape.textframe.paragraphs)
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = base_shape.textframe.paragraphs[0].runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_text_setter_raises_on_no_textframe(self):
        """assignment to BaseShape.text raises for shape with no text frame"""
        with self.assertRaises(TypeError):
            self.base_shape.text = 'test text'


class TestBaseSlide(TestCase):
    """Test BaseSlide"""
    def setUp(self):
        self.base_slide = BaseSlide()

    def test_name_value(self):
        """BaseSlide.name value is correct"""
        # setup -----------------------
        self.base_slide._element = _sldLayout1()
        # exercise --------------------
        name = self.base_slide.name
        # verify ----------------------
        expected = 'Title Slide'
        actual = name
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_shapes_size_after__load(self):
        """BaseSlide.shapes is expected size after _load()"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        pkgpart = Mock(name='pptx.packaging.Part')
        pkgpart.partname = '/ppt/slides/slide1.xml'
        with open(path, 'r') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = []
        part_dict = {}
        self.base_slide._load(pkgpart, part_dict)
        # exercise --------------------
        shapes = self.base_slide.shapes
        # verify ----------------------
        self.assertLength(shapes, 9)


class TestCollection(TestCase):
    """Test Collection"""
    def setUp(self):
        self.collection = Collection()

    def test_indexable(self):
        """Collection is indexable (e.g. no TypeError on 'collection[0]')"""
        # verify ----------------------
        try:
            self.collection[0]
        except TypeError:
            msg = "'Collection' object does not support indexing"
            self.fail(msg)
        except IndexError:
            pass

    def test_is_container(self):
        """Collection is container (e.g. 'x in collection' works)"""
        # verify ----------------------
        try:
            1 in self.collection
        except TypeError:
            msg = "'Collection' object is not container"
            self.fail(msg)

    def test_iterable(self):
        """Collection is iterable"""
        # verify ----------------------
        try:
            for x in self.collection:
                pass
        except TypeError:
            msg = "'Collection' object is not iterable"
            self.fail(msg)

    def test_sized(self):
        """Collection is sized (e.g. 'len(collection)' works)"""
        # verify ----------------------
        try:
            len(self.collection)
        except TypeError:
            msg = "object of type 'Collection' has no len()"
            self.fail(msg)

    def test__values_property_empty_on_construction(self):
        """Collection._values property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.collection, '_values', 0)


class Test_Font(TestCase):
    """Test _Font class"""
    def setUp(self):
        self.rPr_xml = '<a:rPr%s/>' % nsprefix_decls
        self.rPr = oxml_fromstring(self.rPr_xml)
        self.font = _Font(self.rPr)

    def test_get_bold_setting(self):
        """_Font.bold returns True on bold font weight"""
        # setup -----------------------
        rPr_xml = '<a:rPr%s b="1"/>' % nsprefix_decls
        rPr = oxml_fromstring(rPr_xml)
        font = _Font(rPr)
        # verify ----------------------
        assert_that(self.font.bold, is_(False))
        assert_that(font.bold, is_(True))

    def test_set_bold(self):
        """Setting _Font.bold to True selects bold font weight"""
        # setup -----------------------
        expected_rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main" b="1"/>')
        # exercise --------------------
        self.font.bold = True
        # verify ----------------------
        rPr_xml = oxml_tostring(self.font._Font__rPr)
        assert_that(rPr_xml, is_(equal_to(expected_rPr_xml)))

    def test_clear_bold(self):
        """Setting _Font.bold to False selects normal font weight"""
        # setup -----------------------
        rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main" b="1"/>')
        rPr = oxml_fromstring(rPr_xml)
        font = _Font(rPr)
        expected_rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main"/>')
        # exercise --------------------
        font.bold = False
        # verify ----------------------
        rPr_xml = oxml_tostring(font._Font__rPr)
        assert_that(rPr_xml, is_(equal_to(expected_rPr_xml)))

    def test_set_font_size(self):
        """Assignment to _Font.size changes font size"""
        # setup -----------------------
        newfontsize = 2400
        expected_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main" sz="%d"/>') % newfontsize
        # exercise --------------------
        self.font.size = newfontsize
        # verify ----------------------
        actual_xml = oxml_tostring(self.font._Font__rPr)
        assert_that(actual_xml, is_(equal_to(expected_xml)))


class TestImage(TestCase):
    """Test Image"""
    def test_construction_from_file(self):
        """Image(path) constructor produces correct attribute values"""
        # exercise --------------------
        image = Image(test_image_path)
        # verify ----------------------
        assert_that(image.ext, is_(equal_to('.jpeg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))

    def test_construction_from_stream(self):
        """Image(stream) construction produces correct attribute values"""
        # exercise --------------------
        with open(test_image_path) as f:
            stream = StringIO(f.read())
        image = Image(stream)
        # verify ----------------------
        assert_that(image.ext, is_(equal_to('.jpg')))
        assert_that(image._content_type, is_(equal_to('image/jpeg')))
        assert_that(len(image._blob), is_(equal_to(3277)))

    def test_construction_from_file_raises_on_bad_path(self):
        """Image(path) constructor raises on bad path"""
        # verify ----------------------
        with self.assertRaises(IOError):
            Image('foobar27.png')

    def test___ext_from_image_stream_raises_on_incompatible_format(self):
        """Image.__ext_from_image_stream() raises on incompatible format"""
        # verify ----------------------
        with self.assertRaises(ValueError):
            with open(test_bmp_path) as stream:
                Image._Image__ext_from_image_stream(stream)

    def test___image_ext_content_type_known_type(self):
        """Image.__image_ext_content_type() correct for known content type"""
        # exercise --------------------
        content_type = Image._Image__image_ext_content_type('.jpeg')
        # verify ----------------------
        expected = 'image/jpeg'
        actual = content_type
        msg = ("expected content type '%s', got '%s'" % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test___image_ext_content_type_raises_on_bad_ext(self):
        """Image.__image_ext_content_type() raises on bad extension"""
        # verify ----------------------
        with self.assertRaises(TypeError):
            Image._Image__image_ext_content_type('.xj7')

    def test___image_ext_content_type_raises_on_non_img_ext(self):
        """Image.__image_ext_content_type() raises on non-image extension"""
        # verify ----------------------
        with self.assertRaises(TypeError):
            Image._Image__image_ext_content_type('.xml')


class TestImageCollection(TestCase):
    """Test ImageCollection"""
    def test_add_image_returns_matching_image(self):
        """ImageCollection.add_image() returns existing image on match"""
        # setup -----------------------
        pkg = Package(images_pptx_path)
        matching_idx = 4
        matching_image = pkg._images[matching_idx]
        # exercise --------------------
        image = pkg._images.add_image(test_image_path)
        # verify ----------------------
        expected = matching_image
        actual = image
        msg = ("expected images[%d], got images[%d]"
               % (matching_idx, pkg._images.index(image)))
        self.assertEqual(expected, actual, msg)

    def test_add_image_adds_new_image(self):
        """ImageCollection.add_image() adds new image on no match"""
        # setup -----------------------
        pkg = Package(images_pptx_path)
        expected_partname = '/ppt/media/image8.png'
        expected_len = len(pkg._images) + 1
        expected_sha1 = '79769f1e202add2e963158b532e36c2c0f76a70c'
        # exercise --------------------
        image = pkg._images.add_image(new_image_path)
        # verify ----------------------
        expected = (expected_partname, expected_len, expected_sha1)
        actual = (image.partname, len(pkg._images), image._sha1)
        msg = "\nExpected: %s\n     Got: %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestPackage(TestCase):
    """Test Package"""
    def setUp(self):
        self.test_pptx_path = absjoin(test_file_dir, 'test_python-pptx.pptx')
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)

    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)

    def test_construction_with_no_path_loads_default_template(self):
        """Package() call with no path loads default template"""
        prs = Package().presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))

    def test_instances_are_tracked(self):
        """Package instances are tracked"""
        pkg = Package()
        self.assertIn(pkg, Package.instances())

    def test_instance_refs_are_garbage_collected(self):
        """Package instances are tracked"""
        pkg = Package()
        pkg1_repr = "%r" % pkg
        pkg = Package()
        # pkg2_repr = "%r" % pkg
        gc.collect()
        reprs = [repr(pkg_inst) for pkg_inst in Package.instances()]
        # log.debug("pkg1, pkg2, reprs: %s, %s, %s"
        #           % (pkg1_repr, pkg2_repr, reprs))
        assert_that(pkg1_repr, is_not(is_in(reprs)))

    def test_containing_returns_correct_pkg(self):
        """Package.containing() returns right package instance"""
        # setup -----------------------
        pkg1 = Package(test_pptx_path)
        pkg2 = Package(test_pptx_path)
        slide = pkg2.presentation.slides[0]
        # exercise --------------------
        found_pkg = Package.containing(slide)
        # verify ----------------------
        expected = pkg2
        actual = found_pkg
        msg = "expected %r, got %r" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_containing_raises_on_no_pkg_contains_part(self):
        """Package.containing(part) raises on no package contains part"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        part = Mock(name='part')
        # verify ----------------------
        with self.assertRaises(KeyError):
            Package.containing(part)

    def test_open_gathers_image_parts(self):
        """Package open gathers image parts into image collection"""
        # exercise --------------------
        pkg = Package(images_pptx_path)
        # verify ----------------------
        expected = 7
        actual = len(pkg._Package__images)
        msg = "expected image count of %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_presentation_presentation_after_open(self):
        """Package.presentation is instance of Presentation after open()"""
        # setup -----------------------
        cls = Presentation
        pkg = Package()
        # exercise --------------------
        obj = pkg.presentation
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)

    def test_saved_file_has_plausible_contents(self):
        """Package.save produces a .pptx with plausible contents"""
        # setup -----------------------
        pkg = Package()
        # exercise --------------------
        pkg.save(self.test_pptx_path)
        # verify ----------------------
        pkg = Package(self.test_pptx_path)
        prs = pkg.presentation
        assert_that(prs, is_not(None))
        slidemasters = prs.slidemasters
        assert_that(slidemasters, is_not(None))
        assert_that(len(slidemasters), is_(1))
        slidelayouts = slidemasters[0].slidelayouts
        assert_that(slidelayouts, is_not(None))
        assert_that(len(slidelayouts), is_(11))


class TestParagraph(TestCase):
    """Test Paragraph"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = oxml_parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody/a:p'
        self.pList = self.sld.xpath(xpath, namespaces=nsmap)

        self.test_text = 'test text'
        self.p_xml = ('<a:p%s><a:r><a:t>%s</a:t></a:r></a:p>' %
                      (nsprefix_decls, self.test_text))
        self.p = oxml_fromstring(self.p_xml)
        self.paragraph = Paragraph(self.p)

    def test_runs_size(self):
        """Paragraph.runs is expected size"""
        # setup -----------------------
        actual_lengths = []
        for p in self.pList:
            paragraph = Paragraph(p)
            # exercise ----------------
            actual_lengths.append(len(paragraph.runs))
        # verify ------------------
        expected = [0, 0, 2, 1, 1, 1]
        actual = actual_lengths
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_run_increments_run_count(self):
        """Paragraph.add_run() increments run count"""
        # setup -----------------------
        p_elm = self.pList[0]
        paragraph = Paragraph(p_elm)
        # exercise --------------------
        paragraph.add_run()
        # verify ----------------------
        expected = 1
        actual = len(paragraph.runs)
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_clear_removes_all_runs(self):
        """Paragraph.clear() removes all runs from paragraph"""
        # setup -----------------------
        p = self.pList[2]
        _SubElement(p, 'a:pPr')
        paragraph = Paragraph(p)
        assert_that(len(paragraph.runs), is_(equal_to(2)))
        # exercise --------------------
        paragraph.clear()
        # verify ----------------------
        assert_that(len(paragraph.runs), is_(equal_to(0)))

    def test_clear_preserves_paragraph_properties(self):
        """Paragraph.clear() preserves paragraph properties"""
        # setup -----------------------
        nsprefix_decls = (
            ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
            'main"')
        p_xml = ('<a:p%s><a:pPr lvl="1"/><a:r><a:t>%s</a:t></a:r></a:p>' %
                 (nsprefix_decls, self.test_text))
        p_elm = oxml_fromstring(p_xml)
        paragraph = Paragraph(p_elm)
        expected_p_xml = '<a:p%s><a:pPr lvl="1"/></a:p>' % nsprefix_decls
        # exercise --------------------
        paragraph.clear()
        # verify ----------------------
        p_xml = oxml_tostring(paragraph._Paragraph__p)
        assert_that(p_xml, is_(equal_to(expected_p_xml)))

    def test_set_font_size(self):
        """Assignment to Paragraph.font.size changes font size"""
        # setup -----------------------
        newfontsize = Pt(54.3)
        expected_xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n<'
            'a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/ma'
            'in">\n  <a:pPr>\n    <a:defRPr sz="5430"/>\n  </a:pPr>\n  <a:r>'
            '\n    <a:t>test text</a:t>\n  </a:r>\n</a:p>\n')
        # exercise --------------------
        self.paragraph.font.size = newfontsize
        # verify ----------------------
        p_xml = prettify_nsdecls(oxml_tostring(self.paragraph._Paragraph__p,
            encoding='UTF-8', pretty_print=True, standalone=True))
        p_xml_lines = p_xml.split('\n')
        expected_xml_lines = expected_xml.split('\n')
        for idx, line in enumerate(p_xml_lines):
            # msg = '\n\n%s' % sld_xml
            msg = "\n\nexpected:\n\n%s\n\nbut got:\n\n%s" % (expected_xml, p_xml)
            self.assertEqual(line, expected_xml_lines[idx], msg)

    def test_text_setter_sets_single_run_text(self):
        """assignment to Paragraph.text creates single run containing value"""
        # setup -----------------------
        test_text = 'python-pptx was here!!'
        p_elm = self.pList[2]
        paragraph = Paragraph(p_elm)
        # exercise --------------------
        paragraph.text = test_text
        # verify run count ------------
        expected = 1
        actual = len(paragraph.runs)
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = paragraph.runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_text_accepts_non_ascii_strings(self):
        """assignment of non-ASCII string to text does not raise"""
        # setup -----------------------
        _7bit_string = 'String containing only 7-bit (ASCII) characters'
        _8bit_string = '8-bit string: Hér er texti með íslenskum stöfum.'
        _utf8_literal = u'unicode literal: Hér er texti með íslenskum stöfum.'
        _utf8_from_8bit = unicode('utf-8 unicode: Hér er texti', 'utf-8')
        # verify ----------------------
        try:
            text = _7bit_string    ; self.paragraph.text = text
            text = _8bit_string    ; self.paragraph.text = text
            text = _utf8_literal   ; self.paragraph.text = text
            text = _utf8_from_8bit ; self.paragraph.text = text
        except ValueError:
            msg = "Paragraph.text rejects valid text string '%s'" % text
            self.fail(msg)


class TestPart(TestCase):
    """Test Part"""
    def test_constructs_presentation_for_rt_officedocument(self):
        """Part() returns Presentation for RT_OFFICEDOCUMENT"""
        # setup -----------------------
        cls = Presentation
        # exercise --------------------
        obj = Part(RT_OFFICEDOCUMENT, CT_PRESENTATION)
        # verify ----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slide_for_rt_slide(self):
        """Part() returns Slide for RT_SLIDE"""
        # setup -----------------------
        cls = Slide
        # exercise --------------------
        obj = Part(RT_SLIDE, CT_SLIDE)
        # verify ----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slidelayout_for_rt_slidelayout(self):
        """Part() returns SlideLayout for RT_SLIDELAYOUT"""
        # setup -----------------------
        cls = SlideLayout
        # exercise --------------------
        obj = Part(RT_SLIDELAYOUT, CT_SLIDELAYOUT)
        # verify ----------------------
        self.assertIsInstance(obj, cls)

    def test_constructs_slidemaster_for_rt_slidemaster(self):
        """Part() returns SlideMaster for RT_SLIDEMASTER"""
        # setup -----------------------
        cls = SlideMaster
        # exercise --------------------
        obj = Part(RT_SLIDEMASTER, CT_SLIDEMASTER)
        # verify ----------------------
        self.assertIsInstance(obj, cls)

    def test_contructor_raises_on_invalid_prs_content_type(self):
        """Part() raises on invalid presentation content type"""
        with self.assertRaises(InvalidPackageError):
            Part(RT_OFFICEDOCUMENT, CT_SLIDEMASTER)


class TestPartCollection(TestCase):
    """Test PartCollection"""
    def test__loadpart_sorts_loaded_parts(self):
        """PartCollection._loadpart sorts loaded parts"""
        # setup -----------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = Mock(name='part1'); part1.partname = partname1
        part2 = Mock(name='part2'); part2.partname = partname2
        part3 = Mock(name='part3'); part3.partname = partname3
        parts = PartCollection()
        # exercise --------------------
        parts._loadpart(part2)
        parts._loadpart(part3)
        parts._loadpart(part1)
        # verify ----------------------
        expected = [partname1, partname2, partname3]
        actual = [part.partname for part in parts]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestPlaceholder(TestCase):
    """Test Placeholder"""
    def test_property_values(self):
        """Placeholder property values are correct"""
        # setup -----------------------
        expected_values =\
            ( (PH_TYPE_CTRTITLE, PH_ORIENT_HORZ, PH_SZ_FULL,     0)
            , (PH_TYPE_DT,       PH_ORIENT_HORZ, PH_SZ_HALF,    10)
            , (PH_TYPE_SUBTITLE, PH_ORIENT_VERT, PH_SZ_FULL,     1)
            , (PH_TYPE_TBL,      PH_ORIENT_HORZ, PH_SZ_QUARTER, 14)
            , (PH_TYPE_SLDNUM,   PH_ORIENT_HORZ, PH_SZ_QUARTER, 12)
            , (PH_TYPE_FTR,      PH_ORIENT_HORZ, PH_SZ_QUARTER, 11)
            )
        shapes = _sldLayout1_shapes()
        # exercise --------------------
        for idx, sp in enumerate(shapes):
            ph = Placeholder(sp)
            values = (ph.type, ph.orient, ph.sz, ph.idx)
            # verify ----------------------
            expected = expected_values[idx]
            actual = values
            msg = "expected shapes[%d] values %s, got %s"\
                   % (idx, expected, actual)
            self.assertEqual(expected, actual, msg)


class TestPresentation(TestCase):
    """Test Presentation"""
    def setUp(self):
        self.prs = Presentation()

    def test__blob_rewrites_sldIdLst(self):
        """Presentation._blob rewrites sldIdLst"""
        # setup -----------------------
        relationships = RelationshipCollectionBuilder()\
                       .with_tuple_targets(2, RT_SLIDEMASTER)\
                       .with_tuple_targets(3, RT_SLIDE)\
                       .with_ordering(RT_SLIDEMASTER, RT_SLIDE)\
                       .build()
        prs = Presentation()
        prs._relationships = relationships
        prs.partname = '/ppt/presentation.xml'
        path = os.path.join(thisdir, 'test_files/presentation.xml')
        prs._element = oxml_parse(path).getroot()
        # exercise --------------------
        blob = prs._blob
        # verify ----------------------
        presentation = oxml_fromstring(blob)
        sldIds = presentation.xpath('./p:sldIdLst/p:sldId', namespaces=nsmap)
        expected = ['rId3', 'rId4', 'rId5']
        actual = [sldId.get(qtag('r:id')) for sldId in sldIds]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_slidemasters_property_empty_on_construction(self):
        """Presentation.slidemasters property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.prs, 'slidemasters', 0)

    def test_slidemasters_correct_length_after_pkg_open(self):
        """Presentation.slidemasters correct length after load"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        prs = pkg.presentation
        # exercise --------------------
        slidemasters = prs.slidemasters
        # verify ----------------------
        self.assertLength(slidemasters, 1)

    def test_slides_property_empty_on_construction(self):
        """Presentation.slides property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.prs, 'slides', 0)

    def test_slides_correct_length_after_pkg_open(self):
        """Presentation.slides correct length after load"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        prs = pkg.presentation
        # exercise --------------------
        slides = prs.slides
        # verify ----------------------
        self.assertLength(slides, 1)


class Test_Relationship(TestCase):
    """Test _Relationship"""
    def setUp(self):
        rId = 'rId1'
        reltype = RT_SLIDE
        target_part = None
        self.rel = _Relationship(rId, reltype, target_part)

    def test_constructor_raises_on_bad_rId(self):
        """_Relationship constructor raises on non-standard rId"""
        with self.assertRaises(AssertionError):
            _Relationship('Non-std14', None, None)

    def test__num_value(self):
        """_Relationship._num value is correct"""
        # setup -----------------------
        num = 91
        rId = 'rId%d' % num
        rel = _Relationship(rId, None, None)
        # verify ----------------------
        assert_that(rel._num, is_(equal_to(num)))

    def test__num_value_on_non_standard_rId(self):
        """_Relationship._num value is correct for non-standard rId"""
        # setup -----------------------
        rel = _Relationship('rIdSm', None, None)
        # verify ----------------------
        assert_that(rel._num, is_(equal_to(9999)))

    def test__rId_setter(self):
        """Relationship._rId setter stores passed value"""
        # setup -----------------------
        rId = 'rId9'
        # exercise ----------------
        self.rel._rId = rId
        # verify ------------------
        expected = rId
        actual = self.rel._rId
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class Test_RelationshipCollection(TestCase):
    """Test _RelationshipCollection"""
    def setUp(self):
        self.relationships = _RelationshipCollection()

    def __reltype_ordering_mock(self):
        """
        Return RelationshipCollection instance with mocked-up contents
        suitable for testing _reltype_ordering.
        """
        # setup -----------------------
        partnames =\
            [ '/ppt/slides/slide4.xml'
            , '/ppt/slideLayouts/slideLayout1.xml'
            , '/ppt/slideMasters/slideMaster1.xml'
            , '/ppt/slides/slide1.xml'
            , '/ppt/presProps.xml'
            ]
        part1 = Mock(name='part1'); part1.partname = partnames[0]
        part2 = Mock(name='part2'); part2.partname = partnames[1]
        part3 = Mock(name='part3'); part3.partname = partnames[2]
        part4 = Mock(name='part4'); part4.partname = partnames[3]
        part5 = Mock(name='part5'); part5.partname = partnames[4]
        rel1 = _Relationship('rId1', RT_SLIDE,       part1)
        rel2 = _Relationship('rId2', RT_SLIDELAYOUT, part2)
        rel3 = _Relationship('rId3', RT_SLIDEMASTER, part3)
        rel4 = _Relationship('rId4', RT_SLIDE,       part4)
        rel5 = _Relationship('rId5', RT_PRESPROPS,   part5)
        relationships = _RelationshipCollection()
        relationships._additem(rel1)
        relationships._additem(rel2)
        relationships._additem(rel3)
        relationships._additem(rel4)
        relationships._additem(rel5)
        return (relationships, partnames)

    def test__additem_raises_on_dup_rId(self):
        """_RelationshipCollection._additem raises on duplicate rId"""
        # setup -----------------------
        part1 = BasePart()
        part2 = BasePart()
        rel1 = _Relationship('rId9', None, part1)
        rel2 = _Relationship('rId9', None, part2)
        self.relationships._additem(rel1)
        # verify ----------------------
        with self.assertRaises(ValueError):
            self.relationships._additem(rel2)

    def test__additem_maintains_rId_ordering(self):
        """_RelationshipCollection maintains rId ordering on additem()"""
        # setup -----------------------
        part1 = BasePart()
        part2 = BasePart()
        part3 = BasePart()
        rel1 = _Relationship('rId1', None, part1)
        rel2 = _Relationship('rId2', None, part2)
        rel3 = _Relationship('rId3', None, part3)
        # exercise --------------------
        self.relationships._additem(rel2)
        self.relationships._additem(rel1)
        self.relationships._additem(rel3)
        # verify ----------------------
        expected = ['rId1', 'rId2', 'rId3']
        actual = [rel._rId for rel in self.relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__additem_maintains_reltype_ordering(self):
        """_RelationshipCollection maintains reltype ordering on additem()"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        ordering = (RT_SLIDEMASTER, RT_SLIDELAYOUT, RT_SLIDE)
        relationships._reltype_ordering = ordering
        partname = '/ppt/slides/slide2.xml'
        part = Mock(name='new_part'); part.partname = partname
        rId = relationships._next_rId
        rel = _Relationship(rId, RT_SLIDE, part)
        # exercise --------------------
        relationships._additem(rel)
        # verify ordering -------------
        expected = [partnames[2], partnames[1], partnames[3],
                    partname, partnames[0], partnames[4]]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_rels_of_reltype_return_value(self):
        """RelationshipCollection._rels_of_reltype returns correct rels"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        # exercise --------------------
        retval = relationships.rels_of_reltype(RT_SLIDE)
        # verify ordering -------------
        expected = ['rId1', 'rId4']
        actual = [rel._rId for rel in retval]
        msg = "expected %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__reltype_ordering_sorts_rels(self):
        """RelationshipCollection._reltype_ordering sorts rels"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        ordering = (RT_SLIDEMASTER, RT_SLIDELAYOUT, RT_SLIDE)
        # exercise --------------------
        relationships._reltype_ordering = ordering
        # verify ordering -------------
        assert_that(relationships._reltype_ordering, is_(equal_to(ordering)))
        expected = [ partnames[2], partnames[1], partnames[3], partnames[0]
                   , partnames[4] ]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__reltype_ordering_renumbers_rels(self):
        """RelationshipCollection._reltype_ordering renumbers rels"""
        # setup -----------------------
        relationships, partnames = self.__reltype_ordering_mock()
        ordering = (RT_SLIDEMASTER, RT_SLIDELAYOUT, RT_SLIDE)
        # exercise --------------------
        relationships._reltype_ordering = ordering
        # verify renumbering ----------
        expected = ['rId1', 'rId2', 'rId3', 'rId4', 'rId5']
        actual = [rel._rId for rel in relationships]
        msg = "expected numbering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__next_rId_fills_gap(self):
        """_RelationshipCollection._next_rId fills gap in rId sequence"""
        # setup -----------------------
        part1 = BasePart()
        part2 = BasePart()
        part3 = BasePart()
        part4 = BasePart()
        rel1 = _Relationship('rId1', None, part1)
        rel2 = _Relationship('rId2', None, part2)
        rel3 = _Relationship('rId3', None, part3)
        rel4 = _Relationship('rId4', None, part4)
        cases =\
            ( ('rId1', (rel2, rel3, rel4))
            , ('rId2', (rel1, rel3, rel4))
            , ('rId3', (rel1, rel2, rel4))
            , ('rId4', (rel1, rel2, rel3))
            )
        # exercise --------------------
        expected_rIds = []
        actual_rIds = []
        for expected_rId, rels in cases:
            expected_rIds.append(expected_rId)
            relationships = _RelationshipCollection()
            for rel in rels:
                relationships._additem(rel)
            actual_rIds.append(relationships._next_rId)
        # verify ----------------------
        expected = expected_rIds
        actual = actual_rIds
        msg = "expected rIds %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_reorders_on_partname_change(self):
        """RelationshipCollection reorders on partname change"""
        # setup -----------------------
        partname1 = '/ppt/slides/slide1.xml'
        partname2 = '/ppt/slides/slide2.xml'
        partname3 = '/ppt/slides/slide3.xml'
        part1 = PartBuilder().with_partname(partname1).build()
        part2 = PartBuilder().with_partname(partname2).build()
        rel1 = _Relationship('rId1', RT_SLIDE, part1)
        rel2 = _Relationship('rId2', RT_SLIDE, part2)
        relationships = _RelationshipCollection()
        relationships._reltype_ordering = (RT_SLIDE)
        relationships._additem(rel1)
        relationships._additem(rel2)
        # exercise --------------------
        part1.partname = partname3
        # verify ----------------------
        expected = [partname2, partname3]
        actual = [rel._target.partname for rel in relationships]
        msg = "expected ordering %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestRun(TestCase):
    """Test Run"""
    def setUp(self):
        self.test_text = 'test text'
        self.r_xml = ('<a:r%s><a:t>%s</a:t></a:r>'
            % (nsprefix_decls, self.test_text))
        self.r = oxml_fromstring(self.r_xml)
        self.run = Run(self.r)

    def test_set_font_size(self):
        """Assignment to Run.font.size changes font size"""
        # setup -----------------------
        newfontsize = 2400
        expected_xml = ('<?xml version=\'1.0\' encoding=\'UTF-8\' standalone='
            '\'yes\'?>\n<a:r xmlns:a="http://schemas.openxmlformats.org/drawi'
            'ngml/2006/main">\n  <a:rPr sz="2400"/>\n  <a:t>test text</a:t>\n'
            '</a:r>\n')
        # exercise --------------------
        self.run.font.size = newfontsize
        # verify ----------------------
        r_xml = prettify_nsdecls(oxml_tostring(self.run._Run__r,
            encoding='UTF-8', pretty_print=True, standalone=True))
        r_xml_lines = r_xml.split('\n')
        expected_xml_lines = expected_xml.split('\n')
        for idx, line in enumerate(r_xml_lines):
            msg = ("\n\nexpected:\n\n%s\n\nbut got:\n\n%s"
                % (expected_xml, r_xml))
            self.assertEqual(line, expected_xml_lines[idx], msg)

    def test_text_value(self):
        """Run.text value is correct"""
        # exercise --------------------
        text = self.run.text
        # verify ----------------------
        expected = self.test_text
        actual = text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_text_setter(self):
        """Run.text setter stores passed value"""
        # setup -----------------------
        new_value = 'new string'
        # exercise --------------------
        self.run.text = new_value
        # verify ----------------------
        expected = new_value
        actual = self.run.text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__to_unicode_raises_on_non_string(self):
        """_to_unicode(text) raises on *text* not a string"""
        # verify ----------------------
        with self.assertRaises(TypeError):
            _to_unicode(999)


class TestShape(TestCase):
    """Test Shape"""
    def __loaded_shape(self):
        """
        Return Shape instance loaded from test file.
        """
        sldLayout = _slideLayout1()
        sp = sldLayout.xpath('p:cSld/p:spTree/p:sp', namespaces=nsmap)[0]
        return Shape(sp)


class TestShapeCollection(TestCase):
    """Test ShapeCollection"""
    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        sld = oxml_parse(path).getroot()
        spTree = sld.xpath('./p:cSld/p:spTree', namespaces=nsmap)[0]
        self.shapes = ShapeCollection(spTree)

    def test_construction_size(self):
        """ShapeCollection is expected size after construction"""
        # verify ----------------------
        self.assertLength(self.shapes, 9)

    @patch('pptx.presentation.Picture')
    @patch('pptx.presentation.Collection._values', new_callable=PropertyMock)
    @patch('pptx.presentation.Package')
    def test_add_picture_collaboration(self, MockPackage, mock_values, MockPicture):
        """ShapeCollection.add_picture() calls the right collaborators"""
        # constant values -------------
        rId = 'rId1'
        left = 1
        top = 2
        # setup mockery ---------------
        pkg      = Mock(name='pkg')
        image    = Mock(name='image')
        rel      = Mock(name='rel')
        pic      = Mock(name='pic')
        slide    = Mock(name='slide')
        __pic    = Mock(name='__pic')
        __spTree = Mock(name='__spTree')
        Picture  = MockPicture
        MockPackage.containing.return_value = pkg
        pkg._images.add_image.return_value = image
        slide._add_relationship.return_value = rel
        rel._rId = rId
        __pic.return_value = pic
        # setup -----------------------
        shapes = ShapeCollection(_empty_spTree(), slide)
        shapes._ShapeCollection__pic = __pic
        shapes._ShapeCollection__spTree = __spTree
        # exercise --------------------
        picture = shapes.add_picture(test_image_path, left, top)
        # verify ----------------------
        MockPackage.containing.assert_called_once_with(slide)
        pkg._images.add_image.assert_called_once_with(test_image_path)
        slide._add_relationship.assert_called_once_with(RT_IMAGE, image)
        __pic.assert_called_once_with(rId, test_image_path,
                                      left, top, None, None)
        __spTree.append.assert_called_once_with(pic)
        Picture.assert_called_once_with(pic)
        shapes._values.append.assert_called_once_with(picture)

    @patch('pptx.presentation.Collection._values', new_callable=PropertyMock)
    @patch('pptx.presentation.Shape')
    @patch('pptx.presentation.ShapeCollection._ShapeCollection__next_shape_id'
          , new_callable=PropertyMock)
    def test_add_textbox_collaboration(self, __next_shape_id, Shape, _values):
        """ShapeCollection.add_textbox() calls the right collaborators"""
        # constant values -------------
        sp_id  = 9
        name   = 'TextBox 8'
        text   = 'Test text'
        left   = Inches(1.0)
        top    = Inches(2.0)
        width  = Inches(1.5)
        height = Inches(0.5)
        # setup mockery ---------------
        __next_shape_id.return_value = sp_id
        sp = Mock(name='sp')
        __sp = Mock(name='__sp', return_value=sp)
        __spTree = Mock(name='__spTree')
        shapes = ShapeCollection(_empty_spTree())
        shapes._ShapeCollection__sp = __sp
        shapes._ShapeCollection__spTree = __spTree
        # exercise --------------------
        shape = shapes.add_textbox(left, top, width, height)
        # verify ----------------------
        __next_shape_id.assert_called_once_with()
        __sp.assert_called_once_with(sp_id, name, left, top,
                                     width, height, is_textbox=True)
        __spTree.append.assert_called_once_with(sp)
        Shape.assert_called_once_with(sp)
        shapes._values.append.assert_called_once_with(shape)

    def test_add_textbox_xml(self):
        """ShapeCollection.add_textbox() generates correct XML"""
        # constant values -------------
        left   = Inches(1.0)
        top    = Inches(2.0)
        width  = Inches(1.5)
        height = Inches(0.5)
        shapes = ShapeCollection(_empty_spTree())
        # exercise --------------------
        shape = shapes.add_textbox(left, top, width, height)
        # verify ----------------------
        xml = oxml_tostring(shape._element, encoding='UTF-8',
                            pretty_print=True, standalone=True)
        xml = prettify_nsdecls(xml)
        xml_lines = xml.split('\n')
        txbox_xml_lines = _txbox_xml().split('\n')
        for idx, line in enumerate(xml_lines):
            msg = "expected:\n%s\n\nbut got:\n\n%s" % (xml, _txbox_xml())
            self.assertEqual(line, txbox_xml_lines[idx], msg)

    def test_title_value(self):
        """ShapeCollection.title value is ref to correct shape"""
        # exercise --------------------
        title_shape = self.shapes.title
        # verify ----------------------
        expected = 0
        actual = self.shapes.index(title_shape)
        msg = "expected shapes[%d], got shapes[%d]" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_placeholders_values(self):
        """ShapeCollection.placeholders values are correct and sorted"""
        # setup -----------------------
        expected_values =\
            ( ('Title 1',                    PH_TYPE_CTRTITLE,  0)
            , ('Vertical Subtitle 2',        PH_TYPE_SUBTITLE,  1)
            , ('Date Placeholder 7',         PH_TYPE_DT,       10)
            , ('Footer Placeholder 4',       PH_TYPE_FTR,      11)
            , ('Slide Number Placeholder 5', PH_TYPE_SLDNUM,   12)
            , ('Table Placeholder 3',        PH_TYPE_TBL,      14)
            )
        shapes = _sldLayout1_shapes()
        # exercise --------------------
        placeholders = shapes.placeholders
        # verify ----------------------
        for idx, ph in enumerate(placeholders):
            values = (ph.name, ph.type, ph.idx)
            expected = expected_values[idx]
            actual = values
            msg = "expected placeholders[%d] values %s, got %s"\
                   % (idx, expected, actual)
            self.assertEqual(expected, actual, msg)

    def test__clone_layout_placeholders_shapes(self):
        """ShapeCollection._clone_layout_placeholders clones shapes"""
        # setup -----------------------
        expected_values =\
            ( [2, 'Title 1',                    PH_TYPE_CTRTITLE,  0]
            , [3, 'Vertical Subtitle 2',        PH_TYPE_SUBTITLE,  1]
            , [4, 'Table Placeholder 3',        PH_TYPE_TBL,      14]
            )
        slidelayout = SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        shapes = ShapeCollection(_empty_spTree())
        # exercise --------------------
        shapes._clone_layout_placeholders(slidelayout)
        # verify ----------------------
        for idx, sp in enumerate(shapes):
            # verify is placeholder ---
            is_placeholder = sp.is_placeholder
            msg = ("expected shapes[%d].is_placeholder == True %r"
                   % (idx, sp))
            self.assertTrue(is_placeholder, msg)
            # verify values -----------
            ph = Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = ("expected placeholder[%d] values %s, got %s"
                   % (idx, expected, actual))
            self.assertEqual(expected, actual, msg)

    def test___clone_layout_placeholder_values(self):
        """ShapeCollection.__clone_layout_placeholder() values correct"""
        # setup -----------------------
        layout_shapes = _sldLayout1_shapes()
        layout_ph_shapes = [sp for sp in layout_shapes if sp.is_placeholder]
        shapes = ShapeCollection(_empty_spTree())
        expected_values = (
            [2, 'Title 1',                    PH_TYPE_CTRTITLE,  0],
            [3, 'Date Placeholder 2',         PH_TYPE_DT,       10],
            [4, 'Vertical Subtitle 3',        PH_TYPE_SUBTITLE,  1],
            [5, 'Table Placeholder 4',        PH_TYPE_TBL,      14],
            [6, 'Slide Number Placeholder 5', PH_TYPE_SLDNUM,   12],
            [7, 'Footer Placeholder 6',       PH_TYPE_FTR,      11])
        # exercise --------------------
        for idx, layout_ph_sp in enumerate(layout_ph_shapes):
            layout_ph = Placeholder(layout_ph_sp)
            sp = shapes._ShapeCollection__clone_layout_placeholder(layout_ph)
            # verify ------------------
            ph = Placeholder(sp)
            expected = expected_values[idx]
            actual = [ph.id, ph.name, ph.type, ph.idx]
            msg = "expected placeholder values %s, got %s" % (expected, actual)
            self.assertEqual(expected, actual, msg)

    def test___clone_layout_placeholder_xml(self):
        """ShapeCollection.__clone_layout_placeholder() produces correct XML"""
        # setup -----------------------
        layout_shapes = _sldLayout1_shapes()
        layout_ph_shapes = [sp for sp in layout_shapes if sp.is_placeholder]
        shapes = ShapeCollection(_empty_spTree())
        xml_template = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone='
            '\'yes\'?>\n<p:sp xmlns:p="http://schemas.openxmlformats.org/pres'
            'entationml/2006/main"\n      xmlns:a="http://schemas.openxmlform'
            'ats.org/drawingml/2006/main">\n  <p:nvSpPr>\n    <p:cNvPr id="%d'
            '" name="%s"/>\n    <p:cNvSpPr>\n      <a:spLocks noGrp="1"/>\n  '
            '  </p:cNvSpPr>\n    <p:nvPr>\n      <p:ph type="%s"%s/>\n    </p'
            ':nvPr>\n  </p:nvSpPr>\n  <p:spPr/>\n%s</p:sp>')
        txBody_snippet = (
            '  <p:txBody>\n    <a:bodyPr/>\n    <a:lstStyle/>\n    <a:p/>\n  '
            '</p:txBody>\n')
        expected_values = [
            (2, 'Title 1', PH_TYPE_CTRTITLE, '', txBody_snippet),
            (3, 'Date Placeholder 2', PH_TYPE_DT, ' sz="half" idx="10"', ''),
            (4, 'Vertical Subtitle 3', PH_TYPE_SUBTITLE,
                ' orient="vert" idx="1"', txBody_snippet),
            (5, 'Table Placeholder 4', PH_TYPE_TBL,
                ' sz="quarter" idx="14"', ''),
            (6, 'Slide Number Placeholder 5', PH_TYPE_SLDNUM,
                ' sz="quarter" idx="12"', ''),
            (7, 'Footer Placeholder 6', PH_TYPE_FTR,
                ' sz="quarter" idx="11"', '')]
                    # verify ----------------------
        for idx, layout_ph_sp in enumerate(layout_ph_shapes):
            # log.debug("layout_ph_sp.name '%s'" % layout_ph_sp.name)
            layout_ph = Placeholder(layout_ph_sp)
            sp = shapes._ShapeCollection__clone_layout_placeholder(layout_ph)
            ph = Placeholder(sp)
            sp_xml = prettify_nsdecls(
                oxml_tostring(ph._element, encoding='UTF-8',
                              pretty_print=True, standalone=True))
            sp_xml_lines = sp_xml.split('\n')
            expected_xml = xml_template % expected_values[idx]
            expected_xml_lines = expected_xml.split('\n')
            for idx, line in enumerate(sp_xml_lines):
                msg = '\n\n%s' % sp_xml
                self.assertEqual(line, expected_xml_lines[idx], msg)
                # assert_that(line, is_(equal_to(expected_xml_lines[idx])))

    def test___next_ph_name_return_value(self):
        """
        ShapeCollection.__next_ph_name() returns correct value

        * basename + 'Placeholder' + num, e.g. 'Table Placeholder 8'
        * numpart of name defaults to id-1, but increments until unique
        * prefix 'Vertical' if orient="vert"

        """
        cases = (
            (PH_TYPE_OBJ,   3, PH_ORIENT_HORZ, 'Content Placeholder 2'),
            (PH_TYPE_TBL,   4, PH_ORIENT_HORZ, 'Table Placeholder 4'),
            (PH_TYPE_TBL,   7, PH_ORIENT_VERT, 'Vertical Table Placeholder 6'),
            (PH_TYPE_TITLE, 2, PH_ORIENT_HORZ, 'Title 2'))
        # setup -----------------------
        shapes = _sldLayout1_shapes()
        for ph_type, id, orient, expected_name in cases:
            # exercise --------------------
            name = shapes._ShapeCollection__next_ph_name(ph_type, id, orient)
            # verify ----------------------
            expected = expected_name
            actual = name
            msg = "expected placeholder name '%s', got '%s'"\
                  % (expected, actual)
            self.assertEqual(expected, actual, msg)

    def test___next_shape_id_value(self):
        """ShapeCollection.__next_shape_id value is correct"""
        # setup -----------------------
        shapes = _sldLayout1_shapes()
        # exercise --------------------
        next_id = shapes._ShapeCollection__next_shape_id
        # verify ----------------------
        expected = 4
        actual = next_id
        msg = "expected %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test___pic_generates_correct_xml(self):
        """ShapeCollection.__pic returns correct value"""
        # setup -----------------------
        test_image = PILImage.open(test_image_path)
        pic_size = tuple(Px(x) for x in test_image.size)
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>'
            '\n<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/20'
            '06/main"\n       xmlns:p="http://schemas.openxmlformats.org/pres'
            'entationml/2006/main"\n       xmlns:r="http://schemas.openxmlfor'
            'mats.org/officeDocument/2006/relationships">\n  <p:nvPicPr>\n   '
            ' <p:cNvPr id="4" name="Picture 3" descr="python-icon.jpeg"/>\n  '
            '  <p:cNvPicPr/>\n    <p:nvPr/>\n  </p:nvPicPr>\n  <p:blipFill>\n'
            '    <a:blip r:embed="rId9"/>\n    <a:stretch>\n      <a:fillRect'
            '/>\n    </a:stretch>\n  </p:blipFill>\n  <p:spPr>\n    <a:xfrm>'
            '\n      <a:off x="0" y="0"/>\n      <a:ext cx="%s" cy="%s"/>\n  '
            '  </a:xfrm>\n    <a:prstGeom prst="rect">\n      <a:avLst/>\n   '
            ' </a:prstGeom>\n  </p:spPr>\n</p:pic>' % pic_size)
        # exercise --------------------
        pic = self.shapes._ShapeCollection__pic('rId9', test_image_path, 0, 0)
        # verify ----------------------
        pic_xml = oxml_tostring(pic, encoding='UTF-8', pretty_print=True,
                                standalone=True)
        pic_xml = prettify_nsdecls(pic_xml)
        pic_xml_lines = pic_xml.split('\n')
        expected_xml_lines = xml.split('\n')
        for idx, line in enumerate(pic_xml_lines):
            msg = "\n\nexpected:\n\n%s\n\nbut got\n\n%s" % (xml, pic_xml)
            self.assertEqual(line, expected_xml_lines[idx], msg)
            # assert_that(line, is_(equal_to(expected_xml_lines[idx])))

    def test___pic_from_stream_generates_correct_xml(self):
        """ShapeCollection.__pic returns correct XML from stream image"""
        # setup -----------------------
        test_image = PILImage.open(test_image_path)
        pic_size = tuple(Px(x) for x in test_image.size)
        xml = (
            '<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2'
            '006/main" xmlns:p="http://schemas.openxmlformats.org/presentatio'
            'nml/2006/main" xmlns:r="http://schemas.openxmlformats.org/office'
            'Document/2006/relationships"><p:nvPicPr><p:cNvPr id="4" name="Pi'
            'cture 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip '
            'r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></p:blipFill'
            '><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%s" cy="%s"/></a'
            ':xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></'
            'p:pic>' % pic_size)
        # exercise --------------------
        with open(test_image_path) as stream:
            pic = self.shapes._ShapeCollection__pic('rId9', stream, 0, 0)
        # verify ----------------------
        assert_that(oxml_tostring(pic), is_(equal_to(xml)))


class TestSlide(TestCase):
    """Test Slide"""
    def setUp(self):
        self.sld = Slide()

    def test_constructor_sets_correct_content_type(self):
        """Slide constructor sets correct content type"""
        # exercise --------------------
        content_type = self.sld._content_type
        # verify ----------------------
        expected = CT_SLIDE
        actual = content_type
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_construction_adds_slide_layout_relationship(self):
        """Slide(slidelayout) adds relationship slide->slidelayout"""
        # setup -----------------------
        slidelayout = SlideLayout()
        slidelayout._shapes = _sldLayout1_shapes()
        # exercise --------------------
        slide = Slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(slide._relationships)
        msg = ("expected len(slide._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = slide._relationships[0]
        expected = ('rId1', RT_SLIDELAYOUT, slidelayout)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = "expected relationship\n%s\ngot\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test__element_minimal_sld_on_construction(self):
        """Slide._element is minimal sld on construction"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/minimal_slide.xml')
        # exercise --------------------
        elm = self.sld._element
        # verify ----------------------
        with open(path, 'r') as f:
            expected = f.read()
        actual = prettify_nsdecls(
            oxml_tostring(elm, encoding='UTF-8', pretty_print=True,
                          standalone=True))
        msg = "\nexpected:\n\n'%s'\n\nbut got:\n\n'%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_slidelayout_property_none_on_construction(self):
        """Slide.slidelayout property None on construction"""
        # verify ----------------------
        self.assertIsProperty(self.sld, 'slidelayout', None)

    def test__load_sets_slidelayout(self):
        """Slide._load() sets slidelayout"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        slidelayout = Mock(name='slideLayout')
        slidelayout.partname = '/ppt/slideLayouts/slideLayout1.xml'
        rel = Mock(name='pptx.packaging._Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT_SLIDELAYOUT
        rel.target = slidelayout
        pkgpart = Mock(name='pptx.packaging.Part')
        with open(path, 'rb') as f:
            pkgpart.blob = f.read()
        pkgpart.relationships = [rel]
        part_dict = {slidelayout.partname: slidelayout}
        slide = self.sld._load(pkgpart, part_dict)
        # exercise --------------------
        retval = slide.slidelayout
        # verify ----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test___minimal_element_xml(self):
        """Slide.__minimal_element generates correct XML"""
        # setup -----------------------
        path = os.path.join(thisdir, 'test_files/minimal_slide.xml')
        # exercise --------------------
        sld = self.sld._Slide__minimal_element
        # verify ----------------------
        with open(path, 'r') as f:
            expected_xml = f.read()
        sld_xml = prettify_nsdecls(
            oxml_tostring(sld, encoding='UTF-8', pretty_print=True,
                          standalone=True))
        sld_xml_lines = sld_xml.split('\n')
        expected_xml_lines = expected_xml.split('\n')
        for idx, line in enumerate(sld_xml_lines):
            # msg = '\n\n%s' % sld_xml
            msg = "expected:\n%s\n, got\n%s" % (expected_xml, sld_xml)
            self.assertEqual(line, expected_xml_lines[idx], msg)


class TestSlideCollection(TestCase):
    """Test SlideCollection"""
    def setUp(self):
        prs = Presentation()
        self.slides = SlideCollection(prs)

    def test_add_slide_returns_slide(self):
        """SlideCollection.add_slide() returns instance of Slide"""
        # exercise --------------------
        retval = self.slides.add_slide(None)
        # verify ----------------------
        self.assertIsInstance(retval, Slide)

    def test_add_slide_sets_slidelayout(self):
        """
        SlideCollection.add_slide() sets Slide.slidelayout

        Kind of a throw-away test, but was helpful for initial debugging.
        """
        # setup -----------------------
        slidelayout = Mock(name='slideLayout')
        slidelayout.shapes = []
        slide = self.slides.add_slide(slidelayout)
        # exercise --------------------
        retval = slide.slidelayout
        # verify ----------------------
        expected = slidelayout
        actual = retval
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_slide_adds_slide_layout_relationship(self):
        """SlideCollection.add_slide() adds relationship prs->slide"""
        # setup -----------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = SlideLayout()
        slidelayout._shapes = []
        # exercise --------------------
        slide = slides.add_slide(slidelayout)
        # verify length ---------------
        expected = 1
        actual = len(prs._relationships)
        msg = ("expected len(prs._relationships) of %d, got %d"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)
        # verify values ---------------
        rel = prs._relationships[0]
        expected = ('rId1', RT_SLIDE, slide)
        actual = (rel._rId, rel._reltype, rel._target)
        msg = ("expected relationship 1:, got 2:\n1: %s\n2: %s"
               % (expected, actual))
        self.assertEqual(expected, actual, msg)

    def test_add_slide_sets_partname(self):
        """SlideCollection.add_slide() sets partname of new slide"""
        # setup -----------------------
        prs = Presentation()
        slides = prs.slides
        slidelayout = SlideLayout()
        slidelayout._shapes = []
        # exercise --------------------
        slide = slides.add_slide(slidelayout)
        # verify ----------------------
        expected = '/ppt/slides/slide1.xml'
        actual = slide.partname
        msg = "expected partname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class TestSlideLayout(TestCase):
    """Test SlideLayout"""
    def setUp(self):
        self.slidelayout = SlideLayout()

    def __loaded_slidelayout(self, prs_slidemaster=None):
        """
        Return SlideLayout instance loaded using mocks. *prs_slidemaster* is
        an already-loaded model-side SlideMaster instance (or mock, as
        appropriate to calling test).
        """
        # partname for related slideMaster
        sldmaster_partname = '/ppt/slideMasters/slideMaster1.xml'
        # path to test slideLayout XML
        slidelayout_path = absjoin(test_file_dir, 'slideLayout1.xml')
        # model-side slideMaster part
        if prs_slidemaster is None:
            prs_slidemaster = Mock(spec=SlideMaster)
        # a part dict containing the already-loaded model-side slideMaster
        loaded_part_dict = {sldmaster_partname: prs_slidemaster}
        # a slideMaster package part for rel target
        pkg_slidemaster_part = Mock(spec=pptx.packaging.Part)
        pkg_slidemaster_part.partname = sldmaster_partname
        # a package-side relationship from slideLayout to its slideMaster
        rel = Mock(name='pptx.packaging._Relationship')
        rel.rId = 'rId1'
        rel.reltype = RT_SLIDEMASTER
        rel.target = pkg_slidemaster_part
        # the slideLayout package part to send to _load()
        pkg_slidelayout_part = Mock(spec=pptx.packaging.Part)
        pkg_slidelayout_part.relationships = [rel]
        with open(slidelayout_path, 'rb') as f:
            pkg_slidelayout_part.blob = f.read()
        # _load and return
        slidelayout = SlideLayout()
        return slidelayout._load(pkg_slidelayout_part, loaded_part_dict)

    def test__load_sets_slidemaster(self):
        """SlideLayout._load() sets slidemaster"""
        # setup -----------------------
        prs_slidemaster = Mock(spec=SlideMaster)
        # exercise --------------------
        loaded_slidelayout = self.__loaded_slidelayout(prs_slidemaster)
        # verify ----------------------
        expected = prs_slidemaster
        actual = loaded_slidelayout.slidemaster
        msg = "expected: %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_slidemaster_is_readonly(self):
        """SlideLayout.slidemaster is read-only"""
        # verify ----------------------
        self.assertIsReadOnly(self.slidelayout, 'slidemaster')

    def test_slidemaster_raises_on_ref_before_assigned(self):
        """SlideLayout.slidemaster raises on referenced before assigned"""
        with self.assertRaises(AssertionError):
            self.slidelayout.slidemaster


class TestSlideMaster(TestCase):
    """Test SlideMaster"""
    def setUp(self):
        self.sldmaster = SlideMaster()

    def test_slidelayouts_property_empty_on_construction(self):
        """SlideMaster.slidelayouts property empty on construction"""
        # verify ----------------------
        self.assertIsSizedProperty(self.sldmaster, 'slidelayouts', 0)

    def test_slidelayouts_correct_length_after_open(self):
        """SlideMaster.slidelayouts correct length after open"""
        # setup -----------------------
        pkg = Package(test_pptx_path)
        slidemaster = pkg.presentation.slidemasters[0]
        # exercise --------------------
        slidelayouts = slidemaster.slidelayouts
        # verify ----------------------
        self.assertLength(slidelayouts, 11)


class TestTextFrame(TestCase):
    """Test TextFrame"""
    def setUp(self):
        path = os.path.join(thisdir, 'test_files/slide1.xml')
        self.sld = oxml_parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody'
        self.txBodyList = self.sld.xpath(xpath, namespaces=nsmap)

    def test_paragraphs_size(self):
        """TextFrame.paragraphs is expected size"""
        # setup -----------------------
        actual_lengths = []
        for txBody in self.txBodyList:
            textframe = TextFrame(txBody)
            # exercise ----------------
            actual_lengths.append(len(textframe.paragraphs))
        # verify ------------------
        expected = [1, 1, 2, 1, 1]
        actual = actual_lengths
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_paragraph_xml(self):
        """TextFrame.add_paragraph does what it says"""
        # setup -----------------------
        txBody_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentatio'
            'nml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawin'
            'gml/2006/main" xmlns:r="http://schemas.openxmlformats.org/office'
            'Document/2006/relationships"><a:bodyPr/><a:p><a:r><a:t>Test text'
            '</a:t></a:r></a:p></p:txBody>')
        expected_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentatio'
            'nml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawin'
            'gml/2006/main"><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></'
            'a:p><a:p/></p:txBody>')
        txBody = oxml_fromstring(txBody_xml)
        textframe = TextFrame(txBody)
        # exercise --------------------
        textframe.add_paragraph()
        # verify ----------------------
        assert_that(len(textframe.paragraphs), is_(equal_to(2)))
        textframe_xml = oxml_tostring(textframe._TextFrame__txBody)
        expected = expected_xml
        actual = textframe_xml
        msg = "\nExpected: '%s'\n\n     Got: '%s'" % (expected, actual)
        if not expected == actual:
            raise AssertionError(msg)

    def test_text_setter_structure_and_value(self):
        """assign to TextFrame.text yields single run para set to value"""
        # setup -----------------------
        test_text = 'python-pptx was here!!'
        txBody = self.txBodyList[2]
        textframe = TextFrame(txBody)
        # exercise --------------------
        textframe.text = test_text
        # verify paragraph count ------
        expected = 1
        actual = len(textframe.paragraphs)
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value ----------------
        expected = test_text
        actual = textframe.paragraphs[0].runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_vertical_anchor_works(self):
        """Assignment to TextFrame.vertical_anchor sets vert anchor"""
        # setup -----------------------
        txBody_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentatio'
            'nml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawin'
            'gml/2006/main"><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></'
            'a:p></p:txBody>')
        expected_xml = (
            '<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentatio'
            'nml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawin'
            'gml/2006/main">\n  <a:bodyPr anchor="ctr"/>\n  <a:p>\n    <a:r>'
            '\n      <a:t>Test text</a:t>\n    </a:r>\n  </a:p>\n</p:txBody>'
            '\n')
        txBody = oxml_fromstring(txBody_xml)
        textframe = TextFrame(txBody)
        # exercise --------------------
        textframe.vertical_anchor = MSO.ANCHOR_MIDDLE
        # verify ----------------------
        expected_xml_lines = expected_xml.split('\n')
        actual_xml = oxml_tostring(textframe._TextFrame__txBody,
                                   pretty_print=True)
        actual_xml_lines = actual_xml.split('\n')
        for idx, actual_line in enumerate(actual_xml_lines):
            expected_line = expected_xml_lines[idx]
            msg = ("\n\nexpected:\n\n'%s'\n\nbut got:\n\n'%s'"
                   % (expected_xml, actual_xml))
            assert_that(actual_line, is_(equal_to(expected_line)), msg)

