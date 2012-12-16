# -*- coding: utf-8 -*-
#
# test.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.packaging module."""

import inspect
import os
import unittest
import zipfile

from lxml import etree
from StringIO import StringIO

import pptx.packaging

from pptx.exceptions import CorruptedPackageError, DuplicateKeyError,\
                            NotXMLError

from testing import TestCase

# ============================================================================
# Utility Items
# ============================================================================

basepath = os.path.split(__file__)[0]
template_dirpath = os.path.abspath(os.path.join(basepath, '../pptx/pptx_template'))
dir_pkg_path = os.path.join(template_dirpath, 'default')
zip_pkg_path = os.path.join(template_dirpath, 'default.pptx')

class MockParent(object):
    """Stub out parent attributes."""
    def __init__(self, baseURI=None, itemURI=None):
        self.baseURI = baseURI
        self.itemURI = itemURI
    


# ============================================================================
# Test Classes
# ============================================================================

class TestContentTypesItem(TestCase):
    """
    Test pptx.packaging.ContentTypesItem
    
    """
    def setUp(self):
        # construct ContentTypesItem instance
        self.cti = pptx.packaging.ContentTypesItem()
    
    def test_class_present(self):
        """
        ContentTypesItem class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'ContentTypesItem')
    
    def test_compose_method_present(self):
        """
        ContentTypesItem class has method 'compose'.
        
        NOTE: This test will fail if method throws an exception.
        
        """
        self.assertClassHasMethod(pptx.packaging.ContentTypesItem, 'compose')
    
    def test_compose_returns_self(self):
        """
        ContentTypesItem.compose() returns self-reference.
        
        """
        # setup -----------------------
        pkg = pptx.packaging.Package().open(zip_pkg_path)
        # exercise --------------------
        retval = self.cti.compose(pkg.parts)
        # verify ----------------------
        expected = self.cti
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_compose_correct_count(self):
        """
        ContentTypesItem.compose() produces expected element count.
        
        """
        # setup -----------------------
        pkg = pptx.packaging.Package().open(zip_pkg_path)
        # exercise --------------------
        self.cti.compose(pkg.parts)
        # verify ----------------------
        expected = 23
        actual = len(self.cti)
        msg = "expected %d elements, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_element_attribute_present(self):
        """
        ContentTypesItem instance has attribute 'element'.
        
        """
        self.assertInstHasAttr(self.cti, 'element')
    
    def test_element_correct_length(self):
        """
        ContentTypesItem.element() has expected element count.
        
        """
        # setup -----------------------
        pkg = pptx.packaging.Package().open(zip_pkg_path)
        # exercise --------------------
        self.cti.compose(pkg.parts)
        # verify ----------------------
        expected = 23
        actual = len(self.cti.element)
        msg = "expected %d elements, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_getitem_method_present(self):
        """
        ContentTypesItem class has method '__getitem__'.
        
        """
        self.assertClassHasMethod(pptx.packaging.ContentTypesItem,
                                  '__getitem__')
    
    def test_getitem_throws_on_bad_partname(self):
        """
        ContentTypesItem.__getitem__() throws error on bad partname.
        
        """
        # verify ----------------------
        with self.assertRaises(LookupError):
            self.cti['!blat/rhumba.xml']
    
    def test_load_works(self):
        """
        ContentTypesItem can load itself from a filesystem instance.
        
        """
        # setup ------------------------
        dir_fs = pptx.packaging.FileSystem(dir_pkg_path)
        zip_fs = pptx.packaging.FileSystem(zip_pkg_path)
        for fs in (dir_fs, zip_fs):
            # exercise ---------------------
            cti = pptx.packaging.ContentTypesItem().load(fs)
            # test -------------------------
            expected = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
            actual = cti['/ppt/slideLayouts/slideLayout1.xml']
            msg = "expected content type '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestFileSystem(TestCase):
    """
    Test pptx.packaging.FileSystem
    
    """
    def setUp(self):
        self.dir_fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.zip_fs = pptx.packaging.FileSystem(zip_pkg_path)
    
    def test_class_present(self):
        """
        FileSystem class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'FileSystem')
    
    def test_construction(self):
        """
        FileSystem factory returns package-appropriate class.
        
        """
        test_datasets =\
            ( (self.dir_fs, 'DirectoryFileSystem')
            , (self.zip_fs, 'ZipFileSystem')
            )
        for fs, expected in test_datasets:
            actual = fs.__class__.__name__
            msg = "expected class name '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_contains(self):
        """
        'in' operator returns True if URI is in filesystem.
        
        """
        expected_URIs =\
            ( '/[Content_Types].xml'
            , '/docProps/app.xml'
            , '/ppt/presentation.xml'
            , '/ppt/slideMasters/slideMaster1.xml'
            , '/ppt/slideLayouts/_rels/slideLayout1.xml.rels'
            )
        for fs in (self.dir_fs, self.zip_fs):
            for uri in expected_URIs:
                self.assertIn(uri, fs)
    
    def test_getelement_return_count(self):
        """
        ElementTree element for specified package item is returned.
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            elm = fs.getelement('/[Content_Types].xml')
            expected = 23
            actual = len(elm)
            msg = "expected %d elements, got %d" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_getelement_throws_on_binary(self):
        """
        Calling getelement() for binary item throws exception.
        
        """
        # call getelement for thumbnail
        for fs in (self.dir_fs, self.zip_fs):
            with self.assertRaises(NotXMLError):
                fs.getelement('/docProps/thumbnail.jpeg')
    
    def test_getstream_correct_length(self):
        """
        StringIO instance for specified package item is returned.
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            stream = fs.getstream('/[Content_Types].xml')
            elm = etree.parse(stream).getroot()
            expected = 23
            actual = len(elm)
            msg = "expected %d elements, got %d" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_getstream_throws_on_bad_URI(self):
        """
        FileSystem.getstream() throws on bad URI.
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            with self.assertRaises(LookupError):
                fs.getstream('!blat/rhumba.xml')
    
    def test_itemURIs_count(self):
        """
        FileSystem.itemURIs has expected count.
        
        """
        # verify ----------------------
        for fs, fsname in ((self.dir_fs, 'dir_fs'), (self.zip_fs, 'zip_fs')):
            expected = 37
            actual = len(fs.itemURIs)
            msg = "expected %d members in %s, got %d" % (expected, fsname, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_itemURIs_plausible(self):
        """
        All URIs in FileSystem.itemURIs are plausible.
        
        """
        # verify ----------------------
        for fs in (self.dir_fs, self.zip_fs):
            for itemURI in fs.itemURIs:
                # plausible segment count
                expected_min = 1
                expected_max = 4
                segment_count = len(itemURI.split('/'))-1  # leading slash produces empty string in split list
                msg = "item URI has implausible number of segments:\nitemURI ==> '%s'" % (itemURI)
                self.assertTrue(segment_count >= expected_min, msg)
                self.assertTrue(segment_count <= expected_max, msg)
                # check for forward slash
                msg = "item URI '%s' does not start with forward slash ('/')" % (itemURI)
                self.assertTrue(itemURI.startswith('/'), msg)
    

class TestPackage(TestCase):
    """
    Test pptx.packaging.Package
    
    """
    def setUp(self):
        # construct Package
        self.pkg = pptx.packaging.Package()
        self.test_pptx_path = os.path.join(basepath, 'test_python-pptx.pptx')
    
    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)
    
    def test_class_present(self):
        """
        Package class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'Package')
    
    def test_loadparts_part_count(self):
        """
        Package.loadparts() produces expected part count.
        
        """
        # setup -----------------------
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        cti = pptx.packaging.ContentTypesItem().load(fs)
        pri = pptx.packaging.PackageRelationshipsItem().load(fs)
        # exercise --------------------
        self.pkg.loadparts(fs, cti, pri.relationships)
        # verify ----------------------
        expected = 21
        actual = len(self.pkg.parts)
        msg = "expected part count %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_open_returns_package(self):
        """
        Package.open() returns object of class 'Package'.
        
        """
        for path in (dir_pkg_path, zip_pkg_path):
            pkg = pptx.packaging.Package().open(path)
            expected = 'Package'
            actual = pkg.__class__.__name__
            msg = "expected class name '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_open_part_count(self):
        """
        Package.open() produces expected part count.
        
        """
        # exercise --------------------
        self.pkg.open(zip_pkg_path)
        # verify ----------------------
        expected = 21
        actual = len(self.pkg.parts)
        msg = "expected part count %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_parts_attribute_present(self):
        """
        Package instance has attribute 'parts'.
        
        """
        self.assertInstHasAttr(self.pkg, 'parts')
    
    def test_parts_empty_on_construction(self):
        """
        Package.parts is empty on construction.
        
        """
        expected = 0
        actual = len(self.pkg.parts)
        msg = "expected %d parts, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_parts_is_readonly(self):
        """
        Package.parts is read-only.
        
        """
        with self.assertRaises(AttributeError):
            self.pkg.parts = None
    
    def test_save_method_present(self):
        """
        Package class has method 'save'.
        
        """
        self.assertClassHasMethod(pptx.packaging.Package, 'save')
    
    def test_save_writes_pptx_zipfile(self):
        """
        Package.save() writes .pptx file.
        
        """
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_path = self.test_pptx_path
        # exercise --------------------
        self.pkg.save(save_path)
        # verify ----------------------
        actual = zipfile.is_zipfile(save_path)
        msg = "no zipfile at %s" % (save_path)
        self.assertTrue(actual, msg)
    
    def test_save_adds_pptx_ext(self):
        """
        Package.save() adds .pptx extension if not supplied.
        
        """
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_path = self.test_pptx_path[:-5]  # trim off .pptx extension
        exp_path = self.test_pptx_path
        # exercise --------------------
        self.pkg.save(save_path)
        # verify ----------------------
        actual = zipfile.is_zipfile(exp_path)
        msg = "no zipfile at %s" % (exp_path)
        self.assertTrue(actual, msg)
    
    def test_save_member_count(self):
        """
        Package.save() produces expected zip member count.
        
        """
        # setup -----------------------
        self.pkg.open(dir_pkg_path)
        save_path = self.test_pptx_path
        # exercise --------------------
        self.pkg.save(save_path)
        # verify ----------------------
        zip = zipfile.ZipFile(save_path)
        names = zip.namelist()
        zip.close()
        partnames = [name for name in names if not name.endswith('/')]
        expected = 37
        actual = len(partnames)
        msg = "expected %d parts, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPackageRelationshipsItem(TestCase):
    """
    Test pptx.packaging.PackageRelationshipsItem
    
    """
    def setUp(self):
        # load a filesystem of each kind
        self.dir_fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.zip_fs = pptx.packaging.FileSystem(zip_pkg_path)
        # create an empty PackageRelationshipsItem instance
        self.pri = pptx.packaging.PackageRelationshipsItem()
    
    def tearDown(self):
        """
        No tear down required.
        
        """
        pass
    
    def test_class_present(self):
        """
        PackageRelationshipsItem class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'PackageRelationshipsItem')
    
    def test_baseuri_attribute_present(self):
        """
        PackageRelationshipsItem instance has attribute 'baseURI'.
        
        """
        self.assertInstHasAttr(self.pri, 'baseURI')
    
    def test_baseuri_is_readonly(self):
        """
        Assignment to PackageRelationshipsItem.baseURI throws AttributeError.
        
        """
        with self.assertRaises(AttributeError):
            self.pri.baseURI = None
    
    def test_baseuri_correct(self):
        """
        PackageRelationshipsItem.baseURI returns correct value.
        
        """
        expected = '/'
        actual = self.pri.baseURI
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_construction(self):
        """
        PackageRelationshipsItem can be constructed.
        
        """
        expected = 'PackageRelationshipsItem'
        actual = self.pri.__class__.__name__
        msg = "expected classname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_itemuri_attribute_present(self):
        """
        PackageRelationshipsItem instance has attribute 'itemURI'.
        
        """
        self.assertInstHasAttr(self.pri, 'itemURI')
    
    def test_itemuri_is_readonly(self):
        """
        Assignment to PackageRelationshipsItem.itemURI throws AttributeError.
        
        """
        with self.assertRaises(AttributeError):
            self.pri.itemURI = None
    
    def test_itemuri_correct(self):
        """
        PackageRelationshipsItem.itemURI returns correct value.
        
        """
        expected = '/_rels/.rels'
        actual = self.pri.itemURI
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_returns_self(self):
        """
        PackageRelationshipsItem.load() returns reference to caller.
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            expected = self.pri
            actual = self.pri.load(fs)
            msg = "expected PackageRelationshipsItem '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_load_rels_count(self):
        """
        PackageRelationshipItem loads expected number of relationships.
        
        """
        # setup ------------------------
        for fs in (self.dir_fs, self.zip_fs):
            # exercise -----------------
            self.pri.load(fs)
            # verify ------------------
            expected = 4
            actual = len(self.pri.relationships)
            msg = "expected message count of %d, got %d" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestPart(TestCase):
    """
    Test pptx.packaging.Part
    
    """
    def setUp(self):
        # create Part instance and ancestors
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.partname = '/ppt/slideMasters/slideMaster1.xml'
        self.content_type = 'application/vnd.openxmlformats-officedocument'\
                            '.presentationml.slideMaster+xml'
        self.part = pptx.packaging.Part()
        self.test_pptx_path = os.path.join(basepath, 'test_python-pptx.pptx')
    
    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)
    
    def test_class_present(self):
        """
        Part class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'Part')
    
    def test_blob_attribute_present(self):
        """
        Part instance has attribute 'blob'.
        
        """
        self.assertInstHasAttr(self.part, 'blob')
    
    def test_blob_none_on_construction(self):
        """
        Part.blob is None on construction.
        
        """
        expected = None
        actual = self.part.blob
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_blob_correct_length_after_load_binary_part(self):
        """
        Part.blob correct length after load binary part.
        
        """
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        content_type = 'image/jpeg'
        self.part.load(self.fs, partname, content_type)
        # exercise --------------------
        blob = self.part.blob
        # verify ----------------------
        expected = 21138
        actual = len(blob)
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_element_attribute_present(self):
        """
        Part instance has attribute 'element'.
        
        """
        self.assertInstHasAttr(self.part, 'element')
    
    def test_element_none_on_construction(self):
        """
        Part.element is None on construction.
        
        """
        expected = None
        actual = self.part.element
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_element_elm_after_load_xml_part(self):
        """
        Part.element instance of etree._Element after load XML part.
        
        """
        # setup -----------------------
        cls = etree._Element
        self.part.load(self.fs, self.partname, self.content_type)
        # exercise --------------------
        obj = self.part.element
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)
    
    def test_itemuri_attribute_present(self):
        """
        Part instance has attribute 'itemURI'.
        
        """
        self.assertInstHasAttr(self.part, 'itemURI')
    
    def test_itemuri_equals_partname(self):
        """
        Part.itemURI is the same as its part name.
        
        """
        # setup -----------------------
        # test before and after load()
        test_cases =\
            (self.part,
             self.part.load(self.fs, self.partname, self.content_type))
        # verify ----------------------
        for part in test_cases:
            expected = part.partname
            actual = part.itemURI
            msg = "expected itemURI %s, got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_load_returns_part(self):
        """
        Part.load() returns value of class Part.
        
        """
        # exercise --------------------
        part = self.part.load(self.fs, self.partname, self.content_type)
        # verify ----------------------
        expected = 'Part'
        actual = part.__class__.__name__
        msg = "expected classname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_loads_relationships(self):
        """
        Part.load() produces expected rels count.
        
        """
        # setup -----------------------
        part = self.part.load(self.fs, self.partname, self.content_type)
        # verify ----------------------
        expected = 12
        actual = len(part.relationships)
        msg = "expected %d relationships, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_throws_on_missing_rels_item(self):
        """
        Part.load() throws exception on missing rels file.
        
        Strategy: create a zip archive that has a part but not its rels file.
        """
        # setup -----------------------
        # get a valid xml part from filesystem
        partname = '/ppt/presentation.xml'
        fs = pptx.packaging.FileSystem(dir_pkg_path)
        elm = fs.getelement(partname)
        xml = etree.tostring(fs.getelement(partname))
        # use it to make a part in a new zip file
        zip = zipfile.ZipFile(self.test_pptx_path, 'w')
        zip.writestr('ppt/presentation.xml', xml)
        zip.close()
        # open new zip as FileSystem with part but no corresponding rels file
        fs = pptx.packaging.FileSystem(self.test_pptx_path)
        content_type = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.presentation.main+xml'
        # verify ----------------------
        with self.assertRaises(CorruptedPackageError):
            self.part.load(fs, partname, content_type)
    
    def test_partname_none_on_construction(self):
        """
        Part.partname is None on construction.
        
        """
        expected = None
        actual = self.part.partname
        msg = "expected partname value '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_partname_is_readonly(self):
        """
        Assignment to Part.partname throws AttributeError.
        
        """
        with self.assertRaises(AttributeError):
            self.part.partname = None
    
    def test_relationshipsitem_attribute_present(self):
        """
        Part instance has attribute 'relationshipsitem'.
        
        """
        self.assertInstHasAttr(self.part, 'relationshipsitem')
    
    def test_relationshipsitem_is_readonly(self):
        """
        Assignment to Part.relationshipsitem throws AttributeError.
        
        """
        with self.assertRaises(AttributeError):
            self.part.relationshipsitem = None
    
    def test_relationshipsitem_none_on_construction(self):
        """
        Part.relationshipsitem is None on construction.
        
        """
        expected = None
        actual = self.part.relationshipsitem
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationshipsitem_none_after_load_part_with_no_rels(self):
        """
        Part.relationshipsitem None after load part with no relationships.
        
        """
        # setup -----------------------
        partname = '/ppt/presProps.xml'
        content_type = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.presProps+xml'
        self.part.load(self.fs, partname, content_type)
        # verify ----------------------
        expected = None
        actual = self.part.relationshipsitem
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationshipsitem_partrelsitem_after_load_part_with_rels(self):
        """
        Part.relationshipsitem PartRelationshipsItem after load part w/rels.
        
        """
        # setup -----------------------
        cls = pptx.packaging.PartRelationshipsItem
        content_type = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.presProps+xml'
        self.part.load(self.fs, self.partname, self.content_type)
        # exercise --------------------
        obj = self.part.relationshipsitem
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("return value not instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)
    
    def test_relationships_attribute_present(self):
        """
        Part instance has attribute 'relationships'.
        
        """
        self.assertInstHasAttr(self.part, 'relationships')
    
    def test_relationships_is_readonly(self):
        """
        Assignment to Part.relationships throws AttributeError.
        
        """
        with self.assertRaises(AttributeError):
            self.part.relationships = None
    
    def test_relationships_none_on_construction(self):
        """
        Part.relationships is None on construction.
        
        """
        expected = None
        actual = self.part.relationships
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationships_none_after_load_part_with_no_rels(self):
        """
        Part.relationships None after load part with no rels.
        
        """
        # setup -----------------------
        partname = '/ppt/presProps.xml'
        content_type = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.presProps+xml'
        self.part.load(self.fs, partname, content_type)
        # verify ----------------------
        expected = None
        actual = self.part.relationshipsitem
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationships_relscollection_after_load_part_with_rels(self):
        """
        Part.relationships RelationshipCollection after load part w/rels.
        
        """
        # setup -----------------------
        cls = pptx.packaging.RelationshipCollection
        self.part.load(self.fs, self.partname, self.content_type)
        # exercise --------------------
        obj = self.part.relationships
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("expected instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)
    
    def test_typespec_attribute_present(self):
        """
        Part instance has attribute 'typespec'.
        
        """
        self.assertInstHasAttr(self.part, 'typespec')
    
    def test_typespec_none_on_construction(self):
        """
        Part.typespec is None on construction.
        
        """
        expected = None
        actual = self.part.typespec
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_typespec_correct_after_load(self):
        """
        Part.typespec is correct after load.
        
        """
        # setup -----------------------
        self.part.load(self.fs, self.partname, self.content_type)
        # exercise --------------------
        typespec = self.part.typespec
        test_cases =\
            ( (typespec.basename    , 'slideMaster')
            , (typespec.cardinality , 'multiple'   )
            , (typespec.has_rels    , 'always'     )
            , (typespec.format      , 'xml'        )
            )
        # verify ----------------------
        for actual, expected in test_cases:
            msg = "expected '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestPartCollection(TestCase):
    """
    Test pptx.packaging.PartCollection
    
    """
    def setUp(self):
        # Create a PartCollection instance
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.content_type = 'application/vnd.openxmlformats-officedocument'\
                            '.presentationml.slideLayout+xml'
        self.partname = '/ppt/slideLayouts/slideLayout1.xml'
        pkg = pptx.packaging.Package()
        self.parts = pkg.parts
    
    def test_class_present(self):
        """
        PartCollection class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'PartCollection')
    
    def test_contains(self):
        """
        'in' operator returns True if partname in collection.
        
        """
        # setup -----------------------
        self.parts.loadpart(self.fs, self.partname, self.content_type)
        # verify ----------------------
        msg = "expected '%s' in parts, got False" % (self.partname)
        self.assertIn(self.partname, self.parts, msg)
    
    def test_loadpart_adds_member(self):
        """
        loadpart() adds a member to PartCollection.
        
        """
        # exercise --------------------
        self.parts.loadpart(self.fs, self.partname, self.content_type)
        # verify ----------------------
        expected = 1
        actual = len(self.parts)
        msg = "expected %d parts, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_loadpart_returns_part(self):
        """
        PartCollection.loadpart() returns part.
        
        """
        # exercise --------------------
        part = self.parts.loadpart(self.fs, self.partname, self.content_type)
        # test ------------------------
        expected = 'Part'
        actual = part.__class__.__name__
        msg = "expected classname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_loadpart_loads_part(self):
        """
        Loaded part has matching partname.
        
        Further verification belongs to TestPart.test_load().
        
        """
        # exercise --------------------
        part = self.parts.loadpart(self.fs, self.partname, self.content_type)
        # test ------------------------
        expected = self.partname
        actual = part.partname
        msg = "expected partname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_loadpart_throws_on_dup_partname(self):
        """
        Attempt to load same part twice raises exception.
        
        """
        # setup -----------------------
        self.parts.loadpart(self.fs, self.partname, self.content_type)
        # trigger exception -----------
        with self.assertRaises(DuplicateKeyError):
            self.parts.loadpart(self.fs, self.partname, self.content_type)
    
    def test_getitem(self):
        """
        Loaded part can be retrieved by partname.
        
        """
        # setup -----------------------
        self.parts.loadpart(self.fs, self.partname, self.content_type)
        # exercise --------------------
        part = self.parts.getitem(self.partname)
        # verify ----------------------
        expected = self.partname
        actual = part.partname
        msg = "expected partname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_getitem_throws_on_bad_partname(self):
        """
        Lookup on part not in collection raises exception.
        
        """
        # setup -----------------------
        partname1 = '/ppt/slideLayouts/slideLayout1.xml'
        partname2 = '/ppt/slideMasters/slideMaster1.xml'
        self.parts.loadpart(self.fs, partname1, self.content_type)
        # trigger exception -----------
        with self.assertRaises(LookupError):
            self.parts.getitem(partname2)
    
    def test_package_attribute_present(self):
        """
        PartCollection instance has attribute 'package'.
        
        """
        self.assertInstHasAttr(self.parts, 'package')
    
    def test_package_ispackage_on_construction(self):
        """
        PartCollection.package is Package instance on construction.
        
        """
        # setup -----------------------
        cls = pptx.packaging.Package
        obj = self.parts.package
        # verify ----------------------
        actual = isinstance(obj, cls)
        msg = ("PartCollection.package not instance of '%s', got type '%s'"
               % (cls.__name__, type(obj).__name__))
        self.assertTrue(actual, msg)
    
    def test_package_is_readonly(self):
        """
        PartCollection.package is read-only.
        
        """
        with self.assertRaises(AttributeError):
            self.parts.package = None
    

class TestPartRelationshipsItem(TestCase):
    """
    Test pptx.packaging.PartRelationshipsItem
    
    """
    def setUp(self):
        # create Part instance
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.partname = '/ppt/slideMasters/slideMaster1.xml'
        self.relsitemURI = '/ppt/slideMasters/_rels/slideMaster1.xml.rels'
        self.content_type = 'application/vnd.openxmlformats-officedocument'\
                            '.presentationml.slideMaster+xml'
        part = pptx.packaging.Part()
        self.part = part.load(self.fs, self.partname, self.content_type)
        self.ri = pptx.packaging.PartRelationshipsItem()
    
    def test_class_present(self):
        """
        PartRelationshipsItem class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'PartRelationshipsItem')
    
    def test_baseuri_attribute_present(self):
        """
        PartRelationshipsItem instance has attribute 'baseURI'.
        
        """
        self.assertInstHasAttr(self.ri, 'baseURI')
    
    def test_baseuri_correct(self):
        """
        PartRelationshipsItem.baseURI returns correct value.
        
        """
        # setup -----------------------
        self.ri.load(self.part, self.fs, self.relsitemURI)
        # verify ----------------------
        expected = '/ppt/slideMasters'
        actual = self.ri.baseURI
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_constructor_throws_on_non_part(self):
        """
        PartRelationshipsItem(part=None) throws exception.
        
        """
        # verify ----------------------
        with self.assertRaises(TypeError):
            pptx.packaging.PartRelationshipsItem(None)
    
    def test_itemuri_attribute_present(self):
        """
        PartRelationshipsItem instance has attribute 'itemURI'.
        
        """
        self.assertInstHasAttr(self.ri, 'itemURI')
    
    def test_itemuri_calculation(self):
        """
        PartRelationshipsItem.itemURI returns correct URI.
        
        """
        # setup -----------------------
        test_datasets =\
            ( ('/ppt/presentation.xml',
               '/ppt/_rels/presentation.xml.rels')
            , ('/ppt/slideMasters/slideMaster1.xml',
               '/ppt/slideMasters/_rels/slideMaster1.xml.rels')
            )
        for partname, relsitemURI in test_datasets:
            part = pptx.packaging.Part()
            part.load(self.fs, partname, self.content_type)
            ri = pptx.packaging.PartRelationshipsItem()
            ri.load(part, self.fs, relsitemURI)
            # exercise ----------------
            itemURI = ri.itemURI
            # verify ------------------
            expected = relsitemURI
            actual = itemURI
            msg = "expected itemURI '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_load_callable(self):
        """
        PartRelationshipsItem has callable 'load'.
        
        The code for this test might be worth building a custom assertion
        around as this test type is likely to recur fairly often.
        """
        # parameters-------------------
        callable_name = 'load'
        cls = pptx.packaging.PartRelationshipsItem
        # setup -----------------------
        callables = [name for name in dir(cls) if callable(getattr(cls, name))]
        # verify ----------------------
        actual = callable_name in callables
        msg = "expected %s to have callable '%s'" % (cls, callable_name)
        self.assertTrue(actual, msg)
    
    def test_load_returns_self(self):
        """
        PartRelationshipsItem.load() returns reference to self.
        
        """
        # exercise --------------------
        retval = self.ri.load(self.part, self.fs, self.relsitemURI)
        # verify ----------------------
        expected = self.ri
        actual = retval
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_rels_count(self):
        """
        Loads expected number of relationships from rels item.
        
        """
        # exercise --------------------
        self.ri.load(self.part, self.fs, self.relsitemURI)
        # verify ----------------------
        expected = 12
        actual = len(self.ri.relationships)
        msg = "expected %d relationships to be loaded, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    

class TestPartTypeSpec(TestCase):
    """
    Test pptx.packaging.PartTypeSpec
    
    """
    def test_class_present(self):
        """
        PartTypeSpec class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'PartTypeSpec')
    
    def test_construction_returns_correct_parttypespec(self):
        """
        PartTypeSpec(content_type) returns correct PartTypeSpec.
        
        """
        # setup -----------------------
        content_type = 'application/vnd.openxmlformats-officedocument.'\
                       'presentationml.notesMaster+xml'
        # exercise --------------------
        pts = pptx.packaging.PartTypeSpec(content_type)
        # verify ----------------------
        actual = (pts.basename    == 'notesMaster'
              and pts.ext         == '.xml'
              and pts.cardinality == 'multiple'
              and pts.required    == False
              and pts.baseURI     == '/ppt/notesMasters'
              and pts.has_rels    == 'always'
              and pts.rel_type    == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster')
        msg = "PartTypeSpec('%s') returns unexpected values" % (content_type)
        self.assertTrue(actual, msg)
    
    def test_construction_throws_on_bad_contenttype(self):
        """
        PartTypeSpec(ct) throws exception on unrecognized content type.
        
        """
        # setup -----------------------
        content_type = 'application/vnd.baloneyMaster+xml'
        # verify ----------------------
        with self.assertRaises(KeyError):
            pptx.packaging.PartTypeSpec(content_type)
    
    def test_format_correct(self):
        """
        PartTypeSpec.format returns correct value.
        
        """
        # setup -----------------------
        ct_slideMaster = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.slideLayout+xml'
        ct_jpeg = 'image/jpeg'
        ct_printerSettings = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.printerSettings'
        ct_slide = 'application/vnd.openxmlformats-officedocument'\
            '.presentationml.slide+xml'
        test_cases =\
            ( ( ct_slideMaster     , 'xml'    )
            , ( ct_jpeg            , 'binary' )
            , ( ct_printerSettings , 'binary' )
            , ( ct_slide           , 'xml'    )
            )
        # exercise --------------------
        for ct, exp_format in test_cases:
            pts = pptx.packaging.PartTypeSpec(ct)
            # verify ----------------------
            expected = exp_format
            actual = pts.format
            msg = "expected '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestRelationship(TestCase):
    """
    Test pptx.packaging.Relationship
    
    """
    def setUp(self):
        """
        Create a new relationship from a string.
        
        """
        rId = 'rId1'
        self.reltype = 'http://schemas.openxmlformats.org/officeDocument/'\
                       '2006/relationships/officeDocument'
        target = 'ppt/presentation.xml'
        tmpl = '<Relationship Id="%s" Type="%s" Target="%s"/>'
        self.rel_xml = tmpl % (rId, self.reltype, target)
        self.rel_elm = etree.fromstring(self.rel_xml)
        # temporarily using None for parent reference
        self.rel = pptx.packaging.Relationship(None, self.rel_elm)
    
    def test_class_present(self):
        """
        Relationship class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'Relationship')
    
    def test_element_attribute_present(self):
        """
        Relationship instance has attribute 'element'.
        
        """
        self.assertInstHasAttr(self.rel, 'element')
    
    def test_element_correct_xml(self):
        """
        Relationship.element produces expected XML.
        
        """
        # exercise --------------------
        elm_out = self.rel.element
        # verify ----------------------
        xml_out = etree.tostring(elm_out)
        expected = self.rel_xml
        actual = xml_out
        msg = "expected '%s'\n%sgot '%s'" % (expected, ' '*21, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load(self):
        """
        Base attributes present after loading from ElementTree.Element.
        
        """
        rel = self.rel
        
        expected = 'rId1'
        actual = rel.rId
        msg = "expected Relationship.rId '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        
        expected = self.reltype
        actual = rel.reltype
        msg = "expected Relationship.reltype '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        
        expected = 'ppt/presentation.xml'
        actual = rel.target
        msg = "expected Relationship.target '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_target_partname(self):
        """
        Relationship.target_partname returns absolute URI.
        
        In relationship elements, the target is expressed as a relative URI
        that in many cases bears a strong resemblance to a part name. However,
        part names always begin with a forward slash (signifying the package
        root), always contain the full path to the part from the package root,
        and never contain '.' or '..' segments.
        
        """
        test_datasets =\
            [ ( '/'
              , '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
              , '/ppt/presentation.xml')
            , ( '/ppt'
              , '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
              , '/ppt/slides/slide1.xml')
            , ( '/ppt/slideLayouts'
              , '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
              , '/ppt/slideMasters/slideMaster1.xml' )
            ]
        for baseURI, rel_xml, expected in test_datasets:
            parent = MockParent(baseURI)
            rel = pptx.packaging.Relationship(parent, etree.fromstring(rel_xml))
            actual = rel.target_partname
            msg = "expected Relationship.target_partname '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    

class TestRelationshipCollection(TestCase):
    """
    Test pptx.packaging.RelationshipCollection
    
    """
    def setUp(self):
        # ancestor object graph
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.part = pptx.packaging.Part()
        self.partname = '/ppt/slideMasters/slideMaster1.xml'
        self.relsitemURI = '/ppt/slideMasters/_rels/slideMaster1.xml.rels'
        self.content_type = 'application/vnd.openxmlformats-officedocument'\
                            '.presentationml.slideMaster+xml'
        self.part.load(self.fs, self.partname, self.content_type)
        # Create a RelationshipCollection.
        mock_source = MockParent('/ppt/slideLayouts')
        self.relationships = pptx.packaging.RelationshipCollection(mock_source)
        # construct a relationship element
        rel_xml = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        self.rel_elm = etree.fromstring(rel_xml)
    
    def test_class_present(self):
        """
        RelationshipCollection class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'RelationshipCollection')
    
    def test_additem(self):
        """
        New relationship can be round-tripped intact.
        
        """
        # add a relationship to the RelationshipCollection
        new_rel = self.relationships.additem(self.rel_elm)
        # test intact retrieval
        expected = new_rel
        actual = self.relationships.getitem(new_rel.rId)
        msg = "newly added relationship not retrieved intact:\n"\
              "    Added ==> '%s'\n"\
              "Retrieved ==> '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_additem_throws_on_duplicate_item(self):
        """
        Attempt to add relationship with duplicate key raises exception.
        
        """
        # add a relationship to the RelationshipCollection
        self.relationships.additem(self.rel_elm)
        # trigger exception by attempt to add it again
        with self.assertRaises(DuplicateKeyError):
            self.relationships.additem(self.rel_elm)
    
    def test_baseuri_attribute_present(self):
        """
        RelationshipsCollection instance has attribute 'baseURI'.
        
        """
        self.assertInstHasAttr(self.relationships, 'baseURI')
    
    def test_baseuri_is_readonly(self):
        """
        RelationshipCollection.baseURI is read-only.
        
        """
        with self.assertRaises(AttributeError):
            self.relationships.baseURI = None
    
    def test_baseuri_same_as_parent(self):
        """
        RelationshipCollection.baseURI same as parent's.
        
        """
        # setup -----------------------
        pkg_ri = pptx.packaging.PackageRelationshipsItem()
        part_ri = pptx.packaging.PartRelationshipsItem()
        part_ri.load(self.part, self.fs, self.relsitemURI)
        test_cases = (pkg_ri, part_ri)
        for ri in test_cases:
            # exercise ----------------
            ri_baseURI = ri.baseURI
            rels_baseURI = ri.relationships.baseURI
            # verify ------------------
            expected = ri_baseURI
            actual = rels_baseURI
            msg = "expected baseURI '%s', got '%s'" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_getitem_throws_on_bad_key(self):
        """
        Attempt to look up relationship not in collection raises exception.
        
        """
        # setup -----------------------
        self.relationships.additem(self.rel_elm)
        # verify ----------------------
        with self.assertRaises(LookupError):
            self.relationships.getitem('rId9')
    

class TestRelationshipsItem(TestCase):
    """
    Test pptx.packaging.RelationshipsItem
    
    """
    def setUp(self):
        """
        Load package relationships item from default filesystem.
        
        """
        self.ri = pptx.packaging.RelationshipsItem()
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.relsitemURI = '/_rels/.rels'
    
    def loaded_ri(self):
        ri = pptx.packaging.RelationshipsItem()
        return ri.load(self.fs, self.relsitemURI)
    
    def test_class_present(self):
        """
        RelationshipsItem class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'RelationshipsItem')
    
    def test_element_attribute_present(self):
        """
        RelationshipsItem instance has attribute 'element'.
        
        """
        self.assertInstHasAttr(self.ri, 'element')
    
    def test_element_correct_length(self):
        """
        RelationshipsItem.element contains expected count of elements.
        
        """
        # setup -----------------------
        loaded_ri = self.loaded_ri()
        # exercise --------------------
        elm = loaded_ri.element
        # verify ----------------------
        expected = 4
        actual = len(elm)
        msg = "expected %d elements, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_element_correct_xml(self):
        """
        RelationshipsItem.element produces expected XML.
        
        """
        # setup -----------------------
        ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        exp_xml = '<Relationships xmlns="%s"/>' % ns
        # exercise --------------------
        elm = self.ri.element
        # verify ----------------------
        xml_out = etree.tostring(elm)
        expected = exp_xml
        actual = xml_out
        msg = "expected '%s'\n%sgot '%s'" % (expected, ' '*21, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_discards_prior_rels(self):
        """
        RelationshipsItem.load() discards existing relationships.
        
        Tries loading /_rels/.rels twice.
        """
        # setup -----------------------
        loaded_ri = self.loaded_ri()
        # exercise --------------------
        try:  # try loading a second time
            loaded_ri.load(self.fs, self.relsitemURI)
        # verify ----------------------
        except DuplicateKeyError:
            msg = "prior load relationship(s) not discarded"
            self.fail(msg)
        expected = 4
        actual = len(loaded_ri.relationships)
        msg = "expected %d relationships, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_rels_count(self):
        """
        Expected number of relationships loaded from rels item.
        
        """
        # setup -----------------------
        loaded_ri = self.loaded_ri()
        # verify ----------------------
        expected = 4
        actual = len(loaded_ri.relationships)
        msg = "expected %d relationships, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationships_empty_on_construction(self):
        """
        RelationshipsItem.relationships empty on construction.
        
        """
        # verify ----------------------
        expected = 0
        actual = len(self.ri.relationships)
        msg = "expected %d relationships, got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationships_is_readonly(self):
        """
        Assignment to RelationshipsItem.relationships throws exception.
        
        """
        with self.assertRaises(AttributeError):
            self.ri.relationships = None
    

class TestZipFileSystem(TestCase):
    """
    Test pptx.packaging.ZipFileSystem (writing aspect)
    
    """
    def setUp(self):
        self.test_pptx_path = os.path.join(basepath, 'test_python-pptx.pptx')
        self.xml_in =\
            """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"""\
            """<p:presentationPr xmlns:a="http://main" """\
            """xmlns:r="http://relationships" """\
            """xmlns:p="http://presentationml">"""\
            """<p:extLst>"""\
            """<p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}">"""\
            """<r:discardImageEditData val="0"/>"""\
            """</p:ext>"""\
            """</p:extLst>"""\
            """</p:presentationPr>"""
        self.xml_out =\
            """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"""\
            """<p:presentationPr xmlns:a="http://main"\n"""\
            """                  xmlns:r="http://relationships"\n"""\
            """                  xmlns:p="http://presentationml">\n"""\
            """  <p:extLst>\n"""\
            """    <p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}">\n"""\
            """      <r:discardImageEditData val="0"/>\n"""\
            """    </p:ext>\n"""\
            """  </p:extLst>\n"""\
            """</p:presentationPr>"""
    
    def tearDown(self):
        if os.path.isfile(self.test_pptx_path):
            os.remove(self.test_pptx_path)
    
    def test_class_present(self):
        """
        ZipFileSystem class present in packaging module.
        
        """
        # verify ----------------------
        self.assertClassInModule(pptx.packaging, 'ZipFileSystem')
    
    def test_getblob_method_present(self):
        """
        ZipFileSystem class has method 'getblob'.
        
        """
        self.assertClassHasMethod(pptx.packaging.ZipFileSystem,
                                  'getblob')
    
    def test_getblob_correct_length(self):
        """
        ZipFileSystem.getblob() returns object of expected length.
        
        """
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        # exercise --------------------
        blob = fs.getblob(partname)
        # verify ----------------------
        expected = 21138
        actual = len(blob)
        msg = "expected %d bytes, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_getblob_throws_on_bad_URI(self):
        """
        ZipFileSystem.getblob() throws on bad URI.
        
        """
        # setup -----------------------
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        # verify ----------------------
        with self.assertRaises(LookupError):
            fs.getstream('!blat/rhumba.xml')
    
    def test_new_method_present(self):
        """
        ZipFileSystem class has method 'new'.
        
        """
        self.assertClassHasMethod(pptx.packaging.ZipFileSystem, 'new')
    
    def test_new_returns_self(self):
        """
        ZipFileSystem.new() returns self-reference.
        
        """
        # exercise --------------------
        zipfs = pptx.packaging.ZipFileSystem(self.test_pptx_path)
        retval = zipfs.new()
        # verify ----------------------
        expected = zipfs
        actual = retval
        msg = "expected %s, got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_write_blob_method_present(self):
        """
        ZipFileSystem class has method 'write_blob'.
        
        """
        self.assertClassHasMethod(pptx.packaging.ZipFileSystem,
                                  'write_blob')
    
    def test_write_blob_round_trips(self):
        """
        ZipFileSystem.write_blob() round-trips intact.
        
        """
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        in_blob = fs.getblob(partname)
        test_fs = pptx.packaging.ZipFileSystem(self.test_pptx_path).new()
        # exercise --------------------
        test_fs.write_blob(in_blob, partname)
        # verify ----------------------
        out_blob = test_fs.getblob(partname)
        expected = in_blob
        actual = out_blob
        msg = "retrived blob (len %d) differs from original (len %d)"\
               % (len(actual), len(expected))
        self.assertEqual(expected, actual, msg)
    
    def test_write_blob_throws_on_dup_itemuri(self):
        """
        ZipFileSystem.write_blob() throws on duplicate itemURI.
        
        """
        # setup -----------------------
        partname = '/docProps/thumbnail.jpeg'
        fs = pptx.packaging.FileSystem(zip_pkg_path)
        blob = fs.getblob(partname)
        test_fs = pptx.packaging.ZipFileSystem(self.test_pptx_path).new()
        test_fs.write_blob(blob, partname)
        # verify ----------------------
        with self.assertRaises(DuplicateKeyError):
            test_fs.write_blob(blob, partname)
    
    def test_write_element_method_present(self):
        """
        ZipFileSystem class has method 'write_element'.
        
        """
        self.assertClassHasMethod(pptx.packaging.ZipFileSystem,
                                  'write_element')
    
    def test_write_element_round_trips(self):
        """
        ZipFileSystem.write_element() round-trips intact.
        
        """
        # setup -----------------------
        elm = etree.fromstring(self.xml_in)
        itemURI = '/ppt/test.xml'
        zipfs = pptx.packaging.ZipFileSystem(self.test_pptx_path).new()
        # exercise --------------------
        zipfs.write_element(elm, itemURI)
        # verify ----------------------
        stream = zipfs.getstream(itemURI)
        xml_out = stream.read()
        stream.close()
        expected = self.xml_out
        actual = xml_out
        msg = "expected \n%s\n, got\n%s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_write_element_throws_on_dup_itemuri(self):
        """
        ZipFileSystem.write_element() throws on duplicate itemURI.
        
        """
        # setup -----------------------
        elm = etree.fromstring(self.xml_in)
        itemURI = '/ppt/test.xml'
        zipfs = pptx.packaging.ZipFileSystem(self.test_pptx_path).new()
        # exercise --------------------
        zipfs.write_element(elm, itemURI)
        # verify ----------------------
        with self.assertRaises(DuplicateKeyError):
            zipfs.write_element(elm, itemURI)
    


