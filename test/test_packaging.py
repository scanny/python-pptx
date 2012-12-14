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

from lxml import etree
from StringIO import StringIO

import pptx.packaging

from pptx.exceptions import CorruptedPackageError, DuplicateKeyError,\
                            NotXMLError


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

class TestContentTypesItem(unittest.TestCase):
    """
    Test pptx.packaging.ContentTypesItem
    
    """
    def setUp(self):
        # construct ContentTypesItem instance
        self.cti = pptx.packaging.ContentTypesItem()
    
    def test_class_exists(self):
        """
        ContentTypesItem class exists.
        
        """
        # verify ----------------------
        try:
            issubclass(pptx.packaging.ContentTypesItem,
                       pptx.packaging.ContentTypesItem)
        except AttributeError:
            msg = "ContentTypesItem class not found"
            self.fail(msg)
        except TypeError:
            msg = "ContentTypesItem is not a class"
            self.fail(msg)
    
    def test_getitem_method_exists(self):
        """
        ContentTypesItem class has method '__getitem__'.
        
        NOTE: This test will fail if attribute is an @property method that
        throws an exception.
        """
        # setup -----------------------
        cls = pptx.packaging.ContentTypesItem
        method_name = '__getitem__'
        methods = inspect.getmembers(cls, inspect.ismethod)
        # verify ----------------------
        actual = method_name in [name for name, value in methods]
        msg = "no method %s.%s()" % (cls.__name__, method_name)
        self.assertTrue(actual, msg)
    
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
    

class TestFileSystem(unittest.TestCase):
    """
    Test pptx.packaging.FileSystem
    
    """
    def setUp(self):
        self.dir_fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.zip_fs = pptx.packaging.FileSystem(zip_pkg_path)
    
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
    
    def test_getelement(self):
        """
        ElementTree element for specified package item is returned.
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            elm = fs.getelement('/[Content_Types].xml')
            expected = 24
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
    
    def test_getstream(self):
        """
        StringIO instance for specified package item is returned.
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            stream = fs.getstream('/[Content_Types].xml')
            elm = etree.parse(stream).getroot()
            expected = 24
            actual = len(elm)
            msg = "expected %d elements, got %d" % (expected, actual)
            self.assertEqual(expected, actual, msg)
    
    def test_getstream_throws_on_bad_URI(self):
        """
        LookupError thrown on bad URI to getstream().
        
        """
        for fs in (self.dir_fs, self.zip_fs):
            with self.assertRaises(LookupError):
                fs.getstream('!blat/rhumba.xml')
    
    def test_itemURIs_count(self):
        """
        FileSystem.itemURIs has expected count.
        
        """
        # verify ----------------------
        for fs in (self.dir_fs, self.zip_fs):
            expected = 37
            actual = len(fs.itemURIs)
            msg = "expected %d filesystem members, got %d" % (expected, actual)
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
    

class TestPackage(unittest.TestCase):
    """
    Test pptx.packaging.Package
    
    """
    def setUp(self):
        # construct Package
        self.pkg = pptx.packaging.Package()
    
    def test_class_exists(self):
        """
        Package class exists.
        
        """
        # verify ----------------------
        try:
            issubclass(pptx.packaging.Package,
                       pptx.packaging.Package)
        except AttributeError:
            msg = "Package class not found"
            self.fail(msg)
        except TypeError:
            msg = "Package is not a class"
            self.fail(msg)
    
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
    
    def test_parts_attribute(self):
        """
        Package instance has attribute 'parts'.
        
        """
        # setup -----------------------
        obj = self.pkg
        name = 'parts'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    

class TestPackageRelationshipsItem(unittest.TestCase):
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
    
    def test_baseuri_attribute(self):
        """
        PackageRelationshipsItem instances have attribute 'baseURI'.
        
        NOTE: This test will fail if attribute is a method that throws an
        exception.
        
        """
        # setup -----------------------
        obj = self.pri
        name = 'baseURI'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    
    def test_itemuri_attribute(self):
        """
        PackageRelationshipsItem instances have attribute 'itemURI'.
        
        NOTE: This test will fail if attribute is an @property method that
        throws an exception.
        """
        # setup -----------------------
        obj = self.pri
        name = 'itemURI'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    

class TestPart(unittest.TestCase):
    """
    Test pptx.packaging.Part
    
    """
    def setUp(self):
        # create Part instance and ancestors
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.content_type = 'application/vnd.openxmlformats-officedocument'\
                            '.presentationml.slideMaster+xml'
        self.part = pptx.packaging.Part()
        self.partname = '/ppt/slideMasters/slideMaster1.xml'
    
    def test_class_exists(self):
        """
        Part class exists.
        
        """
        # verify ----------------------
        try:
            issubclass(pptx.packaging.Part,
                       pptx.packaging.Part)
        except AttributeError:
            msg = "Part class not found"
            self.fail(msg)
        except TypeError:
            msg = "Part is not a class"
            self.fail(msg)
    
    def test_itemuri_attribute(self):
        """
        Part instances have attribute 'itemURI'.
        
        """
        # setup -----------------------
        obj = self.part
        name = 'itemURI'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
        
        """
        # setup -----------------------
        partname = '/ppt/slides/slide99.xml'
        content_type = 'application/vnd.openxmlformats-officedocument'\
                       '.presentationml.slide+xml'
        # verify ----------------------
        with self.assertRaises(CorruptedPackageError):
            self.part.load(self.fs, partname, content_type)
    
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
    
    def test_relationshipsitem_attribute(self):
        """
        Part instances have attribute 'relationshipsitem'.
        
        """
        # setup -----------------------
        obj = self.part
        name = 'relationshipsitem'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    
    def test_relationships_attribute(self):
        """
        Part instances have attribute 'relationships'.
        
        """
        # setup -----------------------
        obj = self.part
        name = 'relationships'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    

class TestPartCollection(unittest.TestCase):
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
    
    def test_class_exists(self):
        """
        PartCollection class exists.
        
        """
        # verify ----------------------
        try:
            issubclass(pptx.packaging.PartCollection,
                       pptx.packaging.PartCollection)
        except AttributeError:
            msg = "PartCollection class not found"
            self.fail(msg)
        except TypeError:
            msg = "PartCollection is not a class"
            self.fail(msg)
    
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
    
    def test_package_attribute(self):
        """
        PartCollection instance has attribute 'package'.
        
        """
        # setup -----------------------
        obj = self.parts
        name = 'package'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    

class TestPartRelationshipsItem(unittest.TestCase):
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
    
    def test_class_exists(self):
        """
        pptx.packaging.PartRelationshipsItem class exists.
        
        """
        # verify ----------------------
        try:
            issubclass(pptx.packaging.PartRelationshipsItem,
                       pptx.packaging.PartRelationshipsItem)
        except AttributeError:
            msg = "PartRelationshipsItem class not found"
            self.fail(msg)
        except TypeError:
            msg = "PartRelationshipsItem is not a class"
            self.fail(msg)
    
    def test_baseuri_attribute(self):
        """
        PartRelationshipsItem instance has attribute 'baseURI'.
        
        """
        # setup -----------------------
        obj = self.ri
        name = 'baseURI'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    
    def test_itemuri_attribute(self):
        """
        PartRelationshipsItem instances have attribute 'itemURI'.
        
        NOTE: This test will fail if attribute is an @property method that
        throws an exception.
        """
        # setup -----------------------
        obj = self.ri
        name = 'itemURI'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    

class TestPartTypeSpec(unittest.TestCase):
    """
    Test pptx.packaging.PartTypeSpec
    
    """
    def test_class_exists(self):
        """
        PartTypeSpec class exists.
        
        """
        # verify ----------------------
        try:
            issubclass(pptx.packaging.PartTypeSpec,
                       pptx.packaging.PartTypeSpec)
        except AttributeError:
            msg = "PartTypeSpec class not found"
            self.fail(msg)
        except TypeError:
            msg = "PartTypeSpec is not a class"
            self.fail(msg)
    
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
    

class TestRelationship(unittest.TestCase):
    """
    Test pptx.packaging.Relationship
    
    """
    def setUp(self):
        """
        Create a new relationship from a string.
        
        """
        rel_xml =\
            '<Relationship Id="rId1"'\
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'\
            ' Target="ppt/presentation.xml"/>'
        # temporarily using None for parent reference
        self.relationship = pptx.packaging.Relationship(None, etree.fromstring(rel_xml))
    
    def test_load(self):
        """
        Base attributes present after loading from ElementTree.Element.
        
        """
        rel = self.relationship
        
        expected = 'rId1'
        actual = rel.rId
        msg = "expected Relationship.rId '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        
        expected = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
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
    

class TestRelationshipCollection(unittest.TestCase):
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
    
    def test_baseuri_attribute(self):
        """
        RelationshipsCollection instances have attribute 'baseURI'.
        
        """
        # setup -----------------------
        obj = self.relationships
        name = 'baseURI'
        # verify ----------------------
        actual = hasattr(obj, name)
        msg = "expected %s to have attribute '%s'" % (obj, name)
        self.assertTrue(actual, msg)
    
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
    

class TestRelationshipsItem(unittest.TestCase):
    """
    Test pptx.packaging.RelationshipsItem
    
    """
    def setUp(self):
        """
        Load package relationships item from default filesystem.
        
        """
        ri = pptx.packaging.RelationshipsItem()
        self.fs = pptx.packaging.FileSystem(dir_pkg_path)
        self.relsitemURI = '/_rels/.rels'
        self.ri = ri.load(self.fs, self.relsitemURI)
    
    def test_construction(self):
        """
        RelationshipsItem instance can be constructed.
        
        """
        # exercise --------------------
        ri = pptx.packaging.RelationshipsItem()
        # verify ----------------------
        expected = 'RelationshipsItem'
        actual = ri.__class__.__name__
        msg = "expected classname '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_discards_prior_rels(self):
        """
        RelationshipsItem.load() discards existing relationships.
        
        Tries loading /_rels/.rels twice.
        """
        # exercise --------------------
        try:  # setUp loads self.ri, so this is second time
            self.ri.load(self.fs, self.relsitemURI)
        # verify ----------------------
        except DuplicateKeyError:
            msg = "prior load relationship(s) not discarded"
            self.fail(msg)
        expected = 4
        actual = len(self.ri.relationships)
        msg = "expected %d relationships, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_load_rels_count(self):
        """
        Expected number of relationships loaded from rels item.
        
        """
        expected = 4
        actual = len(self.ri.relationships)
        msg = "expected %d relationships, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationships_empty_on_construction(self):
        """
        RelationshipsItem.relationships empty on construction.
        
        """
        ri = pptx.packaging.RelationshipsItem()
        actual = len(ri.relationships)
        expected = 0
        msg = "expected %d relationships, got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def test_relationships_is_readonly(self):
        """
        Assignment to RelationshipsItem.relationships throws exception.
        
        """
        with self.assertRaises(AttributeError):
            self.ri.relationships = None
    


