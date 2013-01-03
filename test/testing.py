# -*- coding: utf-8 -*-
#
# testing.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Testing utilities for python-pptx."""

import inspect
import unittest


class TestCase(unittest.TestCase):
    """Additional assert methods for python-pptx unit testing."""
    def assertClassInModule(self, module, cls_name):
        """
        Raise AssertionError if class *cls_name* does not exist in *module*.
        
        """
        # verify ----------------------
        members = {name: obj for name, obj in inspect.getmembers(module)}
        if cls_name not in members:
            tmpl = "no class '%s' in module '%s'"
            raise AssertionError(tmpl % (cls_name, module.__name__))
        if not inspect.isclass(members[cls_name]):
            tmpl = "'%s' in module '%s' is not a class"
            raise AssertionError(tmpl % (cls_name, module.__name__))
    
    def assertClassHasMethod(self, cls, method_name):
        """
        Raise AssertionError if no method *method_name* in *cls*.
        
        NOTE: This test will fail if method raises an exception.
        """
        # members = {name: obj for name, obj in inspect.getmembers(cls)}
        # if method_name not in members:
        if not hasattr(cls, method_name):
            tmpl = "no method %s.%s()"
            raise AssertionError(tmpl % (cls.__name__, method_name))
        # if not inspect.ismethod(members[method_name]):
        if not inspect.ismethod(getattr(cls, method_name)):
            tmpl = "'%s' in class '%s' is not a method"
            raise AssertionError(tmpl % (method_name, cls.__name__))
    
    def assertInstHasAttr(self, inst, attr_name):
        """
        Raise AssertionError if no name *attr_name* in *inst*.
        
        NOTE: This test will fail if attribute is a property method that
              raises an exception when called.
        """
        if not hasattr(inst, attr_name):
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl % (inst, attr_name))
    
    def assertIsInstance(self, obj, cls):
        """Raise AssertionError if *obj* is not instance of *cls*."""
        tmpl = "expected instance of '%s', got type '%s'"
        if not isinstance(obj, cls):
            raise AssertionError(tmpl % (cls.__name__, type(obj).__name__))
    
    def assertIsProperty(self, inst, propname, value, read_only=True):
        """
        Raise AssertionError if *propname* is not a property of *obj* having
        the specified characteristics. Will raise AssertionError if *propname*
        does not exist in *obj*, if its value is not equal to *value*, or if
        *read_only* is True and assignment does not raise AttributeError.
        """
        if not hasattr(inst, propname):
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl % (inst, propname))
        expected = value
        actual = getattr(inst, propname)
        if actual != expected:
            tmpl = "expected '%s', got '%s'"
            raise AssertionError(tmpl % (expected, actual))
        if read_only:
            try:
                with self.assertRaises(AttributeError):
                    setattr(inst, propname, None)
            except AssertionError:
                tmpl = "property '%s' on class '%s' is not read-only"
                clsname = inst.__class__.__name__
                raise AssertionError(tmpl % (propname, clsname))
    
    def assertIsReadOnly(self, inst, propname):
        """
        Raise AssertionError if *propname* does not raise AttributeError when
        assignment is attempted.
        """
        try:
            with self.assertRaises(AttributeError):
                setattr(inst, propname, None)
        except AssertionError:
            tmpl = "%s.%s is not read-only"
            clsname = inst.__class__.__name__
            raise AssertionError(tmpl % (clsname, propname))
    
    def assertIsSizedProperty(self, inst, propname, length, read_only=True):
        """
        Raise AssertionError if *propname* is not a property of *obj* having
        the specified characteristics. Will raise AssertionError if *propname*
        does not exist in *obj*, if len(inst.propname) is not equal to
        *length*, or if *read_only* is True and assignment does not raise
        AttributeError.
        """
        if not hasattr(inst, propname):
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl % (inst, propname))
        expected = length
        actual = len(getattr(inst, propname))
        if actual != expected:
            tmpl = "expected length %d, got %d"
            raise AssertionError(tmpl % (expected, actual))
        if read_only:
            try:
                with self.assertRaises(AttributeError):
                    setattr(inst, propname, None)
            except AssertionError:
                tmpl = "property '%s' on class '%s' is not read-only"
                clsname = inst.__class__.__name__
                raise AssertionError(tmpl % (propname, clsname))
    
    def assertLength(self, sized, length):
        """Raise AssertionError if len(*sized*) != *length*"""
        expected = length
        actual = len(sized)
        msg = "expected length %d, got %d" % (expected, actual)
        self.assertEqual(expected, actual, msg)
    
    def assertPropertyRaisesOnNone(self, inst, propname, read_only=True):
        """
        Raise AssertionError if *propname* not in *inst* or doesn't raise
        ValueError. If *read_only* is True, raise if property accepts
        assignment.
        """
        # setup -----------------------
        clsname = inst.__class__.__name__
        # verify present --------------
        try:
            getattr(inst, propname)
        # raises AttributeError if name doesn't exist
        except AttributeError:
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl % (inst, propname))
        # raises ValueError if called when None
        except ValueError:
            pass
        # verify read-only ------------
        if read_only:
            try:
                setattr(inst, propname, None)
                tmpl = "property '%s' on class '%s' is not read-only"
                raise AssertionError(tmpl % (propname, clsname))
            except AttributeError:
                pass
        # verify raises on None -------
        try:
            getattr(inst, propname)
            tmpl = "%s.%s did not raise ValueError on None value"
            raise AssertionError(tmpl % (clsname, propname))
        except ValueError:
            pass
    

