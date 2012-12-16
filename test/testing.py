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
    """
    Additional assert methods for python-pptx unit testing.
    
    """
    def assertClassInModule(self, module, cls_name):
        """
        Throw AssertionError if class *cls_name* does not exist in *module*.
        
        """
        # verify ----------------------
        members = inspect.getmembers(module)
        cls_members = inspect.getmembers(module, inspect.isclass)
        names = [name for name, obj in members]
        cls_names = [name for name, cls in cls_members]
        if cls_name not in names:
            tmpl = "no class '%s' in module '%s'"
            raise AssertionError(tmpl % (cls_name, module.__name__))
        if cls_name not in cls_names:
            tmpl = "'%s' in module '%s' is not a class"
            raise AssertionError(tmpl % (cls_name, module.__name__))
    
    def assertClassHasMethod(self, cls, method_name):
        """
        Throw AssertionError if no method *method_name* in *cls*.
        
        NOTE: This test will fail if method throws an exception.
        
        """
        methods = inspect.getmembers(cls, inspect.ismethod)
        if method_name not in [name for name, meth_obj in methods]:
            tmpl = "no method %s.%s()"
            raise AssertionError(tmpl % (cls.__name__, method_name))
    
    def assertInstHasAttr(self, inst, attr_name):
        """
        Throw AssertionError if no name *attr_name* in *inst*.
        
        NOTE: This test will fail if attribute is a property method that
              throws an exception when called.
        
        """
        names = [name for name, obj in inspect.getmembers(inst)]
        if attr_name not in names:
            tmpl = "expected %s to have attribute '%s'"
            raise AssertionError(tmpl %  (inst, attr_name))
    

