# encoding: utf-8

"""Testing utilities for python-pptx."""

import os
import unittest2

from lxml import etree, objectify
from mock import create_autospec, Mock, patch, PropertyMock

from pptx.oxml import oxml_parser


_thisdir = os.path.split(__file__)[0]
test_file_dir = os.path.abspath(os.path.join(_thisdir, 'test_files'))


def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))


def parse_xml_file(file_):
    """
    Return ElementTree for XML contained in *file_*
    """
    return objectify.parse(file_, oxml_parser)


def relpath(relpath):
    thisdir = os.path.split(__file__)[0]
    return os.path.relpath(os.path.join(thisdir, relpath))


def serialize_xml(elm, pretty_print=False):
    objectify.deannotate(elm, xsi=False)
    xml = etree.tostring(elm, pretty_print=pretty_print)
    return xml


class TestCase(unittest2.TestCase):
    """Additional assert methods for python-pptx unit testing."""
    def assertEqualLineByLine(self, expected_xml, element):
        """
        Apply assertEqual() to each line of *expected_xml* and corresponding
        line of XML derived from *element*.
        """
        actual_xml = serialize_xml(element, pretty_print=True)
        actual_xml_lines = actual_xml.split('\n')
        expected_xml_lines = expected_xml.split('\n')
        for idx, line in enumerate(actual_xml_lines):
            msg = ("\n\nexpected:\n\n%s'\nbut got\n\n%s'" %
                   (expected_xml, actual_xml))
            self.assertEqual(line, expected_xml_lines[idx], msg)

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


def class_mock(request, q_class_name, autospec=True, **kwargs):
    """
    Return a mock patching the class with qualified name *q_class_name*.
    The mock is autospec'ed based on the patched class unless the optional
    argument *autospec* is set to False. Any other keyword arguments are
    passed through to Mock(). Patch is reversed after calling test returns.
    """
    _patch = patch(q_class_name, autospec=autospec, **kwargs)
    request.addfinalizer(_patch.stop)
    return _patch.start()


def function_mock(request, q_function_name):
    """
    Return a mock patching the function with qualified name
    *q_function_name*. Patch is reversed after calling test returns.
    """
    _patch = patch(q_function_name)
    request.addfinalizer(_patch.stop)
    return _patch.start()


def initializer_mock(request, cls):
    """
    Return a mock for the __init__ method on *cls* where the patch is
    reversed after pytest uses it.
    """
    _patch = patch.object(cls, '__init__', return_value=None)
    request.addfinalizer(_patch.stop)
    return _patch.start()


def instance_mock(request, cls, name=None, spec_set=True, **kwargs):
    """
    Return a mock for an instance of *cls* that draws its spec from the class
    and does not allow new attributes to be set on the instance. If *name* is
    missing or |None|, the name of the returned |Mock| instance is set to
    *request.fixturename*. Additional keyword arguments are passed through to
    the Mock() call that creates the mock.
    """
    if name is None:
        name = request.fixturename
    return create_autospec(cls, _name=name, spec_set=spec_set, instance=True,
                           **kwargs)


def loose_mock(request, name=None, **kwargs):
    """
    Return a "loose" mock, meaning it has no spec to constrain calls on it.
    Additional keyword arguments are passed through to Mock(). If called
    without a name, it is assigned the name of the fixture.
    """
    if name is None:
        name = request.fixturename
    return Mock(name=name, **kwargs)


def method_mock(request, cls, method_name):
    """
    Return a mock for method *method_name* on *cls* where the patch is
    reversed after pytest uses it.
    """
    _patch = patch.object(cls, method_name)
    request.addfinalizer(_patch.stop)
    return _patch.start()


def property_mock(request, q_property_name):
    """
    Return a mock for property with fully qualified name *q_property_name*
    where the patch is reversed after pytest uses it.
    """
    _patch = patch(q_property_name, new_callable=PropertyMock)
    request.addfinalizer(_patch.stop)
    return _patch.start()


def var_mock(request, q_var_name, **kwargs):
    """
    Return a mock patching the variable with qualified name *q_var_name*.
    Patch is reversed after calling test returns.
    """
    _patch = patch(q_var_name, **kwargs)
    request.addfinalizer(_patch.stop)
    return _patch.start()
