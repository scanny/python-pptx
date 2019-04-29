# encoding: utf-8

"""
Test suite for pptx.oxml.ns.py module.
"""

from __future__ import print_function, unicode_literals

import pytest

from pptx.oxml.ns import NamespacePrefixedTag, namespaces, nsdecls, nsuri, qn


class DescribeNamespacePrefixedTag(object):
    def it_behaves_like_a_string_when_you_want_it_to(self, nsptag):
        s = "- %s -" % nsptag
        assert s == "- a:foobar -"

    def it_knows_its_clark_name(self, nsptag, clark_name):
        assert nsptag.clark_name == clark_name

    def it_knows_its_local_part(self, nsptag, local_part):
        assert nsptag.local_part == local_part

    def it_can_compose_a_single_entry_nsmap_for_itself(self, nsptag, namespace_uri_a):
        expected_nsmap = {"a": namespace_uri_a}
        assert nsptag.nsmap == expected_nsmap

    def it_knows_its_namespace_prefix(self, nsptag):
        assert nsptag.nspfx == "a"

    def it_knows_its_namespace_uri(self, nsptag, namespace_uri_a):
        assert nsptag.nsuri == namespace_uri_a


class DescribeNamespaces(object):
    def it_composes_a_dict_of_ns_uris_keyed_by_ns_pfx(self, nsmap):
        assert namespaces("a", "p") == nsmap


class DescribeNsdecls(object):
    def it_formats_namespace_declarations_from_a_list_of_prefixes(self, nsdecls_str):
        assert nsdecls("a", "p") == nsdecls_str


class DescribeNsuri(object):
    def it_finds_the_namespace_uri_corresponding_to_a_namespace_prefix(
        self, namespace_uri_a
    ):
        assert nsuri("a") == namespace_uri_a


class DescribeQn(object):
    def it_calculates_the_clark_name_for_an_ns_prefixed_tag_string(
        self, nsptag_str, clark_name
    ):
        assert qn(nsptag_str) == clark_name


# ===========================================================================
# fixtures
# ===========================================================================


@pytest.fixture
def clark_name(namespace_uri_a, local_part):
    return "{%s}%s" % (namespace_uri_a, local_part)


@pytest.fixture
def local_part():
    return "foobar"


@pytest.fixture
def namespace_uri_a():
    return "http://schemas.openxmlformats.org/drawingml/2006/main"


@pytest.fixture
def namespace_uri_p():
    return "http://schemas.openxmlformats.org/presentationml/2006/main"


@pytest.fixture
def nsdecls_str(namespace_uri_a, namespace_uri_p):
    return 'xmlns:a="%s" xmlns:p="%s"' % (namespace_uri_a, namespace_uri_p)


@pytest.fixture
def nsmap(namespace_uri_a, namespace_uri_p):
    return {"a": namespace_uri_a, "p": namespace_uri_p}


@pytest.fixture
def nsptag(nsptag_str):
    return NamespacePrefixedTag(nsptag_str)


@pytest.fixture
def nsptag_str(local_part):
    return "a:%s" % local_part
