# encoding: utf-8

"""
lxml custom element classes for core properties-related XML elements.
"""

from __future__ import absolute_import

import re

from datetime import datetime, timedelta

from lxml import objectify

from pptx.oxml import element_class_lookup, oxml_fromstring
from pptx.spec import nsmap

from .util import nsdecls, qn


class CT_CoreProperties(objectify.ObjectifiedElement):
    """
    ``<cp:coreProperties>`` element, the root element of the Core Properties
    part stored as ``/docProps/core.xml``. Implements many of the Dublin Core
    document metadata elements. String elements resolve to an empty string
    ('') if the element is not present in the XML. String elements are
    limited in length to 255 unicode characters.
    """
    _date_tags = {
        'created':      'dcterms:created',
        'last_printed': 'cp:lastPrinted',
        'modified':     'dcterms:modified',
    }
    _str_tags = {
        'author':           'dc:creator',
        'category':         'cp:category',
        'comments':         'dc:description',
        'content_status':   'cp:contentStatus',
        'identifier':       'dc:identifier',
        'keywords':         'cp:keywords',
        'language':         'dc:language',
        'last_modified_by': 'cp:lastModifiedBy',
        'subject':          'dc:subject',
        'title':            'dc:title',
        'version':          'cp:version',
    }
    _coreProperties_tmpl = (
        '<cp:coreProperties %s/>\n' % nsdecls('cp', 'dc', 'dcterms')
    )

    @staticmethod
    def new_coreProperties():
        """Return a new ``<cp:coreProperties>`` element"""
        xml = CT_CoreProperties._coreProperties_tmpl
        coreProperties = oxml_fromstring(xml)
        return coreProperties

    def __getattribute__(self, name):
        """
        Intercept attribute access to generalize property getters.
        """
        if name in CT_CoreProperties._str_tags:
            return self._get_str_prop(name)
        elif name in CT_CoreProperties._date_tags:
            return self._get_date_prop(name)
        elif name == 'revision':
            return self._get_revision()
        else:
            return super(CT_CoreProperties, self).__getattribute__(name)

    def __setattr__(self, name, value):
        """
        Override ``__setattr__`` defined in ObjectifiedElement super class
        to intercept messages intended for custom property setters.
        """
        if name in CT_CoreProperties._str_tags:
            self._set_str_prop(name, value)
        elif name in CT_CoreProperties._date_tags:
            self._set_date_prop(name, value)
        elif name == 'revision':
            self._set_revision(value)
        else:
            super(CT_CoreProperties, self).__setattr__(name, value)

    def _get_str_prop(self, name):
        """Return string value of *name* property."""
        # explicit class reference avoids another pass through getattribute
        tag = qn(CT_CoreProperties._str_tags[name])
        if not hasattr(self, tag):
            return ''
        return getattr(self, tag).text

    def _get_date_prop(self, name):
        """Return datetime value of *name* property."""
        # explicit class reference avoids another pass through getattribute
        tag = qn(CT_CoreProperties._date_tags[name])
        # date properties return None when property element not present
        if not hasattr(self, tag):
            return None
        datetime_str = getattr(self, tag).text
        try:
            return self._parse_W3CDTF_to_datetime(datetime_str)
        except ValueError:
            # invalid datetime strings are ignored
            return None

    def _get_revision(self):
        """Return integer value of revision property."""
        tag = qn('cp:revision')
        # revision returns zero when element not present
        if not hasattr(self, tag):
            return 0
        revision_str = getattr(self, tag).text
        try:
            revision = int(revision_str)
        except ValueError:
            # non-integer revision strings also resolve to 0
            revision = 0
        # as do negative integers
        if revision < 0:
            revision = 0
        return revision

    def _set_str_prop(self, name, value):
        """Set string value of *name* property to *value*"""
        value = str(value)
        if len(value) > 255:
            tmpl = ("exceeded 255 char max length of property '%s', got:"
                    "\n\n'%s'")
            raise ValueError(tmpl % (name, value))
        tag = qn(CT_CoreProperties._str_tags[name])
        setattr(self, tag, value)
        # objectify will leave in a py: namespace without this cleanup
        elm = getattr(self, tag)
        objectify.deannotate(elm, cleanup_namespaces=True)

    def _set_date_prop(self, name, value):
        """Set datetime value of *name* property to *value*"""
        if not isinstance(value, datetime):
            tmpl = ("'%s' property requires <type 'datetime.datetime'> objec"
                    "t, got %s")
            raise ValueError(tmpl % (name, type(value)))
        tagname = CT_CoreProperties._date_tags[name]
        tag = qn(tagname)
        dt_str = value.strftime('%Y-%m-%dT%H:%M:%SZ')
        setattr(self, tag, dt_str)
        # objectify will leave in a py: namespace without this cleanup
        elm = getattr(self, tag)
        objectify.deannotate(elm, cleanup_namespaces=True)
        if name in ('created', 'modified'):
            # these two require an explicit 'xsi:type' attribute
            # first and last line are a hack required to add the xsi
            # namespace to the root element rather than each child element
            # in which it is referenced
            self.set(qn('xsi:foo'), 'bar')
            self[tag].set(qn('xsi:type'), 'dcterms:W3CDTF')
            del self.attrib[qn('xsi:foo')]

    def _set_revision(self, value):
        """Set integer value of revision property to *value*"""
        if not isinstance(value, int) or value < 1:
            tmpl = "revision property requires positive int, got '%s'"
            raise ValueError(tmpl % value)
        tag = qn('cp:revision')
        setattr(self, tag, str(value))
        # objectify will leave in a py: namespace without this cleanup
        elm = getattr(self, tag)
        objectify.deannotate(elm, cleanup_namespaces=True)

    _offset_pattern = re.compile('([+-])(\d\d):(\d\d)')

    @classmethod
    def _offset_dt(cls, dt, offset_str):
        """
        Return a |datetime| instance that is offset from datetime *dt* by
        the timezone offset specified in *offset_str*, a string like
        ``'-07:00'``.
        """
        match = cls._offset_pattern.match(offset_str)
        if match is None:
            raise ValueError("'%s' is not a valid offset string" % offset_str)
        sign, hours_str, minutes_str = match.groups()
        sign_factor = -1 if sign == '+' else 1
        hours = int(hours_str) * sign_factor
        minutes = int(minutes_str) * sign_factor
        td = timedelta(hours=hours, minutes=minutes)
        return dt + td

    @classmethod
    def _parse_W3CDTF_to_datetime(cls, w3cdtf_str):
        # valid W3CDTF date cases:
        # yyyy e.g. '2003'
        # yyyy-mm e.g. '2003-12'
        # yyyy-mm-dd e.g. '2003-12-31'
        # UTC timezone e.g. '2003-12-31T10:14:55Z'
        # numeric timezone e.g. '2003-12-31T10:14:55-08:00'
        templates = (
            '%Y-%m-%dT%H:%M:%S',
            '%Y-%m-%d',
            '%Y-%m',
            '%Y',
        )
        # strptime isn't smart enough to parse literal timezone offsets like
        # '-07:30', so we have to do it ourselves
        parseable_part = w3cdtf_str[:19]
        offset_str = w3cdtf_str[19:]
        dt = None
        for tmpl in templates:
            try:
                dt = datetime.strptime(parseable_part, tmpl)
            except ValueError:
                continue
        if dt is None:
            tmpl = "could not parse W3CDTF datetime string '%s'"
            raise ValueError(tmpl % w3cdtf_str)
        if len(offset_str) == 6:
            return cls._offset_dt(dt, offset_str)
        return dt

a_namespace = element_class_lookup.get_namespace(nsmap['cp'])
a_namespace['coreProperties'] = CT_CoreProperties
