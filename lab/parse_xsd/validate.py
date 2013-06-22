#!/usr/bin/env python
# -*- coding: utf-8 -*-

# validate.py
#

"""
Experimental code to validate XML against an XML Schema
"""

import os
import sys

from lxml import etree

# # re-zip tweaked version of expanded package into a .pptx file
# $ cd test  # 'test' is the expanded package directory
# $ zip -rD -FS ../test-tweaked.pptx .

# ============================================================================
# Code templates
# ============================================================================

# xmlschema_doc = etree.parse('xsd/pml.xsd')
# xmlschema = etree.XMLSchema(xmlschema_doc)

# test_path = '../../../ar-fy-plans/test/ppt/slides'
#
# xmlschema = etree.XMLSchema(etree.parse('xsd/pml.xsd'))
#
# slide_path = os.path.join(test_path, 'slide1.xml')
# slide = etree.parse(slide_path)
# xmlschema.assert_(slide)
#
# sys.exit()
#
# for idx in range(1, 29):
#     slideLayout_path = os.path.join(test_path, 'slideLayout%d.xml' % idx)
#     sldLayout = etree.parse(slideLayout_path)
#     xmlschema.assert_(sldLayout)
#
# sys.exit()


test_path = '../../../ar-fy-plans/test/ppt/slideLayouts'

xmlschema = etree.XMLSchema(etree.parse('xsd/pml.xsd'))  # more compact form

for idx in range(1, 29):
    print 'slideLayout%d.xml' % idx
    slideLayout_path = os.path.join(test_path, 'slideLayout%d.xml' % idx)
    sldLayout = etree.parse(slideLayout_path)
    xmlschema.assert_(sldLayout)

sys.exit()


# assemble the necessary items
xmlschema = etree.XMLSchema(etree.parse('xsd/pml.xsd'))  # more compact form
test_path = '../../../ar-fy-plans/test/ppt/slides'
slide_path = os.path.join(test_path, 'slide1.xml')
sld = etree.parse(slide_path)  # should be ElementTree, not Element

# get valid True/False, no message
valid = xmlschema.validate(sld)
# print "valid => %s" % valid

# print out validation log for messages
# if not valid:
#     log = xmlschema.error_log
#     print type(log)
#     print(log.last_error)

# if not valid:
for e in xmlschema.error_log:
    print('%s:%s %s %s' % (e.filename, e.line, e.level_name, e.message))

# # raise an exception with validation error message if doesn't validate
# xmlschema.assertValid(sld)
# xmlschema.assert_(sld)  # bit shorter error message

sys.exit()
