#!/usr/bin/env python
# -*- coding: utf-8 -*-

# unpack-pptx.py
#
# Unzip a PowerPoint Open XML package (.pptx file) and pretty-print all
# the .xml and .rels files it contains.
#     Writes the files to a directory named to match the rootname of the
# pptx file, e.g. 'test.pptx' is extracted to directory 'test' in the same
# directory as the pptx file.

import os
import re
import shutil
import sys

from lxml    import etree
from zipfile import ZipFile


# indents second and later attributes on the root element so namespace
# declarations don't spread off the page in the text editor and can be more
# easily inspected
def prettify_nsdecls(xml):
    lines = xml.splitlines()
    if len(lines) < 2                   : return xml  # if XML is all on one line, don't mess with it
    if not lines[0].startswith('<?xml') : return xml  # if don't find xml declaration on first line, pass
    if not lines[1].startswith('<')     : return xml  # if don't find an unindented opening element on line 2, pass
    rootline = lines[1]
    # split rootline into element tag part and attributes parts
    attrib_re = re.compile(r'([-a-zA-Z0-9_:.]+="[^"]*" *>?)')
    substrings = [substring.strip() for substring in attrib_re.split(rootline) if substring]
    # substrings looks something like ['<p:sld', 'xmlns:p="html://..."', 'name="Office Theme>"']
    if len(substrings) < 3: # means there's at most one attributes so no need to indent
        return xml
    indent = ' ' * (len(substrings[0])+1)
    newrootline = ' '.join(substrings[:2])  # join element tag and first attribute onto same line
    for substring in substrings[2:]:        # indent remaining attributes on following lines
        newrootline += '\n%s%s' % (indent, substring)
    lines[1] = newrootline
    return '\n'.join(lines)


#=========================================================================
# main
#=========================================================================

if len(sys.argv) < 2:
    sys.stderr.write("\nusage: unpack-pptx.py PPTXFILE\n")
    exit()
else:
    path = sys.argv[1]

# split path into directory and filename
dir, filename = os.path.split(path)

# print """path='%s', dir='%s', filename='%s'""" % (path, dir, filename)
# sys.exit()

# get the rootname of the .pptx file as name for directory to unzip into
root, ext = os.path.splitext(filename)
outdir = os.path.join(dir, root)

# print """root='%s', ext='%s', outdir='%s'""" % (root, ext, outdir)
# sys.exit()

# remove the output directory if it exists so no old files linger
if os.path.isdir(outdir):
    shutil.rmtree(outdir)

# sys.exit()

# open the zipfile and extract all of its members to the output directory
zip = ZipFile(path)
zip.extractall(outdir)
zip.close()

# sys.exit()

# walk the tree and pretty print all the .xml and .rels files
targets = ['.xml', '.rels']
for dirpath, dirnames, filenames in os.walk(outdir):
    for filename in filenames:
        root, ext = os.path.splitext(filename)
        if ext in targets:
            filepath = os.path.join(dirpath, filename)
            print filepath
            f = open(filepath)
            xml = f.read()
            f.close()
            tree = etree.fromstring(xml)
            prettyxml = prettify_nsdecls(etree.tostring(tree, encoding='UTF-8', pretty_print=True, standalone=True))
            f = open(filepath, 'w')
            f.write(prettyxml)
            f.close()
