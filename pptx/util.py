# -*- coding: utf-8 -*-
#
# util.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

'''Utility functions that come in handy when working with PowerPoint and
Open XML.'''

import os

# utility function for calculating EMUs from inches
def emu(inches):
    return int(inches * 914400)


# utility sequential integer generator, suitable for generating unique ids.
def intsequence():
    num = 1
    while True:
        yield num
        num += 1


def prettify_nsdecls(xml):
    lines = xml.splitlines()
    if len(lines) < 2:         # if XML is all on one line, don't mess with it
        return xml
    rootline = lines[1]
    parts = rootline.split()
    if len(parts) < 2:
        return xml
    root_tag = parts[0]
    indent = len(root_tag) + 1
    newrootline = parts[0] + ' ' + parts[1]
    for part in parts[2:]:
        newrootline += ('\n' + ' '*indent + part)
    lines[1] = newrootline
    return '\n'.join(lines)

def sortedtemplatefilepaths(templatedir, searchdir, filenameroot, ext):
    # trim leading slash off of searchdir if present
    searchdir = searchdir[1:] if searchdir.startswith('/') else searchdir
    # form fully qualified path to search directory
    dirpath = os.path.join(templatedir, searchdir)
    # form list of all files in the directory
    filepaths = {}
    for name in os.listdir(dirpath):
        fqname = os.path.join(dirpath, name)
        if not os.path.isfile(fqname):
            continue
        if not name.startswith(filenameroot) or not name.endswith(ext):
            raise Exception('''Unexpected file '%s' found in template.''' % os.path.join(searchdir, name))
        filenamenumber = name[len(filenameroot):-(len(ext)+1)]
        sortkey = int(filenamenumber) if filenamenumber else 0
        filepaths[sortkey] = fqname

    #DELETEME:
    # singledigitones = [fqname for fqname in slidelayoutpaths if len(os.path.basename(fqname)) == 16]
    # doubledigitones = [fqname for fqname in slidelayoutpaths if len(os.path.basename(fqname)) == 17]
    # return sorted(singledigitones) + sorted(doubledigitones)

    # return file paths sorted in numerical order (not lexicographic order)
    return [filepaths[key] for key in sorted(filepaths.keys())]


#TECHDEBT: Not all files in the media directory are necessarily image files.
#          Audio and Video media can show up there too, although are perhaps
#          less likely to appear in a presentation template.
def templatemediafilepaths(templatedir):
    # form fully qualified path to search directory
    dirpath = os.path.join(templatedir, 'ppt/media')
    # if directory doesn't exist, it means the template doesn't contain any media files
    if not os.path.isdir(dirpath):
        return []
    # form list of all files in the directory
    filepaths = []
    for name in os.listdir(dirpath):
        filepath = os.path.join(dirpath, name)
        if not os.path.isfile(filepath):
            continue
        filepaths.append(filepath)
    return filepaths


    