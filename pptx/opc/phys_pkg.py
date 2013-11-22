# encoding: utf-8

"""
Provides a general interface to a *physical* OPC package, such as a zip file.
"""

from __future__ import absolute_import

import os

from lxml import etree

from StringIO import StringIO
from zipfile import ZipFile, is_zipfile, ZIP_DEFLATED

from pptx.exceptions import NotXMLError, PackageNotFoundError

from .packuri import CONTENT_TYPES_URI


class FileSystem(object):
    """
    Factory for filesystem interface instances.

    A FileSystem object provides access to on-disk package items via their URI
    (e.g. ``/_rels/.rels`` or ``/ppt/presentation.xml``). This allows parts to
    be accessed directly by part name, which for a part is identical to its
    item URI. The complexities of translating URIs into file paths or zip item
    names, and file and zip file access specifics are all hidden by the
    filesystem class. |FileSystem| acts as the Factory, returning the
    appropriate concrete filesystem class depending on what it finds at *path*.
    """
    def __new__(cls, file_):
        # if *file_* is a string, treat it as a path
        if isinstance(file_, basestring):
            path = file_
            if is_zipfile(path):
                fs = ZipFileSystem(path)
            elif os.path.isdir(path):
                fs = DirectoryFileSystem(path)
            else:
                raise PackageNotFoundError("Package not found at '%s'" % path)
        else:
            fs = ZipFileSystem(file_)
        return fs


class BaseFileSystem(object):
    """
    Base class for FileSystem classes, providing common methods.
    """
    def __init__(self):
        super(BaseFileSystem, self).__init__()

    def __contains__(self, itemURI):
        """
        Allows use of 'in' operator to test whether an item with the specified
        URI exists in this filesystem.
        """
        return itemURI in self.itemURIs

    def getblob(self, itemURI):
        """Return byte string of item identified by *itemURI*."""
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        stream = self.getstream(itemURI)
        blob = stream.read()
        stream.close()
        return blob

    def getelement(self, itemURI):
        """
        Return ElementTree element of XML item identified by *itemURI*.
        """
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        stream = self.getstream(itemURI)
        try:
            parser = etree.XMLParser(remove_blank_text=True)
            element = etree.parse(stream, parser).getroot()
        except etree.XMLSyntaxError:
            raise NotXMLError("package item %s is not XML" % itemURI)
        stream.close()
        return element


class DirectoryFileSystem(BaseFileSystem):
    """
    Provides access to package members that have been expanded into an on-disk
    directory structure.

    Inherits __contains__(), getelement(), and path from BaseFileSystem.
    """
    def __init__(self, path):
        """
        *path* is the path to a directory containing an expanded package.
        """
        super(DirectoryFileSystem, self).__init__()
        if not os.path.isdir(path):
            tmpl = "path '%s' not a directory"
            raise ValueError(tmpl % path)
        self._path = os.path.abspath(path)

    def close(self):
        """
        Provides interface consistency with |ZipFileSystem|, but does nothing,
        a directory file system doesn't need closing.
        """
        pass

    def getstream(self, itemURI):
        """
        Return file-like object containing package item identified by
        *itemURI*. Remember to call close() on the stream when you're done
        with it to free up the memory it uses.
        """
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        path = os.path.join(self._path, itemURI[1:])
        with open(path, 'rb') as f:
            stream = StringIO(f.read())
        return stream

    @property
    def itemURIs(self):
        """
        Return list of all filenames under filesystem root directory,
        formatted as item URIs. Each URI is the relative path of that file
        with a leading slash added, e.g. '/ppt/slides/slide1.xml'. Although
        not strictly necessary, the results are sorted for neatness' sake.
        """
        itemURIs = []
        for dirpath, dirnames, filenames in os.walk(self._path):
            for filename in filenames:
                item_path = os.path.join(dirpath, filename)
                itemURI = item_path[len(self._path):]  # leave leading slash
                itemURIs.append(itemURI.replace(os.sep, '/'))
        return sorted(itemURIs)


class _DirPkgReader(object):
    """
    Implements |PhysPkgReader| interface for an OPC package extracted into a
    directory.
    """
    def __init__(self, path):
        """
        *path* is the path to a directory containing an expanded package.
        """
        super(_DirPkgReader, self).__init__()
        self._path = os.path.abspath(path)

    def blob_for(self, pack_uri):
        """
        Return contents of file corresponding to *pack_uri* in package
        directory.
        """
        path = os.path.join(self._path, pack_uri.membername)
        with open(path, 'rb') as f:
            blob = f.read()
        return blob

    def close(self):
        """
        Provides interface consistency with |ZipFileSystem|, but does
        nothing, a directory file system doesn't need closing.
        """
        pass

    @property
    def content_types_xml(self):
        """
        Return the `[Content_Types].xml` blob from the package.
        """
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """
        Return rels item XML for source with *source_uri*, or None if the
        item has no rels item.
        """
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except IOError:
            rels_xml = None
        return rels_xml


class PhysPkgReader(object):
    """
    Factory for physical package reader objects.
    """
    def __new__(cls, pkg_file):
        # if *pkg_file* is a string, treat it as a path
        if isinstance(pkg_file, basestring):
            if os.path.isdir(pkg_file):
                reader_cls = _DirPkgReader
            elif is_zipfile(pkg_file):
                reader_cls = _ZipPkgReader
            else:
                raise PackageNotFoundError(
                    "Package not found at '%s'" % pkg_file
                )
        else:  # assume it's a stream and pass it to Zip reader to sort out
            reader_cls = _ZipPkgReader

        return super(PhysPkgReader, cls).__new__(reader_cls)


class PhysPkgWriter(object):
    """
    Factory for physical package writer objects.
    """
    def __new__(cls, pkg_file):
        return _ZipPkgWriter(pkg_file)


class ZipFileSystem(BaseFileSystem):
    """
    Return new instance providing access to zip-format OPC package contained
    in *file*, where *file* can be either a path to a zip file (a string) or a
    file-like object. If mode is 'w', a new zip archive is written to *file*.
    If *file* is a path and a file with that name already exists, it is
    truncated.

    Inherits :meth:`__contains__`, :meth:`getelement`, and :attr:`path` from
    BaseFileSystem.
    """
    def __init__(self, file_, mode='r'):
        super(ZipFileSystem, self).__init__()
        if 'w' in mode:
            self.zipf = ZipFile(file_, 'w', compression=ZIP_DEFLATED)
        else:
            self.zipf = ZipFile(file_, 'r')

    def close(self):
        """
        Close the |ZipFileSystem| instance, necessary to complete the write
        process with the instance is opened for writing.
        """
        self.zipf.close()

    def getstream(self, itemURI):
        """
        Return file-like object containing package item identified by
        *itemURI*. Remember to call close() on the stream when you're done
        with it to free up the memory it uses.
        """
        if itemURI not in self:
            raise LookupError("No package item with URI '%s'" % itemURI)
        membername = itemURI[1:]  # trim off leading slash
        stream = StringIO(self.zipf.read(membername))
        return stream

    @property
    def itemURIs(self):
        """
        Return list of archive members formatted as item URIs. Each member
        name is the archive-relative path of that file. A forward-slash is
        prepended to form the URI, e.g. '/ppt/slides/slide1.xml'. Although
        not strictly necessary, the results are sorted for neatness' sake.
        """
        names = self.zipf.namelist()
        # zip archive can contain entries for directories, so get rid of those
        itemURIs = [('/%s' % nm) for nm in names if not nm.endswith('/')]
        return sorted(itemURIs)


class _ZipPkgReader(object):
    """
    Implements |PhysPkgReader| interface for a zip file OPC package.
    """
    def __init__(self, pkg_file):
        super(_ZipPkgReader, self).__init__()
        self._zipf = ZipFile(pkg_file, 'r')

    def blob_for(self, pack_uri):
        """
        Return blob corresponding to *pack_uri*. Raises |ValueError| if no
        matching member is present in zip archive.
        """
        return self._zipf.read(pack_uri.membername)

    def close(self):
        """
        Close the zip archive, releasing any resources it is using.
        """
        self._zipf.close()

    @property
    def content_types_xml(self):
        """
        Return the `[Content_Types].xml` blob from the zip package.
        """
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """
        Return rels item XML for source with *source_uri* or None if no rels
        item is present.
        """
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except KeyError:
            rels_xml = None
        return rels_xml


class _ZipPkgWriter(object):
    """
    Implements |PhysPkgWriter| interface for a zip file OPC package.
    """
    def __init__(self, pkg_file):
        super(_ZipPkgWriter, self).__init__()
        self._zipf = ZipFile(pkg_file, 'w', compression=ZIP_DEFLATED)

    def close(self):
        """
        Close the zip archive, flushing any pending physical writes and
        releasing any resources it's using.
        """
        self._zipf.close()

    def write(self, pack_uri, blob):
        """
        Write *blob* to this zip package with the membername corresponding to
        *pack_uri*.
        """
        self._zipf.writestr(pack_uri.membername, blob)
