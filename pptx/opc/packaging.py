# encoding: utf-8

"""
The :mod:`pptx.packaging` module coheres around the concerns of reading and
writing presentations to and from a .pptx file. In doing so, it hides the
complexities of the package "directory" structure, reading and writing parts
to and from the package, zip file manipulation, and traversing relationship
items.

The main API class is :class:`pptx.packaging.Package` which provides the
methods :meth:`open`, :meth:`marshal`, and :meth:`save`.
"""

import os
import posixpath

from StringIO import StringIO
from lxml import etree
from zipfile import ZipFile, is_zipfile, ZIP_DEFLATED

import pptx.spec

from pptx.exceptions import (
    CorruptedPackageError, DuplicateKeyError, NotXMLError,
    PackageNotFoundError)

from pptx.spec import qtag
from pptx.spec import PTS_HASRELS_NEVER, PTS_HASRELS_OPTIONAL


PKG_BASE_URI = '/'


# ============================================================================
# API Classes
# ============================================================================

class Package(object):
    """
    Return a new package instance. Package is initially empty, call
    :meth:`open` to open an on-disk package or ``marshal()`` followed by
    ``save()`` to save an in-memory Office document.
    """

    PKG_RELSITEM_URI = '/_rels/.rels'

    def __init__(self):
        super(Package, self).__init__()
        self._relationships = []

    @property
    def parts(self):
        """
        Return a list of :class:`pptx.packaging.Part` corresponding to the
        parts in this package.
        """
        return [part for part in self._walkparts(self.relationships)]

    @property
    def relationships(self):
        """
        A tuple of :class:`pptx.packaging.Relationship` containing the package
        relationships for this package. Note these are not all the
        relationships in the package, just those from the package to top-level
        parts such as ``/ppt/presentation.xml`` and ``/docProps/core.xml``.
        These are useful primarily as the starting point to walk the part
        graph via its relationships.
        """
        return tuple(self._relationships)

    def open(self, file_):
        """
        Load the package contained in *file_*, where *file_* can be a path to
        a file or directory (a string), or a file-like object. If *file_* is
        a path to a directory, the directory must contain an expanded package
        such as is produced by unzipping an OPC package file.
        """
        fs = FileSystem(file_)
        cti = _ContentTypesItem().load(fs)
        self._relationships = []  # discard any rels from prior load
        parts_dict = {}           # track loaded parts, graph is cyclic
        pkg_rel_elms = fs.getelement(Package.PKG_RELSITEM_URI)\
                         .findall(qtag('pr:Relationship'))
        for rel_elm in pkg_rel_elms:
            rId = rel_elm.get('Id')
            reltype = rel_elm.get('Type')
            partname = '/%s' % rel_elm.get('Target')
            part = Part()
            parts_dict[partname] = part
            part._load(fs, partname, cti, parts_dict)
            rel = Relationship(rId, self, reltype, part)
            self._relationships.append(rel)
        fs.close()
        return self

    def marshal(self, model_pkg):
        """
        Load the contents of a model-side package such that it can be saved to
        a package file.
        """
        part_dict = {}  # keep track of marshaled parts, graph is cyclic
        for rel in model_pkg._relationships:
            # unpack working values for target part and relationship
            rId = rel._rId
            reltype = rel._reltype
            model_part = rel._target
            partname = model_part.partname
            # create package-part for target
            part = Part()
            part_dict[partname] = part
            part._marshal(model_part, part_dict)
            # create marshaled version of relationship
            marshaled_rel = Relationship(rId, self, reltype, part)
            self._relationships.append(marshaled_rel)
        return self

    def save(self, file):
        """
        Save this package to *file*, where *file* can be either a path to a
        file (a string) or a file-like object.
        """
        # open a zip filesystem for writing package
        zipfs = ZipFileSystem(file, 'w')
        # write [Content_Types].xml
        cti = _ContentTypesItem().compose(self.parts)
        zipfs.write_element(cti.element, '/[Content_Types].xml')
        # write pkg rels item
        zipfs.write_element(self.__relsitem_element, self.PKG_RELSITEM_URI)
        for part in self.parts:
            # write part item
            zipfs.write_blob(part.blob, part.partname)
            # write rels item if part has one
            if part.relationships:
                zipfs.write_element(part._relsitem_element, part._relsitemURI)
        zipfs.close()

    @property
    def __relsitem_element(self):
        nsmap = {None: pptx.spec.nsmap['pr']}
        element = etree.Element(qtag('pr:Relationships'), nsmap=nsmap)
        for rel in self._relationships:
            element.append(rel._element)
        return element

    @classmethod
    def _walkparts(cls, rels, parts=None):
        """
        Recursive generator method, walk relationships to iterate over all
        parts in this package. Leave out *parts* parameter in call to visit
        all parts.

        """
        # initial call can leave out parts parameter as a signal to initialize
        if parts is None:
            parts = []
        for rel in rels:
            part = rel.target
            if part in parts:  # only visit each part once (graph is cyclic)
                continue
            parts.append(part)
            yield part
            for part in cls._walkparts(part.relationships, parts):
                yield part


class Part(object):
    """
    Part instances are not intended to be constructed externally.
    :class:`pptx.packaging.Part` instances are constructed and initialized
    internally to the :meth:`Package.open` or :meth:`Package.marshal` methods.

    The following |Part| instance attributes can be accessed once the part has
    been loaded as part of a package:

    .. attribute:: typespec

       An instance of |PartTypeSpec| appropriate to the type of this part. The
       |PartTypeSpec| instance provides attributes such as *content_type*,
       *baseURI*, etc. That are useful in several contexts.

    .. attribute:: blob

       The binary contents of this part contained in a byte string. For XML
       parts, this is simply the XML text. For binary parts such as an image,
       this is the string of bytes corresponding exactly to the bytes on disk
       for the binary object.

    """
    def __init__(self):
        super(Part, self).__init__()
        self._partname = None
        self._relationships = []
        self.typespec = None
        self.blob = None

    @property
    def content_type(self):
        """Content type of this part"""
        assert self.typespec, 'Part.content_type called before typespec set'
        return self.typespec.content_type

    @property
    def partname(self):
        """
        Package item URI for this part, commonly known as its part name,
        e.g. ``/ppt/slides/slide1.xml``
        """
        return self._partname

    @property
    def relationships(self):
        """
        Tuple of |Relationship| instances, each representing a relationship
        from this part to another part.
        """
        return tuple(self._relationships)

    def _load(self, fs, partname, ct_dict, parts_dict):
        """
        Load part identified as *partname* from filesystem *fs* and propagate
        the load to related parts.
        """
        # calculate working values
        baseURI = os.path.split(partname)[0]
        content_type = ct_dict[partname]

        # set persisted attributes
        self._partname = partname
        self.blob = fs.getblob(partname)
        self.typespec = PartTypeSpec(content_type)

        # load relationships and propagate load to target parts
        self._relationships = []  # discard any rels from prior load
        rel_elms = self.__get_rel_elms(fs)
        for rel_elm in rel_elms:
            rId = rel_elm.get('Id')
            reltype = rel_elm.get('Type')
            target_relpath = rel_elm.get('Target')
            target_partname = posixpath.abspath(posixpath.join(baseURI,
                                                               target_relpath))
            if target_partname in parts_dict:
                target_part = parts_dict[target_partname]
            else:
                target_part = Part()
                parts_dict[target_partname] = target_part
                target_part._load(fs, target_partname, ct_dict, parts_dict)

            # create relationship to target_part
            rel = Relationship(rId, self, reltype, target_part)
            self._relationships.append(rel)

        return self

    def _marshal(self, model_part, part_dict):
        """
        Load the contents of model-side part such that it can be saved to
        disk. Propagate marshalling to related parts.
        """
        # unpack working values
        content_type = model_part._content_type
        # assign persisted attributes from model part
        self._partname = model_part.partname
        self.blob = model_part._blob
        self.typespec = PartTypeSpec(content_type)

        # load relationships and propagate marshal to target parts
        for rel in model_part._relationships:
            # unpack working values for target part and relationship
            rId = rel._rId
            reltype = rel._reltype
            model_target_part = rel._target
            partname = model_target_part.partname
            # create package-part for target
            if partname in part_dict:
                part = part_dict[partname]
            else:
                part = Part()
                part_dict[partname] = part
                part._marshal(model_target_part, part_dict)
            # create marshalled version of relationship
            marshalled_rel = Relationship(rId, self, reltype, part)
            self._relationships.append(marshalled_rel)

    @property
    def _relsitem_element(self):
        nsmap = {None: pptx.spec.nsmap['pr']}
        element = etree.Element(qtag('pr:Relationships'), nsmap=nsmap)
        for rel in self._relationships:
            element.append(rel._element)
        return element

    @property
    def _relsitemURI(self):
        """
        Return theoretical package URI for this part's relationships item,
        without regard to whether this part actually has a relationships item.
        """
        head, tail = os.path.split(self._partname)
        return '%s/_rels/%s.rels' % (head, tail)

    def __get_rel_elms(self, fs):
        """
        Helper method for _load(). Return list of this relationship elements
        for this part from *fs*. Returns empty list if there are no
        relationships for this part, either because parts of this type never
        have relationships or its relationships are optional and none exist in
        this filesystem (package).
        """
        relsitemURI = self.__relsitemURI(self.typespec, self._partname, fs)
        if relsitemURI is None:
            return []
        if relsitemURI not in fs:
            tmpl = "required relationships item '%s' not found in package"
            raise CorruptedPackageError(tmpl % relsitemURI)
        root_elm = fs.getelement(relsitemURI)
        return root_elm.findall(qtag('pr:Relationship'))

    @staticmethod
    def __relsitemURI(typespec, partname, fs):
        """
        REFACTOR: Combine this logic into __get_rel_elms, it's the only caller
        and logic is partially redundant.

        Return package URI for this part's relationships item. Returns None if
        a part of this type never has relationships. Also returns None if a
        part of this type has only optional relationships and the package
        contains no rels item for this part.
        """
        if typespec.has_rels == PTS_HASRELS_NEVER:
            return None
        head, tail = os.path.split(partname)
        relsitemURI = '%s/_rels/%s.rels' % (head, tail)
        if typespec.has_rels == PTS_HASRELS_OPTIONAL:
            return relsitemURI if relsitemURI in fs else None
        return relsitemURI


class Relationship(object):
    """
    Return a new |Relationship| instance with local identifier *rId* that
    associates *source* with *target*. *source* is an instance of either
    |Package| or |Part|. *target* is always an instance of |Part|. Note that
    *rId* is only unique within the scope of *source*. Relationships do not
    have a globally unique identifier.

    The following attributes are available from |Relationship| instances:

    .. attribute:: rId

       The source-local identifier for this relationship.

    .. attribute:: reltype

       The relationship type URI for this relationship. These are defined in
       the ECMA spec and look something like:
       'http://schemas.openxmlformats.org/.../relationships/slide'

    .. attribute:: target

       The target :class:`pptx.packaging.Part` instance in this relationship.

    """
    def __init__(self, rId, source, reltype, target):
        super(Relationship, self).__init__()
        self._source = source
        self.rId = rId
        self.reltype = reltype
        self.target = target

    @property
    def _element(self):
        """
        The :class:`ElementTree._Element` instance containing the XML
        representation of this Relationship.
        """
        element = etree.Element('Relationship')
        element.set('Id', self.rId)
        element.set('Type', self.reltype)
        element.set('Target', self._target_relpath)
        return element

    @property
    def _baseURI(self):
        """Return the directory part of the source itemURI."""
        if isinstance(self._source, Part):
            return os.path.split(self._source.partname)[0]
        return PKG_BASE_URI

    @property
    def _target_relpath(self):
        # workaround for posixpath bug in 2.6, doesn't generate correct
        # relative path when *start* (second) parameter is root ('/')
        if self._baseURI == '/':
            relpath = self.target.partname[1:]
        else:
            relpath = posixpath.relpath(self.target.partname, self._baseURI)
        return relpath


class PartTypeSpec(object):
    """
    Return an instance of |PartTypeSpec| containing metadata for parts of type
    *content_type*. Instances are cached, so no more than one instance for a
    particular content type is in memory.

    Instances provide the following attributes:

    .. attribute:: content_type

       MIME type-like string that identifies how content is encoded for parts
       of this type. In most cases it corresponds to a particular XML
       sub-schema, although binary parts have other encoding schemes. As an
       example, the content type for a theme part is
       ``application/vnd.openxmlformats-officedocument.theme+xml``. Each
       part's content type is indicated in the content types item
       (``[Content_Types].xml``) located in the package root.

    .. attribute:: basename

       The root of the partname "filename" segment for parts of this type. For
       example, *basename* for slide layout parts is ``slideLayout`` and the
       partname for a slide layout matches the regular expression
       ``/ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``, e.g.
       ``/ppt/slideLayouts/slideLayout1.xml``.

       Note that while *basename* also usually corresponds to the base of the
       immediate parent "directory" name for tuple parts, this is not
       guaranteed. One example is theme parts, with partnames like
       ``/ppt/theme/theme1.xml``. Use *baseURI* to determine the "directory"
       portion of a partname.

    .. attribute:: ext

       The extension of the partname "filename" segment for parts of this
       type. For example, *ext* for a presentation part
       (``/ppt/presentation.xml``) is ``.xml``. Note that the leading period
       is included in the extension, consistent with the behavior of
       :func:`os.path.split`.

    .. attribute:: cardinality

       One of :attr:`pptx.spec.PTS_CARDINALITY_SINGLETON` or
       :attr:`pptx.spec.PTS_CARDINALITY_TUPLE`, corresponding to whether at
       most one or multiple parts of this type may appear in the package.
       ``/ppt/presentation.xml`` is an example of a singleton part.
       ``/ppt/slideLayouts/slideLayout4.xml`` is an example of a tuple part.
       The term *tuple* in this context is drawn from set theory in math and
       has no direct relationship to the Python tuple class.

    .. attribute:: required

       Boolean expressing whether at least one instance of this part type must
       appear in the package. ``presentation`` is an example of a required
       part type. ``notesMaster`` is an example of a optional part type.

    .. attribute:: baseURI

       The "directory" portion of the partname for parts of this type. The
       term *URI* is used because although part names (and other package item
       URIs) strongly resemble filesystem paths, and are readily operated on
       with functions from :mod:`os.path`, they have no direct correspondence
       to location in a file system (otherwise all packages would overwrite
       each other in the root directory :).

       For example, *baseURI* for slide layout parts is ``/ppt/slideLayouts``.

    .. attribute:: has_rels

       One of ``pptx.spec.PTS_HASRELS_ALWAYS``,
       ``pptx.spec.PTS_HASRELS_NEVER``, or ``pptx.spec.PTS_HASRELS_OPTIONAL``,
       indicating whether parts of this type always, never, or sometimes have
       relationships, respectively.

    .. attribute:: reltype

       The string used in the ``Type`` attribute of a ``Relationship`` XML
       element where a part of this content type is the target of the
       relationship. A relationship type is a URI string of the same form as a
       web page URL. For example, *reltype* for a part named
       ``/ppt/slides/slide1.xml`` would look something like
       ``http://schemas.openxmlformats.org/.../relationships/slide``.

    """
    _instances = {}

    def __new__(cls, content_type):
        """
        Only create new instance on first call for content_type. After that,
        use cached instance.
        """
        # if there's not a matching instance in the cache, create one
        if content_type not in cls._instances:
            inst = super(PartTypeSpec, cls).__new__(cls)
            cls._instances[content_type] = inst
        # return the instance; note that __init__() gets called either way
        return cls._instances[content_type]

    def __init__(self, content_type):
        """Initialize spec attributes from constant values in pptx.spec."""
        # skip loading if this instance is from the cache
        if hasattr(self, '_loaded'):
            return
        # otherwise initialize new instance
        self._loaded = True
        if content_type not in pptx.spec.pml_parttypes:
            tmpl = "no content type '%s' in pptx.spec.pml_parttypes"
            raise KeyError(tmpl % content_type)
        ptsdict = pptx.spec.pml_parttypes[content_type]
        # load attributes from spec constants dictionary
        # e.g. 'application/vnd.open...ment.presentationml.slideMaster+xml'
        self.content_type = content_type
        # e.g. 'slideMaster'
        self.basename = ptsdict['basename']
        # e.g. '.xml'
        self.cardinality = ptsdict['cardinality']
        # e.g. PTS_CARDINALITY_SINGLETON or PTS_CARDINALITY_TUPLE
        self.ext = ptsdict['ext']
        # e.g. False
        self.required = ptsdict['required']
        # e.g. '/ppt/slideMasters'
        self.baseURI = ptsdict['baseURI']
        # e.g. PTS_HASRELS_ALWAYS, PTS_HASRELS_NEVER, or PTS_HASRELS_OPTIONAL
        self.has_rels = ptsdict['has_rels']
        # e.g. 'http://schemas.openxmlformats.org/.../metadata/core-properties'
        self.reltype = ptsdict['reltype']

    @property
    def format(self):
        """One of ``'xml'`` or ``'binary'``."""
        return 'xml' if self.ext == '.xml' else 'binary'


# ============================================================================
# Support Classes
# ============================================================================

class _ContentTypesItem(object):
    """
    Lookup content type by part name using dictionary syntax, e.g.
    ``content_type = cti['/ppt/presentation.xml']``.
    """
    def __init__(self):
        super(_ContentTypesItem, self).__init__()
        self._defaults = {}
        self._overrides = {}

    def __getitem__(self, partname):
        """
        Return the content type for the part with *partname*.
        """
        # look for explicit partname match in overrides (case-insensitive)
        for override_partname, content_type in self._overrides.items():
            if override_partname.lower() == partname.lower():
                return content_type
        # look for case-insensitive match on extension in default element
        ext = os.path.splitext(partname)[1]  # get extension of partname
        # with leading dot trimmed off
        ext = ext[1:] if ext.startswith('.') else ext
        for extension, content_type in self._defaults.items():
            if extension.lower() == ext.lower():
                return content_type
        # if neither of those work, raise an exception
        tmpl = "no content type for part '%s' in [Content_Types].xml"
        raise LookupError(tmpl % partname)

    def __len__(self):
        """
        Return sum count of Default and Override elements.
        """
        return len(self._defaults) + len(self._overrides)

    def compose(self, parts):
        """
        Assemble a [Content_Types].xml item based on the contents of *parts*.
        """
        # extensions in this dict include leading '.'
        def_cts = pptx.spec.default_content_types
        # initialize working dictionaries for defaults and overrides
        self._defaults = dict((ext[1:], def_cts[ext])
                              for ext in ('.rels', '.xml'))
        self._overrides = {}
        # compose appropriate element for each part
        for part in parts:
            ext = os.path.splitext(part.partname)[1]
            # if extension is '.xml', assume an override. There might be a
            # fancier way to do this, otherwise I don't know what 'xml'
            # Default entry is for.
            if ext == '.xml':
                self._overrides[part.partname] = part.content_type
            elif ext in def_cts:
                self._defaults[ext[1:]] = def_cts[ext]
            else:
                tmpl = "extension '%s' not found in default_content_types"
                raise LookupError(tmpl % (ext))
        return self

    @property
    def element(self):
        nsmap = {None: pptx.spec.nsmap['ct']}
        element = etree.Element(qtag('ct:Types'), nsmap=nsmap)
        if self._defaults:
            for ext in sorted(self._defaults.keys()):
                subelm = etree.SubElement(element, qtag('ct:Default'))
                subelm.set('Extension', ext)
                subelm.set('ContentType', self._defaults[ext])
        if self._overrides:
            for partname in sorted(self._overrides.keys()):
                subelm = etree.SubElement(element, qtag('ct:Override'))
                subelm.set('PartName', partname)
                subelm.set('ContentType', self._overrides[partname])
        return element

    def load(self, fs):
        """
        Retrieve [Content_Types].xml from specified file system and load it.
        Returns a reference to this _ContentTypesItem instance to allow
        generative call, e.g. ``cti = _ContentTypesItem().load(fs)``.
        """
        element = fs.getelement('/[Content_Types].xml')
        defaults = element.findall(qtag('ct:Default'))
        overrides = element.findall(qtag('ct:Override'))
        self._defaults = dict((d.get('Extension'), d.get('ContentType'))
                              for d in defaults)
        self._overrides = dict((o.get('PartName'), o.get('ContentType'))
                               for o in overrides)
        return self


# ============================================================================
# FileSystem Classes
# ============================================================================

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

    def write_blob(self, blob, itemURI):
        """
        Write *blob* to zip file as binary stream named *itemURI*.
        """
        if itemURI in self:
            tmpl = "Item with URI '%s' already in package"
            raise DuplicateKeyError(tmpl % itemURI)
        membername = itemURI[1:]  # trim off leading slash
        self.zipf.writestr(membername, blob)

    def write_element(self, element, itemURI):
        """
        Write *element* to zip file as an XML document named *itemURI*.
        """
        if itemURI in self:
            tmpl = "Item with URI '%s' already in package"
            raise DuplicateKeyError(tmpl % itemURI)
        membername = itemURI[1:]  # trim off leading slash
        xml = etree.tostring(element, encoding='UTF-8', pretty_print=False,
                             standalone=True)
        self.zipf.writestr(membername, xml)
