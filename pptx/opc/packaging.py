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

from __future__ import absolute_import

import os
import posixpath

from lxml import etree

import pptx.spec

from pptx.exceptions import CorruptedPackageError
from pptx.oxml.core import serialize_part_xml
from pptx.oxml.ns import nsuri, qn
from pptx.spec import PTS_HASRELS_NEVER, PTS_HASRELS_OPTIONAL

from .packuri import CONTENT_TYPES_URI, PackURI, PACKAGE_URI
from .phys_pkg import FileSystem, PhysPkgWriter


class Package(object):
    """
    Return a new package instance. Package is initially empty, call
    :meth:`open` to open an on-disk package or ``marshal()`` followed by
    ``save()`` to save an in-memory Office document.
    """

    PKG_RELSITEM_URI = PackURI('/_rels/.rels')

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
                         .findall(qn('pr:Relationship'))
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
            rId = rel.rId
            reltype = rel.reltype
            model_part = rel.target
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
        phys_pkg_writer = PhysPkgWriter(file)

        # write [Content_Types].xml
        cti = _ContentTypesItem().compose(self.parts)
        phys_pkg_writer.write(CONTENT_TYPES_URI, cti.xml)

        # write pkg rels item
        phys_pkg_writer.write(self.PKG_RELSITEM_URI, self._relsitem_xml)

        # write part items and their rels item when they have one
        for part in self.parts:
            partname = PackURI(part.partname)
            phys_pkg_writer.write(partname, part.blob)
            # write rels item if part has one
            if part.relationships:
                relsitemURI = PackURI(part._relsitemURI)
                phys_pkg_writer.write(relsitemURI, part._relsitem_xml)

        phys_pkg_writer.close()

    @property
    def _relsitem_element(self):
        nsmap = {None: nsuri('pr')}
        element = etree.Element(qn('pr:Relationships'), nsmap=nsmap)
        for rel in self._relationships:
            element.append(rel._element)
        return element

    @property
    def _relsitem_xml(self):
        return serialize_part_xml(self._relsitem_element)

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
        rel_elms = self._get_rel_elms(fs)
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
            rId = rel.rId
            reltype = rel.reltype
            model_target_part = rel.target
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
        nsmap = {None: nsuri('pr')}
        element = etree.Element(qn('pr:Relationships'), nsmap=nsmap)
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

    @property
    def _relsitem_xml(self):
        return serialize_part_xml(self._relsitem_element)

    def _get_rel_elms(self, fs):
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
        return root_elm.findall(qn('pr:Relationship'))

    @staticmethod
    def __relsitemURI(typespec, partname, fs):
        """
        REFACTOR: Combine this logic into _get_rel_elms, it's the only caller
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
        return PACKAGE_URI

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
        nsmap = {None: nsuri('ct')}
        element = etree.Element(qn('ct:Types'), nsmap=nsmap)
        if self._defaults:
            for ext in sorted(self._defaults.keys()):
                subelm = etree.SubElement(element, qn('ct:Default'))
                subelm.set('Extension', ext)
                subelm.set('ContentType', self._defaults[ext])
        if self._overrides:
            for partname in sorted(self._overrides.keys()):
                subelm = etree.SubElement(element, qn('ct:Override'))
                subelm.set('PartName', partname)
                subelm.set('ContentType', self._overrides[partname])
        return element

    @property
    def xml(self):
        return serialize_part_xml(self.element)

    def load(self, fs):
        """
        Retrieve [Content_Types].xml from specified file system and load it.
        Returns a reference to this _ContentTypesItem instance to allow
        generative call, e.g. ``cti = _ContentTypesItem().load(fs)``.
        """
        element = fs.getelement('/[Content_Types].xml')
        defaults = element.findall(qn('ct:Default'))
        overrides = element.findall(qn('ct:Override'))
        self._defaults = dict((d.get('Extension'), d.get('ContentType'))
                              for d in defaults)
        self._overrides = dict((o.get('PartName'), o.get('ContentType'))
                               for o in overrides)
        return self
