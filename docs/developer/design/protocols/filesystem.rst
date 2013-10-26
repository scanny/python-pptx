=================================
FileSystem interface and protocol
=================================

The FileSystem interface and protocol provides access to packages conforming
to the Open Packaging Convention (OPC) specified in ISO/IEC 29500 Part 2.

OPC packages may take different forms, for example the ISO spec mentions the
possibility of streaming-access network-based packages, but the only
file-based packages are dealt with in |pp|. Most commonly these are ``.pptx``
files, which are a zip format file, but |pp| also supports reading from (but
not saving to) an expanded form, which is essentially the same directory
structure you get when you unzip a ``.pptx`` file. This can be handy when you
want to swap out parts of a template without having to zip up the directory
structure every time you make a change.

The FileSystem interface encapsulates the details of accessing and writing
individual items in a package, raising the level of abstraction to a namespace
of item URIs. A partname, e.g. ``/ppt/slides/slide1.xml`` is a package item
URI, as is ``/ppt/slides/_rels/slide1.xml.rels``. Note that the term
*filesystem* here does not correspond directly to the conventional meaning of
the entire set of files and directories on a particular computer. Rather it is
meant to reflect an abstraction of how the individual streams that constitute
a package are assembled as a whole.

FileSystem objects
==================

Filesystem services are provided by a collaboration of four classes:

* |FileSystem| is a factory class which returns an appropriate instance of
  either |DirectoryFileSystem| or |ZipFileSystem| based on the provided path.
* |BaseFileSystem| is inherited by both |DirectoryFileSystem| and
  |ZipFileSystem| and provides methods that are common to both of them.
* |DirectoryFileSystem| provides services on a package that has been expanded
  into a directory hierarchy.
* |ZipFileSystem| provides services on a zipfile package, typically a ``.pptx``
  file.

The following code sample illustrates the protocol::

    from pptx.opc.packaging import FileSystem
    
    # construction
    fs = FileSystem(path)
    
    # retrieve an XML part as etree.Element instance
    element = fs.getelement('/[Content_Types].xml')
    
    # TODO: Add others, new(), write, etc. ...


|FileSystem| objects
--------------------

.. autoclass:: pptx.opc.packaging.FileSystem
   :members: __new__
   :member-order: bysource
   :undoc-members:


|BaseFileSystem| objects
------------------------

.. autoclass:: pptx.opc.packaging.BaseFileSystem
   :members:
   :member-order: bysource
   :undoc-members:


|DirectoryFileSystem| objects
-----------------------------

.. autoclass:: pptx.opc.packaging.DirectoryFileSystem
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:


|ZipFileSystem| objects
-----------------------

.. autoclass:: pptx.opc.packaging.ZipFileSystem
   :members:
   :member-order: bysource
   :undoc-members:
   :show-inheritance:

