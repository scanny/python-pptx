##################
Security model
##################

This page documents the security assumptions and defenses applied to
python-pptx. It is aimed at developers integrating python-pptx into larger
systems that may need to read or write presentations derived from untrusted
sources.

Trust model
===========

python-pptx is primarily an authoring tool for legitimate PowerPoint
documents. The library is frequently used to read files as well, however, and
some of those files may originate from untrusted sources (user uploads,
third-party services, etc.). python-pptx applies the mitigations described
below so that *parsing a file* does not expose the host to the most common
classes of XML and zip attacks, but **a .pptx file is still a complex data
structure and should be treated as untrusted input**. In particular:

* python-pptx does not evaluate macros, scripts, linked content, or OLE
  payloads and makes no representation about the safety of content it
  exposes to the caller.
* Users embedding python-pptx in a server-side service are encouraged to
  run presentation parsing in an isolated process with OS-level memory and
  CPU limits (``ulimit``, containers, or a subprocess pool).

XML parser hardening
====================

All XML parsed by python-pptx is handled by an lxml ``XMLParser`` configured
with:

* ``resolve_entities=False`` - blocks both the "billion laughs" exponential
  entity-expansion attack and XML external entity ("XXE") attacks that
  would otherwise read local files or open network sockets in the course
  of resolving a ``SYSTEM`` or ``PUBLIC`` identifier.
* ``no_network=True`` - forbids network access for any related files.
* ``load_dtd=False`` - disables DTD loading. python-pptx never relies on
  DTD-declared defaults, so this has no functional impact.
* ``huge_tree=False`` - preserves libxml2's built-in limits on tree depth
  and text-node length.

The parser lives at :data:`pptx.oxml.oxml_parser` and is shared by every
call that ultimately reaches :func:`pptx.oxml.parse_xml`. A second, similarly
hardened parser (``pptx.chart.xlsx._xlsx_xml_parser``) is used when reading
the workbook embedded in a chart part.

Zip "bomb" defense
==================

A single ``.pptx`` file is a Zip archive. python-pptx loads every member of
the archive into memory when the package is opened. To prevent a crafted
archive from requesting an unreasonable allocation, the reader inspects the
uncompressed-size fields recorded in the Zip central directory before
reading any member bytes and raises :class:`pptx.exc.PackageTooLargeError`
when the declared total exceeds a configurable limit.

The default limit is **2 GiB**, chosen to leave sufficient headroom for
large legitimate presentations (hundreds of high-resolution images,
embedded video, etc.) while still blocking archives that declare petabyte-
scale payloads.

Two mechanisms are available for adjusting the limit:

* Set the ``PPTX_MAX_UNCOMPRESSED_SIZE`` environment variable to the
  desired limit in bytes (``0`` disables the guard).
* Assign to ``pptx.opc.serialized.MAX_UNCOMPRESSED_PACKAGE_SIZE`` before
  opening the package.

python-pptx **does not extract zip members to disk**. Member names are
used only as dictionary keys inside an in-memory mapping, so classic
"zip-slip" path-traversal attacks are not applicable.

Reporting a vulnerability
=========================

Please do *not* open a public GitHub issue for suspected security
vulnerabilities. Instead, contact the maintainer by email using the
address on the https://github.com/scanny profile page. Reports that
include a minimal reproducer and an assessment of the user impact will
be triaged most quickly.
