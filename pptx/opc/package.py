# encoding: utf-8

"""
The :mod:`pptx.packaging` module coheres around the concerns of reading and
writing presentations to and from a .pptx file.
"""

from __future__ import absolute_import


class PartFactory(object):
    """
    Provides a way for client code to specify a subclass of |Part| to be
    constructed by |Unmarshaller| based on its content type.
    """
    part_type_for = {}
    # default_part_type = Part
    default_part_type = None

    def __new__(cls, partname, content_type, blob):
        if content_type in cls.part_type_for:
            CustomPartClass = cls.part_type_for[content_type]
            return CustomPartClass.load(partname, content_type, blob)
        return cls.default_part_type(partname, content_type, blob)


class Unmarshaller(object):
    """
    Hosts static methods for unmarshalling a package from a |PackageReader|
    instance.
    """
    @staticmethod
    def unmarshal(pkg_reader, pkg, part_factory):
        """
        Construct graph of parts and realized relationships based on the
        contents of *pkg_reader*, delegating construction of each part to
        *part_factory*. Package relationships are added to *pkg*.
        """
        parts = Unmarshaller._unmarshal_parts(pkg_reader, part_factory)
        Unmarshaller._unmarshal_relationships(pkg_reader, pkg, parts)
        for part in parts.values():
            part.after_unmarshal()
        pkg.after_unmarshal()

    @staticmethod
    def _unmarshal_parts(pkg_reader, part_factory):
        """
        Return a dictionary of |Part| instances unmarshalled from
        *pkg_reader*, keyed by partname. Side-effect is that each part in
        *pkg_reader* is constructed using *part_factory*.
        """
        parts = {}
        for partname, content_type, blob in pkg_reader.iter_sparts():
            parts[partname] = part_factory(partname, content_type, blob)
        return parts

    @staticmethod
    def _unmarshal_relationships(pkg_reader, pkg, parts):
        """
        Add a relationship to the source object corresponding to each of the
        relationships in *pkg_reader* with its target_part set to the actual
        target part in *parts*.
        """
        for source_uri, srel in pkg_reader.iter_srels():
            source = pkg if source_uri == '/' else parts[source_uri]
            target = (srel.target_ref if srel.is_external
                      else parts[srel.target_partname])
            source._add_relationship(srel.reltype, target, srel.rId,
                                     srel.is_external)
