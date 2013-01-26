===============================
ElementTree manipulation domain
===============================

One of the domains in this code base is the *ElementTree manipulation domain*.

It will be well to encapsulate operations in this domain. ... the following
*smells* can be detected ...

1. Code to locate a particular child element within a local ElementTree graph
   is tedious, repetitive, and tends to obscure intent by taking up too many
   lines to get a single thing done.::

     xpath = './a:endParaRPr'
     elms = self.__p.xpath(xpath, namespaces=self.__nsmap)
     if len(elms):
         endParaRPr = elms[0]
         endParaRPr.addprevious(r)
     else:
         self.__p.append(r)

2. Working out that an optional element is not there and then working out
   where in the hierarchy to insert a new one is also repetitive and takes
   several lines that obscure the surrounding code.

3. Collection operations (add, delete, reorder, etc.) need to happen in both
   an XML domain collection and a user domain collection, with care taken to
   make sure they remain in sync. Not sure if its possible to link them in a
   way that makes operations on one happen transparently in the other.

4. ... others ...


Refactoring ideas
=================

1. **Bookmarks**. Introduce the notion of a *bookmark property* that provides
   ready access to a particular child element.

   :func:`_child` utility function factors bookmark properties down to a
   single line::

    def _child(element, child_tagname, nsmap):
        """
        Return direct child of *element* having *child_tagname* or
        :class:`None` if no such child element is present.
        """
        xpath = './%s' % child_tagname
        matching_children = element.xpath(xpath, namespaces=nsmap)
        return matching_children[0] if len(matching_children) else None
    
    # ...
    
    @property
    def __sldIdLst(self):
        """Bookmark to ``<p:sldIdLst>`` child element"""
        return _child(self._element, 'p:sldIdLst', self._nsmap)

#. **Wrapper object**. ... could develop a wrapper (decorator?) for main
   element that provides properties and methods that allow for convenient
   access. Some schema-like **metadata** could be added as state, things like
   child sequence and maybe cardinality for example, to allow it to figure
   most things out for itself.

   ... perhaps a sequence of possible child elements, and insert can traverse
   the sequence to find first one that's not None or append to parent if all
   None::

    child_sequence = ('p:sldMasterIdLst', 'p:notesMasterIdLst',
                      'p:handoutMasterIdLst', 'p:sldIdLst', ...)

#. **Encapsulate and compose**. ... wrapper object could be a separate class
   for each element that needed wrapping, inherit from a base wrapper ... and
   that object held as a member in the API class. Maybe those XML wrapper
   objects could be segregated into a separate module.

   Might be mostly like an unmarshaler, but without the render_xml() method
   because it works on a live ElementTree graph. Could possibly incrementally
   refine that until it could handle a whole part, then might be able to
   postpone marshaling until _blob() call.

#. ... some way to add missing elements in-line, like self.__elm(add=true) to
   change behavior on access to a missing element ... substitute for::

    if self.__elm is None:
        self.__add_elm()
    
    # maybe this would suffice
    elm = (self.__elm if self.__elm is not None
                      else self.__add_elm())

#. ... perhaps some way to **lazy-add** and **auto-delete** an optional
   element ... auto-delete would mean getting rid of the element if it no
   longer had any children, that sort of thing.

