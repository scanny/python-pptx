
Understanding ``xmlchemy``
==========================

``xmlchemy`` is an object-XML mapping layer somewhat reminiscent of
SQLAlchemy, hence the name. Mapping XML elements to objects is not nearly as
challenging as mapping to a relational database, so this layer is
substantially more modest. Nevertheless, it provides a powerful and very
useful abstraction layer around ``lxml``, particularly well-suited to
providing access to the broad schema of XML elements involved in the Open XML
standard.


Additional topics to add ...
----------------------------

* understanding complex types in Open XML
* understanding attribute definitions Open XML
* understanding simple types in Open XML


Adding support for a new element type
-------------------------------------

* add a new custom element mapping to ``pptx.oxml.__init__``
* add a new custom element class in the appropriate ``pptx.oxml`` subpackage
  module
* Add element definition members to the class
* Add attribute definition members to the class
* Add simple type definitions to ``pptx.oxml.simpletype``


Example
-------

::

    from pptx.oxml.xmlchemy import BaseOxmlElement


    class CT_Foobar(BaseOxmlElement):
        """
        Custom element class corresponding to ``CT_Foobar`` complex type
        definition in pml.xsd or other Open XML schema.
        """
        hlink = ZeroOrOne('a:hlink', successors=('a:rtl', 'a:extLst'))
        eg_fillProperties = ZeroOrOneChoice(
            (Choice('a:noFill'), Choice('a:solidFill'), Choice('a:gradFill'),
             Choice('a:blipFill'), Choice('a:pattFill')),
            successors=(
                'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx',
                'a:uLn', 'a:uFillTx' 'a:extLst'
            )
        )

        sz = OptionalAttribute(
            'sz', ST_SimpleType, default=ST_SimpleType.OPTION
        )
        anchor = OptionalAttribute('i', XsdBoolean)
        rId = RequiredAttribute('r:id', XsdString)


Protocol
--------

::

    >>> assert isinstance(foobar, CT_Foobar)
    >>> foobar.hlink
    None
    >>> hlink = foobar._add_hlink()
    >>> hlink
    <pptx.oxml.xyz.CT_Hyperlink object at 0x10ab4b2d0>
    >>> assert foobar.hlink is hlink

    >>> foobar.eg_fillProperties
    None
    >>> foobar.solidFill
    None
    >>> solidFill = foobar.get_or_change_to_solidFill()
    >>> solidFill
    <pptx.oxml.xyz.CT_SolidFill object at 0x10ab4b2d0>
    >>> assert foobar.eg_fillProperties is solidFill
    >>> assert foobar.solidFill is solidFill
    >>> foobar.remove_eg_fillProperties()
    >>> foobar.eg_fillProperties
    None
    >>> foobar.solidFill
    None


``OneAndOnlyOne`` element declaration
-------------------------------------

The ``OneAndOnlyOne`` callable generates the API for a required child
element::

    childElement = OneAndOnlyOne('ns:localTagName')

Unlike the other element declarations, the call does not include
a ``successors`` argument. Since no API for inserting a new element is
generated, a successors list is not required.


Generated API
~~~~~~~~~~~~~

``childElement`` property (read-only)
    Holds a reference to the child element object. Raises |InvalidXmlError|
    on access if the required child element is not present.


Protocol
~~~~~~~~

::

    >>> foobar.childElement
    <pptx.oxml.xyz.CT_ChildElement object at 0x10ab4b2d0>


``RequiredAttribute`` attribute declaration
-------------------------------------------

::

    reqAttr = RequiredAttribute('reqAtr', ST_SimpleType)


Generated API
~~~~~~~~~~~~~

``childElement`` property (read/write)
    Referencing the property returns the type-converted value of the
    attribute as determined by the ``from_xml()`` method of the simple type
    class appearing in the declaration (e.g. ST_SimpleType above).
    Assignments to the property are validated by the ``validate()`` method of
    the simple type class, potentially raising ``TypeError`` or
    ``ValueError``. Values are assigned in their natural Python type and are
    encoded to the appropriate string value by the ``to_xml()`` method of the
    simple type class.


``ZeroOrOne`` element declaration
---------------------------------

::

    childElement = ZeroOrOne(
        'ns:localTagName', successors=('ns:abc', 'ns:def')
    )


Generated API
~~~~~~~~~~~~~

``childElement`` property (read-only)
    Holds a reference to the child element object, or None if the element is
    not present.

``get_or_add_childElement()`` method
    Returns the child element object, newly added if not present.

``_add_childElement()`` empty element adder method
    Returns a newly added empty child element having the declared tag name.
    Adding is unconditional and assumes the element is not already present.
    This method is called by the ``get_or_add_childElement()`` method as
    needed and may be called by a hand-coded ``add_childElement()`` method
    as needed. May be overridden to produce customized behavior.

``_new_childElement()`` empty element creator method
    Returns a new "loose" child element of the declared tag name. Called by
    ``_add_childElement()`` to obtain a new child element, it may be
    overridden to customize the element creation process.

``_insert_childElement(childElement)`` element inserter method
    Returns the passed ``childElement`` after inserting it before any
    successor elements, as listed in the ``successors`` argument of the
    declaration. Called by ``_add_childElement()`` to insert the new element
    it creates using ``_new_childElement()``.

``_remove_childElement()`` element remover method
    Removes all instances of the child element. Does not raise an error if no
    matching child elements are present.
