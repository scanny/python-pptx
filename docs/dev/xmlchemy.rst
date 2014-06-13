
Understanding ``xmlchemy``
==========================

``xmlchemy`` is an object-XML mapping layer somewhat reminiscent of
SQLAlchemy, hence the name. Mapping XML elements to objects is not nearly as
challenging as mapping to a relational database, so this layer is
substantially more modest. Nevertheless, it provides a powerful and very
useful abstraction layer around ``lxml``, particularly well-suited to
providing access to the broad schema of XML elements involved in the Open XML
standard.


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


``ZeroOrOne`` element declarator
--------------------------------

::

    childElement = ZeroOrOne(
        'ns:localTagName', successors=('ns:abc', 'ns:def')
    )


**Generated API**

* ``childElement`` property (read-only)
* ``get_or_add_childElement()`` method
* ``_add_childElement()`` empty element adder method
* ``_new_childElement()`` empty element creator method
* ``_insert_childElement(childElement)`` element inserter method
* ``_remove_childElement()`` element remover method
