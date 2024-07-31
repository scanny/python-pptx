# pyright: reportPrivateUsage=false

from __future__ import annotations

from typing import Literal

from ._classlookup import ElementClassLookup
from .._types import _ET_co, _NSMapArg, _TagName, SupportsLaxedItems

class HTMLParser:
    def __init__(
        self,
        *,
        encoding: str | None = None,
        remove_blank_text: bool = False,
        remove_comments: bool = False,
        remove_pis: bool = False,
        strip_cdata: bool = True,
        no_network: bool = True,
        recover: bool = True,
        compact: bool = True,
        default_doctype: bool = True,
        collect_ids: bool = True,
        huge_tree: bool = False,
    ) -> None: ...
    def set_element_class_lookup(self, lookup: ElementClassLookup | None = None) -> None: ...

class XMLParser:
    def __init__(
        self,
        *,
        attribute_defaults: bool = False,
        collect_ids: bool = True,
        compact: bool = True,
        dtd_validation: bool = False,
        encoding: str | None = None,
        huge_tree: bool = False,
        load_dtd: bool = False,
        no_network: bool = True,
        ns_clean: bool = False,
        recover: bool = False,
        remove_blank_text: bool = False,
        remove_comments: bool = False,
        remove_pis: bool = False,
        resolve_entities: bool | Literal["internal"] = "internal",
        strip_cdata: bool = True,
    ) -> None: ...
    def makeelement(
        self,
        _tag: _TagName,
        /,
        attrib: SupportsLaxedItems[str, str] | None = None,
        nsmap: _NSMapArg | None = None,
        **_extra: str,
    ) -> _ET_co: ...
    def set_element_class_lookup(self, lookup: ElementClassLookup | None = None) -> None:
        """
        Notes
        -----
        When calling this method, it is advised to also change typing
        specialization of concerned parser too, because current python
        typing system can't change it automatically.

        Example
        -------
        Following code demonstrates how to create ``lxml.html.HTMLParser``
        manually from ``lxml.etree.HTMLParser``::

        ```python
        parser = etree.HTMLParser()
        reveal_type(parser)  # HTMLParser[_Element]
        if TYPE_CHECKING:
            parser = cast('etree.HTMLParser[HtmlElement]', parser)
        else:
            parser.set_element_class_lookup(
                html.HtmlElementClassLookup())
        result = etree.fromstring(data, parser=parser)
        reveal_type(result)  # HtmlElement
        ```
        """
        ...
