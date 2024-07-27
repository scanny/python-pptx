# pyright: reportPrivateUsage=false

from __future__ import annotations

from typing import Literal, overload

from .._types import _ElementOrTree, _OutputMethodArg
from ..etree import HTMLParser, XMLParser
from ._element import _Element

def fromstring(text: str | bytes, parser: XMLParser | HTMLParser) -> _Element: ...

# -- Native str, no XML declaration --
@overload
def tostring(  # type: ignore[overload-overlap]
    element_or_tree: _ElementOrTree,
    *,
    encoding: type[str] | Literal["unicode"],
    method: _OutputMethodArg = "xml",
    pretty_print: bool = False,
    with_tail: bool = True,
    standalone: bool | None = None,
    doctype: str | None = None,
) -> str: ...

# -- bytes, str encoded with `encoding`, no XML declaration --
@overload
def tostring(
    element_or_tree: _ElementOrTree,
    *,
    encoding: str | None = None,
    method: _OutputMethodArg = "xml",
    xml_declaration: bool | None = None,
    pretty_print: bool = False,
    with_tail: bool = True,
    standalone: bool | None = None,
    doctype: str | None = None,
) -> bytes: ...
