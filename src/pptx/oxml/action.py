"""lxml custom element classes for click-action (hyperlink) elements."""

from __future__ import annotations

from typing import Callable

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, OxmlElement

# -- URI used by python-pptx to mark a run as opting out of the theme's
# -- hyperlink-color override. PowerPoint ignores unknown extensions, so the
# -- marker is lossless across round-trips through python-pptx and does not
# -- confuse PowerPoint. See `pptx.text.text.Font.use_theme_hyperlink_color`
# -- for the API that reads and writes this marker. The URI is rooted under
# -- the python-pptx project domain so it won't collide with any Microsoft
# -- or third-party extension.
_NO_THEME_HLINK_COLOR_URI = "{PY-PPTX-940}"
from pptx.oxml.simpletypes import ST_RelationshipId, XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)


class CT_EmbeddedWAVAudioFile(BaseOxmlElement):
    """`a:snd` element.

    Child of `a:hlinkClick` / `a:hlinkHover` (or `a:hlinkMouseOver` on text runs) that
    specifies an embedded WAV-audio file to be played when the containing hyperlink is
    activated. See `CT_EmbeddedWAVAudioFile` in the ISO/IEC 29500-1 DrawingML schema.
    """

    rEmbed: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "r:embed", ST_RelationshipId
    )
    name: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString, default=""
    )


class CT_Hyperlink(BaseOxmlElement):
    """Custom element class for <a:hlinkClick> elements."""

    get_or_add_snd: Callable[[], CT_EmbeddedWAVAudioFile]

    _tag_seq = ("a:snd", "a:extLst")
    snd: CT_EmbeddedWAVAudioFile | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:snd", successors=_tag_seq[1:]
    )
    del _tag_seq

    rId: str = OptionalAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]
    action: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "action", XsdString
    )
    tooltip: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "tooltip", XsdString
    )

    @property
    def action_fields(self) -> dict[str, str]:
        """Query portion of the `ppaction://` URL as dict.

        For example `{'id':'0', 'return':'true'}` in 'ppaction://customshow?id=0&return=true'.

        Returns an empty dict if the URL contains no query string or if no action attribute is
        present.
        """
        url = self.action

        if url is None:
            return {}

        halves = url.split("?")
        if len(halves) == 1:
            return {}

        key_value_pairs = halves[1].split("&")
        return dict([pair.split("=") for pair in key_value_pairs])

    @property
    def action_verb(self) -> str | None:
        """The host portion of the `ppaction://` URL contained in the action attribute.

        For example 'customshow' in 'ppaction://customshow?id=0&return=true'. Returns |None| if no
        action attribute is present.
        """
        url = self.action

        if url is None:
            return None

        protocol_and_host = url.split("?")[0]
        host = protocol_and_host[11:]

        return host

    # -- hyperlink-color override marker (issue #940) ------------------

    @property
    def suppress_theme_color(self) -> bool:
        """True when a python-pptx `no-theme-hlink-color` marker is present under this element.

        The marker is an `a:extLst/a:ext` child whose `uri` attribute is the python-pptx
        sentinel value `{PY-PPTX-940}`. PowerPoint ignores unrecognized `a:ext` extensions,
        so the marker is harmless cosmetically; it is used by python-pptx to round-trip the
        caller's intent that the run's explicit `a:solidFill` color should be preferred over
        the theme's `a:hlink` color.
        """
        return self._find_no_theme_ext() is not None

    @suppress_theme_color.setter
    def suppress_theme_color(self, value: bool) -> None:
        ext = self._find_no_theme_ext()
        if value:
            if ext is None:
                self._add_no_theme_ext()
        else:
            if ext is not None:
                extLst = ext.getparent()
                assert extLst is not None  # -- for pyright --
                extLst.remove(ext)
                # -- remove the extLst container when it's empty --
                if len(extLst) == 0:
                    parent = extLst.getparent()
                    assert parent is not None
                    parent.remove(extLst)

    def _find_no_theme_ext(self):
        """Return the `a:ext` element marking theme-color suppression, or None."""
        extLst = self.find(qn("a:extLst"))
        if extLst is None:
            return None
        for ext in extLst.findall(qn("a:ext")):
            if ext.get("uri") == _NO_THEME_HLINK_COLOR_URI:
                return ext
        return None

    def _add_no_theme_ext(self):
        """Append an `a:extLst/a:ext` marker indicating theme hyperlink-color is suppressed."""
        extLst = self.find(qn("a:extLst"))
        if extLst is None:
            extLst = OxmlElement("a:extLst")
            self.append(extLst)
        ext = OxmlElement("a:ext")
        ext.set("uri", _NO_THEME_HLINK_COLOR_URI)
        extLst.append(ext)
        return ext
