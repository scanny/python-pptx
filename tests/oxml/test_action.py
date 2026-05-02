"""Unit-test suite for `pptx.oxml.action` module."""

from __future__ import annotations

from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

_A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


class DescribeCT_Hyperlink(object):
    """Unit-test suite for `pptx.oxml.action.CT_Hyperlink` suppress_theme_color accessor."""

    def it_knows_when_suppress_theme_color_marker_is_absent(self):
        hlinkClick = parse_xml(f"<a:hlinkClick {_A_NS}/>")
        assert hlinkClick.suppress_theme_color is False

    def it_knows_when_suppress_theme_color_marker_is_present(self):
        hlinkClick = parse_xml(
            f"<a:hlinkClick {_A_NS}>"
            '<a:extLst><a:ext uri="{PY-PPTX-940}"/></a:extLst>'
            "</a:hlinkClick>"
        )
        assert hlinkClick.suppress_theme_color is True

    def it_ignores_a_foreign_ext_child(self):
        hlinkClick = parse_xml(
            f"<a:hlinkClick {_A_NS}>"
            '<a:extLst><a:ext uri="{OTHER-EXT-UUID}"/></a:extLst>'
            "</a:hlinkClick>"
        )
        assert hlinkClick.suppress_theme_color is False

    def it_can_add_the_suppress_theme_color_marker(self):
        hlinkClick = parse_xml(f"<a:hlinkClick {_A_NS}/>")
        hlinkClick.suppress_theme_color = True
        assert hlinkClick.suppress_theme_color is True
        ext = hlinkClick.find(qn("a:extLst")).find(qn("a:ext"))
        assert ext.get("uri") == "{PY-PPTX-940}"

    def it_is_idempotent_when_adding_the_marker_twice(self):
        hlinkClick = parse_xml(f"<a:hlinkClick {_A_NS}/>")
        hlinkClick.suppress_theme_color = True
        hlinkClick.suppress_theme_color = True
        extLst = hlinkClick.find(qn("a:extLst"))
        exts = extLst.findall(qn("a:ext"))
        assert len(exts) == 1

    def it_can_remove_the_suppress_theme_color_marker(self):
        hlinkClick = parse_xml(f"<a:hlinkClick {_A_NS}/>")
        hlinkClick.suppress_theme_color = True
        hlinkClick.suppress_theme_color = False
        assert hlinkClick.suppress_theme_color is False
        # -- empty extLst was removed too --
        assert hlinkClick.find(qn("a:extLst")) is None

    def it_preserves_foreign_extensions_when_removing_marker(self):
        hlinkClick = parse_xml(
            f"<a:hlinkClick {_A_NS}>"
            '<a:extLst>'
            '<a:ext uri="{OTHER-EXT-UUID}"/>'
            '<a:ext uri="{PY-PPTX-940}"/>'
            "</a:extLst>"
            "</a:hlinkClick>"
        )
        hlinkClick.suppress_theme_color = False
        extLst = hlinkClick.find(qn("a:extLst"))
        assert extLst is not None
        exts = extLst.findall(qn("a:ext"))
        assert len(exts) == 1
        assert exts[0].get("uri") == "{OTHER-EXT-UUID}"
