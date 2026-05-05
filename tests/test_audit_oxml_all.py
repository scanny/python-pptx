"""Regression test for audit item #5 (seal oxml public surface).

Asserts that ``from pptx.oxml import *`` exposes only a small, curated set
of public names. Historically the ``pptx.oxml`` package bound ~235 ``CT_*``
classes at module scope as a side-effect of its element-class registration
imports, meaning ``from pptx.oxml import *`` would drag the entire internal
XML-binding machinery into the caller's namespace. The explicit ``__all__``
on ``pptx.oxml`` restricts the wildcard surface to the intentionally-public
helpers (``parse_xml``, ``qn``, etc.).

This test locks that behaviour in so future additions to the side-effect
registration block can't accidentally re-expand the wildcard surface.
"""

from __future__ import annotations


def test_pptx_wildcard_exports_only_Presentation():
    """``from pptx import *`` exposes only the documented public factory."""
    ns: dict[str, object] = {}
    exec("from pptx import *", ns)  # noqa: S102 - intentional wildcard probe
    public = {n for n in ns if not n.startswith("_")}
    assert public == {"Presentation"}


def test_pptx_oxml_wildcard_does_not_leak_CT_or_ST_names():
    """``from pptx.oxml import *`` must not drag in CT_*/ST_* internals."""
    ns: dict[str, object] = {}
    exec("from pptx.oxml import *", ns)  # noqa: S102 - intentional wildcard probe
    public = {n for n in ns if not n.startswith("_")}
    ct_leaks = {n for n in public if n.startswith("CT_")}
    st_leaks = {n for n in public if n.startswith("ST_")}
    assert ct_leaks == set(), f"CT_* leaked via `from pptx.oxml import *`: {sorted(ct_leaks)}"
    assert st_leaks == set(), f"ST_* leaked via `from pptx.oxml import *`: {sorted(st_leaks)}"


def test_pptx_oxml_wildcard_exposes_curated_helpers():
    """The curated ``pptx.oxml`` wildcard surface is small and stable."""
    ns: dict[str, object] = {}
    exec("from pptx.oxml import *", ns)  # noqa: S102 - intentional wildcard probe
    public = {n for n in ns if not n.startswith("_")}
    # -- Required helpers that have always been part of the public `pptx.oxml`
    # -- wildcard surface and that real callers rely on.
    required = {
        "NamespacePrefixedTag",
        "oxml_parser",
        "parse_from_template",
        "parse_xml",
        "qn",
        "register_element_cls",
    }
    assert required <= public, f"missing expected public names: {sorted(required - public)}"
    # -- Upper bound: nothing outside this curated set should appear. If this
    # -- ever fails, consider whether the newcomer really belongs in the
    # -- wildcard surface or whether `__all__` should stay as-is.
    assert public == required, f"unexpected wildcard surface: {sorted(public - required)}"


def test_pptx_oxml_CT_classes_still_importable_explicitly():
    """Adding ``__all__`` must not break explicit ``from pptx.oxml import CT_X``.

    ``__all__`` only affects wildcard imports; the side-effect imports that
    register element classes must keep those classes reachable by name for
    downstream callers that import them directly.
    """
    # -- a handful of CT_* classes that have been reachable as top-level
    # -- `pptx.oxml` attributes since the module was written; import them
    # -- explicitly to confirm nothing broke.
    from pptx.oxml import CT_Hyperlink, CT_Presentation, CT_Slide

    assert CT_Hyperlink is not None
    assert CT_Presentation is not None
    assert CT_Slide is not None
