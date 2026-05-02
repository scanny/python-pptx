"""View-properties part, corresponds to ``/ppt/viewProps.xml`` in the package.

This part records editor-view settings that PowerPoint restores when the
presentation is re-opened: which editor view was last active, whether the
comments pane is visible, and the zoom / scroll-position state of each
individual view (normal, slide, notes, outline, sorter). See issue #94.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.viewprops import CT_ViewProperties

if TYPE_CHECKING:
    from pptx.package import Package


class ViewPropsPart(XmlPart):
    """Corresponds to part named ``/ppt/viewProps.xml``.

    Holds editor-view settings for the presentation. Created lazily when the
    package does not already contain one; see :attr:`.Presentation.view_props`.
    """

    _element: CT_ViewProperties

    @classmethod
    def default(cls, package: Package) -> ViewPropsPart:
        """Return a default, empty view-properties part.

        Suitable as the starting point for a package that does not yet have
        a ``ppt/viewProps.xml``. Carries only the bare ``<p:viewPr/>`` root,
        which PowerPoint and the schema both tolerate.
        """
        return cls(
            PackURI("/ppt/viewProps.xml"),
            CT.PML_VIEW_PROPS,
            package,
            CT_ViewProperties.new_default(),
        )
