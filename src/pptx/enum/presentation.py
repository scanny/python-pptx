"""Enumerations that describe presentation-level settings.

Currently carries :class:`PP_VIEW_TYPE`, the value-space of
``p:viewPr/@lastView`` — the "remembered" editor view PowerPoint restores
when the presentation is re-opened.
"""

from __future__ import annotations

from pptx.enum.base import BaseXmlEnum


class PP_VIEW_TYPE(BaseXmlEnum):
    """Identifies a PowerPoint editor view.

    Maps ``ST_ViewType`` from ``pml.xsd``; each member's ``xml_value`` is the
    token PowerPoint writes into ``p:viewPr/@lastView`` (the view it will
    restore when the presentation is re-opened).

    MS API Name: *none — this enum is python-pptx-specific.*

    .. versionadded:: 2026.05.0
    """

    NORMAL = (1, "sldView", "Normal (slide-editing) view.")
    """Normal (slide-editing) view."""

    SLIDE_MASTER = (2, "sldMasterView", "Slide-master view.")
    """Slide-master view."""

    NOTES_PAGE = (3, "notesView", "Notes-page view.")
    """Notes-page view."""

    HANDOUT = (4, "handoutView", "Handout-master view.")
    """Handout-master view."""

    NOTES_MASTER = (5, "notesMasterView", "Notes-master view.")
    """Notes-master view."""

    OUTLINE = (6, "outlineView", "Outline view.")
    """Outline view."""

    SLIDE_SORTER = (7, "sldSorterView", "Slide-sorter view.")
    """Slide-sorter view."""

    SLIDE_THUMBNAIL = (
        8,
        "sldThumbnailView",
        "Slide-thumbnail view (Reading / Thumbnail pane).",
    )
    """Slide-thumbnail view (Reading / Thumbnail pane).

    PowerPoint's default template persists this value on a freshly-saved
    deck, so :meth:`.ViewProps.view_type` commonly returns
    :attr:`SLIDE_THUMBNAIL` on packages that have not been edited since
    their template was generated."""
