.. _PpViewType:

``PP_VIEW_TYPE``
================

Identifies an editor view in PowerPoint. Value-space of
``p:viewPr/@lastView`` — the view PowerPoint restores when the
presentation is re-opened. See issue #94.

Example::

    from pptx.enum.presentation import PP_VIEW_TYPE

    prs.view_props.view_type = PP_VIEW_TYPE.OUTLINE

----

NORMAL
    Normal (slide-editing) view.

SLIDE_MASTER
    Slide-master view.

NOTES_PAGE
    Notes-page view.

HANDOUT
    Handout-master view.

NOTES_MASTER
    Notes-master view.

OUTLINE
    Outline view.

SLIDE_SORTER
    Slide-sorter view.

SLIDE_THUMBNAIL
    Slide-thumbnail view (Reading / Thumbnail pane). PowerPoint's
    default template persists this value on a freshly-saved deck.
