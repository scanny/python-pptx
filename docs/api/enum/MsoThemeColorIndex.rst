.. _MsoThemeColorIndex:

``MSO_THEME_COLOR_INDEX``
=========================

Indicates the Office theme color, one of those shown in the color gallery on
the formatting ribbon.

Alias: ``MSO_THEME_COLOR``

Example::

    from pptx.enum.dml import MSO_THEME_COLOR

    shape.fill.solid()
    shape.fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1

XML mapping
-----------

Each enum member writes a specific ``val`` on ``<a:schemeClr>``. OOXML
defines *two* parallel sets of scheme-color names, connected at the theme
level by ``<a:clrMap>``:

**Slide-level names** — what PowerPoint writes in slide / layout /
master content (shape fills, run text colors, etc.):

==================================== ===============
``MSO_THEME_COLOR_INDEX`` member     ``val`` written
==================================== ===============
``TEXT_1``                           ``tx1``
``BACKGROUND_1``                     ``bg1``
``TEXT_2``                           ``tx2``
``BACKGROUND_2``                     ``bg2``
``ACCENT_1`` … ``ACCENT_6``          ``accent1`` … ``accent6``
``HYPERLINK``                        ``hlink``
``FOLLOWED_HYPERLINK``               ``folHlink``
==================================== ===============

**Theme-level names** — what appears inside the theme part's own
``<a:clrScheme>``:

==================================== ===============
``MSO_THEME_COLOR_INDEX`` member     ``val`` written
==================================== ===============
``DARK_1``                           ``dk1``
``LIGHT_1``                          ``lt1``
``DARK_2``                           ``dk2``
``LIGHT_2``                          ``lt2``
``ACCENT_1`` … ``ACCENT_6``          ``accent1`` … ``accent6``
``HYPERLINK``                        ``hlink``
``FOLLOWED_HYPERLINK``               ``folHlink``
==================================== ===============

The slide master's ``<p:clrMap>`` connects the two sets, typically as
``bg1 → lt1``, ``tx1 → dk1``, ``bg2 → lt2``, ``tx2 → dk2`` (accents and
hyperlinks pass through by name). Authors can remap to swap light and
dark schemes without editing any slide content.

**Which member should I assign?** Match what PowerPoint writes at the
same location:

* On a run, shape fill, or other slide-level content, use
  ``TEXT_1`` / ``BACKGROUND_1`` / ``TEXT_2`` / ``BACKGROUND_2`` (these
  write ``tx1`` / ``bg1`` / ``tx2`` / ``bg2``). This is what the
  PowerPoint UI emits.

* Inside a theme's own ``<a:clrScheme>``, use ``DARK_1`` / ``LIGHT_1``
  / ``DARK_2`` / ``LIGHT_2`` (these write ``dk1`` / ``lt1`` / ``dk2`` /
  ``lt2``).

Both sets are accepted by PowerPoint in slide content — the slide-level
names are strongly preferred for interoperability and round-trip
stability.

----

NOT_THEME_COLOR
    Indicates the color is not a theme color.

ACCENT_1
    Specifies the Accent 1 theme color. Writes ``<a:schemeClr val="accent1">``.

ACCENT_2
    Specifies the Accent 2 theme color. Writes ``<a:schemeClr val="accent2">``.

ACCENT_3
    Specifies the Accent 3 theme color. Writes ``<a:schemeClr val="accent3">``.

ACCENT_4
    Specifies the Accent 4 theme color. Writes ``<a:schemeClr val="accent4">``.

ACCENT_5
    Specifies the Accent 5 theme color. Writes ``<a:schemeClr val="accent5">``.

ACCENT_6
    Specifies the Accent 6 theme color. Writes ``<a:schemeClr val="accent6">``.

BACKGROUND_1
    Specifies the Background 1 theme color (slide-level). Writes
    ``<a:schemeClr val="bg1">``. Mapped to ``lt1`` by a default
    ``<p:clrMap>``. See also ``LIGHT_1``.

BACKGROUND_2
    Specifies the Background 2 theme color (slide-level). Writes
    ``<a:schemeClr val="bg2">``. Mapped to ``lt2`` by a default
    ``<p:clrMap>``. See also ``LIGHT_2``.

DARK_1
    Specifies the Dark 1 theme color (theme-level). Writes
    ``<a:schemeClr val="dk1">``. Used inside a theme's ``<a:clrScheme>``;
    slide content typically uses ``TEXT_1`` (which maps to this via
    ``<p:clrMap>``).

DARK_2
    Specifies the Dark 2 theme color (theme-level). Writes
    ``<a:schemeClr val="dk2">``. Used inside a theme's ``<a:clrScheme>``;
    slide content typically uses ``TEXT_2`` (which maps to this via
    ``<p:clrMap>``).

FOLLOWED_HYPERLINK
    Specifies the theme color for a clicked hyperlink. Writes
    ``<a:schemeClr val="folHlink">``.

HYPERLINK
    Specifies the theme color for a hyperlink. Writes
    ``<a:schemeClr val="hlink">``.

LIGHT_1
    Specifies the Light 1 theme color (theme-level). Writes
    ``<a:schemeClr val="lt1">``. Used inside a theme's ``<a:clrScheme>``;
    slide content typically uses ``BACKGROUND_1`` (which maps to this via
    ``<p:clrMap>``).

LIGHT_2
    Specifies the Light 2 theme color (theme-level). Writes
    ``<a:schemeClr val="lt2">``. Used inside a theme's ``<a:clrScheme>``;
    slide content typically uses ``BACKGROUND_2`` (which maps to this via
    ``<p:clrMap>``).

TEXT_1
    Specifies the Text 1 theme color (slide-level). Writes
    ``<a:schemeClr val="tx1">``. Mapped to ``dk1`` by a default
    ``<p:clrMap>``. See also ``DARK_1``.

TEXT_2
    Specifies the Text 2 theme color (slide-level). Writes
    ``<a:schemeClr val="tx2">``. Mapped to ``dk2`` by a default
    ``<p:clrMap>``. See also ``DARK_2``.

MIXED
    Indicates multiple theme colors are used, such as in a group shape.
