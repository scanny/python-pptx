"""DrawingML objects related to color, ColorFormat being the most prominent."""

from __future__ import annotations

import colorsys
from typing import Mapping

from pptx.dml._preset_colors import PRESET_COLORS
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.oxml.dml.color import (
    CT_HslColor,
    CT_PresetColor,
    CT_SchemeColor,
    CT_ScRgbColor,
    CT_SRgbColor,
    CT_SystemColor,
)


class ColorFormat(object):
    """
    Provides access to color settings such as RGB color, theme color, and
    luminance adjustments.
    """

    def __init__(self, eg_colorChoice_parent, color):
        super(ColorFormat, self).__init__()
        self._xFill = eg_colorChoice_parent
        self._color = color

    @property
    def alpha(self):
        """Read/write opacity of this color as a float in ``[0.0, 1.0]``.

        Corresponds to the ``<a:alpha val="N"/>`` child element where ``N``
        is an `ST_PositiveFixedPercentage` (``0`` through ``100000`` in
        OOXML, thousandths of a percent). A value of ``1.0`` is fully
        opaque and ``0.0`` is fully transparent. Returns ``1.0`` when no
        ``<a:alpha>`` element is present (opacity is the OOXML default).

        Assigning ``None`` removes any ``<a:alpha>`` child so the color
        inherits its opacity (effectively ``1.0``). Raises
        :class:`ValueError` when the assigned value is outside ``[0.0,
        1.0]`` or when this :class:`ColorFormat` has no color (setting
        alpha requires a color — set ``.rgb`` or ``.theme_color`` first).
        """
        return self._color.alpha

    @alpha.setter
    def alpha(self, value):
        if value is None:
            self._color.alpha = None
            return
        if value < 0.0 or value > 1.0:
            raise ValueError("alpha must be a number in range 0.0 to 1.0")
        if isinstance(self._color, _NoneColor):
            raise ValueError(
                "can't set alpha when color.type is None. Set color.rgb or"
                " .theme_color first."
            )
        self._color.alpha = value

    @property
    def brightness(self):
        """
        Read/write float value between -1.0 and 1.0 indicating the brightness
        adjustment for this color, e.g. -0.25 is 25% darker and 0.4 is 40%
        lighter. 0 means no brightness adjustment.
        """
        return self._color.brightness

    @brightness.setter
    def brightness(self, value):
        self._validate_brightness_value(value)
        self._color.brightness = value

    @classmethod
    def from_colorchoice_parent(cls, eg_colorChoice_parent):
        xClr = eg_colorChoice_parent.eg_colorChoice
        color = _Color(xClr)
        color_format = cls(eg_colorChoice_parent, color)
        return color_format

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, or None if no RGB color is explicitly
        defined for this font. Setting this value to an |RGBColor| instance
        causes its type to change to MSO_COLOR_TYPE.RGB. If the color was a
        theme color with a brightness adjustment, the brightness adjustment
        is removed when changing it to an RGB color.
        """
        return self._color.rgb

    @rgb.setter
    def rgb(self, rgb):
        if not isinstance(rgb, RGBColor):
            raise ValueError("assigned value must be type RGBColor")
        # change to rgb color format if not already
        if not isinstance(self._color, _SRgbColor):
            srgbClr = self._xFill.get_or_change_to_srgbClr()
            self._color = _SRgbColor(srgbClr)
        # call _SRgbColor instance to do the setting
        self._color.rgb = rgb

    @property
    def theme_color(self):
        """Theme color value of this color.

        Value is a member of :ref:`MsoThemeColorIndex`, e.g.
        ``MSO_THEME_COLOR.ACCENT_1``. Raises AttributeError on access if the
        color is not type ``MSO_COLOR_TYPE.SCHEME``. Assigning a member of
        :ref:`MsoThemeColorIndex` causes the color's type to change to
        ``MSO_COLOR_TYPE.SCHEME``.
        """
        return self._color.theme_color

    @theme_color.setter
    def theme_color(self, mso_theme_color_idx):
        # change to theme color format if not already
        if not isinstance(self._color, _SchemeColor):
            schemeClr = self._xFill.get_or_change_to_schemeClr()
            self._color = _SchemeColor(schemeClr)
        self._color.theme_color = mso_theme_color_idx

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor | None:
        """Return |RGBColor| resolved from whichever color type is present.

        Return |None| when no color is defined (color-type is None).

        `theme_colors` is an optional mapping from scheme-color name
        (e.g. ``"accent1"``, ``"bg1"``, ``"dk1"``) to |RGBColor| and is
        required only when this color is a scheme (theme) color. The mapping
        can be obtained from :attr:`SlideMaster.theme_colors`.

        The returned value reflects any luminance modifiers (``<a:lumMod>``
        and ``<a:lumOff>``, set via :attr:`brightness`) present on the color
        element, so the result corresponds to the RGB PowerPoint would
        actually render for the color — not just the unmodified base.

        Raises :class:`ValueError` for a scheme color when `theme_colors` is
        not provided or does not contain the required entry, and for a
        preset color whose name is not recognized.

        .. versionadded:: 2026.05.0
        """
        return self._color.to_rgb(theme_colors)

    @property
    def type(self):
        """
        Read-only. A value from :ref:`MsoColorType`, either RGB or SCHEME,
        corresponding to the way this color is defined, or None if no color
        is defined at the level of this font.
        """
        return self._color.color_type

    def _validate_brightness_value(self, value):
        if value < -1.0 or value > 1.0:
            raise ValueError("brightness must be number in range -1.0 to 1.0")
        if isinstance(self._color, _NoneColor):
            msg = (
                "can't set brightness when color.type is None. Set color.rgb"
                " or .theme_color first."
            )
            raise ValueError(msg)


class _Color(object):
    """
    Object factory for color object of the appropriate type, also the base
    class for all color type classes such as SRgbColor.
    """

    def __new__(cls, xClr):
        color_cls = {
            type(None): _NoneColor,
            CT_HslColor: _HslColor,
            CT_PresetColor: _PrstColor,
            CT_SchemeColor: _SchemeColor,
            CT_ScRgbColor: _ScRgbColor,
            CT_SRgbColor: _SRgbColor,
            CT_SystemColor: _SysColor,
        }[type(xClr)]
        return super(_Color, cls).__new__(color_cls)

    def __init__(self, xClr):
        super(_Color, self).__init__()
        self._xClr = xClr

    @property
    def alpha(self):
        """Return opacity as a float in ``[0.0, 1.0]``.

        Returns ``1.0`` when no ``<a:alpha>`` child is present (opacity is
        the default).
        """
        alpha_elm = self._xClr.alpha
        if alpha_elm is None:
            return 1.0
        return float(alpha_elm.val)

    @alpha.setter
    def alpha(self, value):
        if value is None:
            self._xClr.clear_alpha()
            return
        self._xClr.set_alpha(value)

    @property
    def brightness(self):
        lumMod, lumOff = self._xClr.lumMod, self._xClr.lumOff
        # a tint is lighter, a shade is darker
        # only tints have lumOff child
        if lumOff is not None:
            brightness = lumOff.val
            return brightness
        # which leaves shades, if lumMod is present
        if lumMod is not None:
            brightness = lumMod.val - 1.0
            return brightness
        # there's no brightness adjustment if no lum{Mod|Off} elements
        return 0

    @brightness.setter
    def brightness(self, value):
        if value > 0:
            self._tint(value)
        elif value < 0:
            self._shade(value)
        else:
            self._xClr.clear_lum()

    @property
    def color_type(self):  # pragma: no cover
        tmpl = ".color_type property must be implemented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)

    @property
    def rgb(self):
        """
        Raises TypeError on access unless overridden by subclass.
        """
        tmpl = "no .rgb property on color type '%s'"
        raise AttributeError(tmpl % self.__class__.__name__)

    @property
    def theme_color(self):
        """
        Raises TypeError on access unless overridden by subclass.
        """
        return MSO_THEME_COLOR.NOT_THEME_COLOR

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor | None:
        """Return |RGBColor| corresponding to this color.

        Must be overridden by subclasses that can be resolved to RGB.
        """
        tmpl = "no .to_rgb() implementation on color type '%s'"
        raise NotImplementedError(tmpl % self.__class__.__name__)

    def _apply_lum_mods(self, rgb: RGBColor) -> RGBColor:
        """Return `rgb` adjusted by any `<a:lumMod>`/`<a:lumOff>` children.

        Applies the ECMA-376 luminance transform::

            new_L = old_L * lumMod + lumOff

        in the HSL color space, with `lumMod` defaulting to `1.0` and
        `lumOff` defaulting to `0.0`. When neither modifier is present the
        input `rgb` is returned unchanged.
        """
        lumMod_elm = self._xClr.lumMod
        lumOff_elm = self._xClr.lumOff
        if lumMod_elm is None and lumOff_elm is None:
            return rgb

        lum_mod = 1.0 if lumMod_elm is None else float(lumMod_elm.val)
        lum_off = 0.0 if lumOff_elm is None else float(lumOff_elm.val)

        r, g, b = rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0
        hue, lum, sat = colorsys.rgb_to_hls(r, g, b)
        new_lum = max(0.0, min(1.0, lum * lum_mod + lum_off))
        nr, ng, nb = colorsys.hls_to_rgb(hue, new_lum, sat)
        return RGBColor(
            max(0, min(255, int(round(nr * 255)))),
            max(0, min(255, int(round(ng * 255)))),
            max(0, min(255, int(round(nb * 255)))),
        )

    def _shade(self, value):
        lumMod_val = 1.0 - abs(value)
        color_elm = self._xClr.clear_lum()
        color_elm.add_lumMod(lumMod_val)

    def _tint(self, value):
        lumOff_val = value
        lumMod_val = 1.0 - lumOff_val
        color_elm = self._xClr.clear_lum()
        color_elm.add_lumMod(lumMod_val)
        color_elm.add_lumOff(lumOff_val)


class _HslColor(_Color):
    @property
    def color_type(self):
        return MSO_COLOR_TYPE.HSL

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor:
        """Return |RGBColor| computed from the HSL triple on `<a:hslClr>`.

        OOXML encodes `hue` as 60000-ths of a degree and `sat`/`lum` as
        1000-ths of a percent. Any `<a:lumMod>` or `<a:lumOff>` child is
        applied to the result.
        """
        hue = _ST_Angle_to_degrees(self._xClr.get("hue", "0"))
        sat = _ST_Percentage_to_unit(self._xClr.get("sat", "0"))
        lum = _ST_Percentage_to_unit(self._xClr.get("lum", "0"))
        # -- colorsys.hls_to_rgb expects (h, l, s) where h is in [0, 1) --
        r, g, b = colorsys.hls_to_rgb(hue / 360.0, lum, sat)
        base = RGBColor(int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))
        return self._apply_lum_mods(base)


class _NoneColor(_Color):
    @property
    def alpha(self):
        """Return 1.0, the default opacity for a color with no explicit value.

        ``_NoneColor`` carries no XML element, so there can be no
        ``<a:alpha>`` child; OOXML treats an absent alpha as fully opaque.
        """
        return 1.0

    @alpha.setter
    def alpha(self, value):  # pragma: no cover
        # -- unreachable: ColorFormat.alpha guards against setting on a
        # -- _NoneColor by raising ValueError before delegating here.
        raise ValueError(
            "can't set alpha when color.type is None. Set color.rgb or"
            " .theme_color first."
        )

    @property
    def color_type(self):
        return None

    @property
    def theme_color(self):
        """
        Raise TypeError on attempt to access .theme_color when no color
        choice is present.
        """
        tmpl = "no .theme_color property on color type '%s'"
        raise AttributeError(tmpl % self.__class__.__name__)

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> None:
        """Return |None|, indicating no color is defined at this level."""
        return None


class _PrstColor(_Color):
    @property
    def color_type(self):
        return MSO_COLOR_TYPE.PRESET

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor:
        """Return |RGBColor| for the preset color name on `<a:prstClr>`.

        Any `<a:lumMod>` or `<a:lumOff>` child is applied to the result.

        Raises :class:`ValueError` when the preset name is not recognized.
        """
        name = self._xClr.get("val", "")
        try:
            hex_str = PRESET_COLORS[name]
        except KeyError as exc:
            raise ValueError("unrecognized preset color name '%s'" % name) from exc
        return self._apply_lum_mods(RGBColor.from_string(hex_str))


class _SchemeColor(_Color):
    def __init__(self, schemeClr):
        super(_SchemeColor, self).__init__(schemeClr)
        self._schemeClr = schemeClr

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.SCHEME

    @property
    def theme_color(self):
        """
        Theme color value of this color, one of those defined in the
        MSO_THEME_COLOR enumeration, e.g. MSO_THEME_COLOR.ACCENT_1. None if
        no theme color is explicitly defined for this font. Setting this to a
        value in MSO_THEME_COLOR causes the color's type to change to
        ``MSO_COLOR_TYPE.SCHEME``.
        """
        return self._schemeClr.val

    @theme_color.setter
    def theme_color(self, mso_theme_color_idx):
        self._schemeClr.val = mso_theme_color_idx

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor:
        """Return |RGBColor| for this scheme color.

        `theme_colors` must be provided and must be a mapping from scheme
        color name (e.g. ``"accent1"``, ``"bg1"``) to |RGBColor|. Such a
        mapping can be obtained from :attr:`SlideMaster.theme_colors`.

        Any `<a:lumMod>` or `<a:lumOff>` child is applied to the result, so
        a tint- or shade-adjusted theme color resolves to the RGB value
        PowerPoint actually renders.

        Raises :class:`ValueError` when `theme_colors` is None or does not
        contain an entry for this color's scheme name.
        """
        if theme_colors is None:
            raise ValueError(
                "scheme color resolution requires a `theme_colors` mapping; "
                "see SlideMaster.theme_colors"
            )
        name = self._schemeClr.get("val", "")
        try:
            base = theme_colors[name]
        except KeyError as exc:
            raise ValueError("theme_colors has no entry for scheme color '%s'" % name) from exc
        return self._apply_lum_mods(base)


class _ScRgbColor(_Color):
    @property
    def color_type(self):
        return MSO_COLOR_TYPE.SCRGB

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor:
        """Return |RGBColor| computed from the percentage-based RGB attrs.

        Any `<a:lumMod>` or `<a:lumOff>` child is applied to the result.
        """
        r = _ST_Percentage_to_unit(self._xClr.get("r", "0"))
        g = _ST_Percentage_to_unit(self._xClr.get("g", "0"))
        b = _ST_Percentage_to_unit(self._xClr.get("b", "0"))
        base = RGBColor(
            max(0, min(255, int(round(r * 255)))),
            max(0, min(255, int(round(g * 255)))),
            max(0, min(255, int(round(b * 255)))),
        )
        return self._apply_lum_mods(base)


class _SRgbColor(_Color):
    def __init__(self, srgbClr):
        super(_SRgbColor, self).__init__(srgbClr)
        self._srgbClr = srgbClr

    @property
    def color_type(self):
        return MSO_COLOR_TYPE.RGB

    @property
    def rgb(self):
        """
        |RGBColor| value of this color, corresponding to the value in the
        required ``val`` attribute of the ``<a:srgbColr>`` element.
        """
        return RGBColor.from_string(self._srgbClr.val)

    @rgb.setter
    def rgb(self, rgb):
        self._srgbClr.val = str(rgb)

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor:
        """Return |RGBColor| for this explicit RGB color.

        Any `<a:lumMod>` or `<a:lumOff>` child is applied to the result.
        """
        return self._apply_lum_mods(RGBColor.from_string(self._srgbClr.val))


class _SysColor(_Color):
    @property
    def color_type(self):
        return MSO_COLOR_TYPE.SYSTEM

    def to_rgb(self, theme_colors: Mapping[str, RGBColor] | None = None) -> RGBColor:
        """Return |RGBColor| stored on the `lastClr` attribute.

        `lastClr` is the rendered RGB value PowerPoint most recently observed
        for the named system color (like ``"windowText"``) and is the closest
        thing to a resolved RGB available without the rendering context. Any
        `<a:lumMod>` or `<a:lumOff>` child is applied to the result.

        Raises :class:`ValueError` when the element has no `lastClr`
        attribute.
        """
        last_clr = self._xClr.get("lastClr")
        if last_clr is None:
            raise ValueError(
                "system color '%s' has no `lastClr` attribute; RGB cannot be "
                "resolved without a rendering context" % self._xClr.get("val", "")
            )
        return self._apply_lum_mods(RGBColor.from_string(last_clr))


class RGBColor(tuple):
    """
    Immutable value object defining a particular RGB color.
    """

    def __new__(cls, r, g, b):
        msg = "RGBColor() takes three integer values 0-255"
        for val in (r, g, b):
            if not isinstance(val, int) or val < 0 or val > 255:
                raise ValueError(msg)
        return super(RGBColor, cls).__new__(cls, (r, g, b))

    def __str__(self):
        """
        Return a hex string rgb value, like '3C2F80'
        """
        return "%02X%02X%02X" % self

    @classmethod
    def from_string(cls, rgb_hex_str):
        """
        Return a new instance from an RGB color hex string like ``'3C2F80'``.
        """
        r = int(rgb_hex_str[:2], 16)
        g = int(rgb_hex_str[2:4], 16)
        b = int(rgb_hex_str[4:], 16)
        return cls(r, g, b)


def _ST_Percentage_to_unit(value: str | float) -> float:
    """Parse an `ST_Percentage`-style OOXML value into a float in [0.0, 1.0].

    Accepts either the OOXML integer form (``"55000"`` = 55%) or the
    schema-form with trailing ``%`` (``"55%"``).
    """
    if isinstance(value, (int, float)):
        # -- already a numeric fraction, assume 1000-ths of a percent --
        return float(value) / 100000.0
    if value.endswith("%"):
        return float(value[:-1]) / 100.0
    return float(value) / 100000.0


def _ST_Angle_to_degrees(value: str | float) -> float:
    """Parse an `ST_Angle`-style OOXML value into a float in [0.0, 360.0).

    Accepts either the OOXML integer form (``"21600000"`` = 360°, i.e.
    60000-ths of a degree) or the plain decimal-degree form.
    """
    if isinstance(value, (int, float)):
        return (float(value) / 60000.0) % 360.0
    return (float(value) / 60000.0) % 360.0
