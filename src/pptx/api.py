"""Directly exposed API classes, Presentation for now.

Provides some syntactic sugar for interacting with the pptx.presentation.Package graph and also
provides some insulation so not so many classes in the other modules need to be named as internal
(leading underscore).
"""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.package import Package

if TYPE_CHECKING:
    from pptx import presentation
    from pptx.parts.presentation import PresentationPart


#: Slide-size preset specifications keyed by preset name.
#:
#: Each value is a 4-tuple ``(template_filename, cx, cy, sld_sz_type)``:
#:
#: * ``template_filename`` - built-in template (under ``src/pptx/templates/``)
#:   used as the starting point. It supplies the slide master / layouts and
#:   has its ``p:sldSz`` rewritten to the preset's dimensions on load.
#: * ``cx`` / ``cy`` - slide width / height in English Metric Units.
#: * ``sld_sz_type`` - value written to ``p:sldSz/@type``. ``None`` omits the
#:   attribute (which ECMA-376 treats as ``custom``).
#:
#: EMU conversions: 914400 EMU/inch and 36000 EMU/mm.
#:
#: The widescreen / 16:9 preset loads the dedicated ``default-16x9.pptx``
#: template (whose placeholders are already sized for widescreen); every
#: other preset starts from the 4:3 ``default.pptx`` and overrides ``p:sldSz``
#: in-memory before parsing, so a single template file backs all paper sizes.
_PRESET_SPECS: dict[str, tuple[str, int, int, str | None]] = {
    # 10" x 7.5" (ECMA `screen4x3`) - historic 4:3 default
    "4x3": ("default.pptx", 9144000, 6858000, "screen4x3"),
    "standard": ("default.pptx", 9144000, 6858000, "screen4x3"),
    # 13.333" x 7.5" (ECMA `screen16x9`)
    "16x9": ("default-16x9.pptx", 12192000, 6858000, "screen16x9"),
    "widescreen": ("default-16x9.pptx", 12192000, 6858000, "screen16x9"),
    # US Letter paper - ECMA `letter`; shares the historic 4:3 dimensions
    # (10" x 7.5") but tags the slide as letter paper so PowerPoint's Page
    # Setup dialog displays "Letter Paper (8.5x11 in)" rather than
    # "On-screen Show (4:3)".
    "letter": ("default.pptx", 9144000, 6858000, "letter"),
    # A4 landscape (297mm x 210mm) - ECMA `A4`
    "a4": ("default.pptx", 10692000, 7560000, "A4"),
}


def Presentation(
    pptx: str | IO[bytes] | None = None,
    pptx_format: str | tuple[int, int] | None = None,
    password: str | None = None,
) -> presentation.Presentation:
    """Return a |Presentation| object loaded from *pptx*.

    *pptx* can be either a path to a ``.pptx`` file (a string) or a file-like
    object. If *pptx* is missing or ``None``, the built-in default presentation
    "template" is loaded.

    When *pptx* is ``None``, *pptx_format* selects the slide size of the
    built-in template. Accepted values are:

    * a preset name (case-insensitive): ``"4x3"`` / ``"standard"``,
      ``"16x9"`` / ``"widescreen"``, ``"letter"``, or ``"a4"``;
    * a ``(cx, cy)`` tuple giving a custom slide width and height in English
      Metric Units (EMU) - the ``p:sldSz/@type`` attribute is set to
      ``"custom"``;
    * ``None`` (the default), which resolves to ``"4x3"`` for backwards
      compatibility.

    Providing *pptx_format* together with a non-``None`` *pptx* raises
    :class:`ValueError`, because the slide size is determined by the opened
    file rather than the preset.

    When *password* is provided, an encrypted (password-protected) ``.pptx``
    file is decrypted before loading. This requires the optional
    ``msoffcrypto-tool`` dependency. Opening an encrypted package without
    a password raises :class:`pptx.exc.EncryptedPackageError`.
    """
    if pptx_format is not None and pptx is not None:
        raise ValueError(
            "pptx_format is only valid when opening the default template (pptx is None)"
        )

    opened_default = pptx is None
    if opened_default:
        template_path, target_cx, target_cy, sld_sz_type = _resolve_preset(pptx_format)
        pptx = template_path
    else:
        target_cx = target_cy = 0
        sld_sz_type = None

    presentation_part = Package.open(pptx, password=password).main_document_part

    if not _is_pptx_package(presentation_part):
        tmpl = "file '%s' is not a PowerPoint file, content type is '%s'"
        raise ValueError(tmpl % (pptx, presentation_part.content_type))

    if opened_default:
        _apply_slide_size(presentation_part, target_cx, target_cy, sld_sz_type)

    return presentation_part.presentation


def _apply_slide_size(
    prs_part: PresentationPart, cx: int, cy: int, sld_sz_type: str | None
) -> None:
    """Rewrite `p:sldSz` on *prs_part*'s element to (*cx*, *cy*, *sld_sz_type*)."""
    prs_elm = prs_part._element  # pyright: ignore[reportPrivateUsage]
    sldSz = prs_elm.get_or_add_sldSz()
    sldSz.cx = cx  # type: ignore[assignment]
    sldSz.cy = cy  # type: ignore[assignment]
    sldSz.type = sld_sz_type


def _resolve_preset(
    pptx_format: str | tuple[int, int] | None,
) -> tuple[str, int, int, str | None]:
    """Return ``(template_path, cx, cy, sld_sz_type)`` for *pptx_format*."""
    _thisdir = os.path.split(__file__)[0]

    if pptx_format is None:
        spec = _PRESET_SPECS["4x3"]
        template_path = os.path.join(_thisdir, "templates", spec[0])
        return template_path, spec[1], spec[2], spec[3]

    if isinstance(pptx_format, tuple):
        if len(pptx_format) != 2:
            raise ValueError(
                "pptx_format tuple must be (cx, cy) in EMU; got %r" % (pptx_format,)
            )
        cx, cy = pptx_format
        if not (isinstance(cx, int) and isinstance(cy, int)):
            raise ValueError(
                "pptx_format tuple (cx, cy) must contain integers in EMU; got %r"
                % (pptx_format,)
            )
        # -- pick the base template whose aspect ratio is closest to the
        # -- requested size so placeholders render cleanly
        base_filename = "default-16x9.pptx" if cx * 3 > cy * 4 else "default.pptx"
        template_path = os.path.join(_thisdir, "templates", base_filename)
        return template_path, cx, cy, "custom"

    if not isinstance(pptx_format, str):
        raise TypeError(
            "pptx_format must be a preset name or (cx, cy) tuple; got %r" % (pptx_format,)
        )

    key = pptx_format.lower()
    if key not in _PRESET_SPECS:
        valid = ", ".join(sorted(_PRESET_SPECS))
        raise ValueError("unknown pptx_format %r; expected one of: %s" % (pptx_format, valid))
    spec = _PRESET_SPECS[key]
    template_path = os.path.join(_thisdir, "templates", spec[0])
    return template_path, spec[1], spec[2], spec[3]


def _default_pptx_path(pptx_format: str | None = None) -> str:
    """Return the path to the built-in default .pptx package.

    *pptx_format* selects the aspect-ratio preset. ``None`` resolves to the
    original 4:3 template. Valid string values are the keys of
    :data:`_PRESET_SPECS` (case-insensitive). An unknown preset name raises
    :class:`ValueError`.

    This helper is retained for backwards compatibility with callers that
    only need the file path; it does not apply the preset's slide-size
    override to the resulting package.
    """
    template_path, _, _, _ = _resolve_preset(pptx_format)
    return template_path


#: Backwards-compatible alias for :data:`_PRESET_SPECS`. Maps each preset
#: name to just the template filename, matching the shape used before the
#: letter / A4 / tuple-size support landed.
_PRESET_TEMPLATES: dict[str, str] = {key: spec[0] for key, spec in _PRESET_SPECS.items()}


def _is_pptx_package(prs_part: PresentationPart):
    """Return |True| if *prs_part* is a valid main document part, |False| otherwise.

    Accepts the four PresentationML main-part content types: a regular presentation
    (``.pptx``), a macro-enabled presentation (``.pptm``), a template (``.potx``), and a
    slideshow (``.ppsx``).
    """
    valid_content_types = (
        CT.PML_PRESENTATION_MAIN,
        CT.PML_PRES_MACRO_MAIN,
        CT.PML_TEMPLATE_MAIN,
        CT.PML_SLIDESHOW_MAIN,
        CT.PML_SLIDESHOW_MACRO_MAIN,
    )
    return prs_part.content_type in valid_content_types
