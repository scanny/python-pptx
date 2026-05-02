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


#: Aspect-ratio preset names accepted by :func:`Presentation`.
#:
#: Each key maps to the filename of a built-in template under
#: ``src/pptx/templates/``. Aliases such as ``"widescreen"`` resolve to the
#: same template as their canonical name (``"16x9"``).
_PRESET_TEMPLATES: dict[str, str] = {
    "4x3": "default.pptx",
    "standard": "default.pptx",
    "16x9": "default-16x9.pptx",
    "widescreen": "default-16x9.pptx",
}


def Presentation(
    pptx: str | IO[bytes] | None = None,
    pptx_format: str | None = None,
    password: str | None = None,
) -> presentation.Presentation:
    """Return a |Presentation| object loaded from *pptx*.

    *pptx* can be either a path to a ``.pptx`` file (a string) or a file-like
    object. If *pptx* is missing or ``None``, the built-in default presentation
    "template" is loaded.

    When *pptx* is ``None``, *pptx_format* selects the aspect ratio of the
    built-in template. Accepted values are ``"4x3"`` (also spelled
    ``"standard"``), ``"16x9"`` (also spelled ``"widescreen"``), and ``None``
    (the default, which resolves to ``"4x3"`` for backwards compatibility).
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

    if pptx is None:
        pptx = _default_pptx_path(pptx_format)

    presentation_part = Package.open(pptx, password=password).main_document_part

    if not _is_pptx_package(presentation_part):
        tmpl = "file '%s' is not a PowerPoint file, content type is '%s'"
        raise ValueError(tmpl % (pptx, presentation_part.content_type))

    return presentation_part.presentation


def _default_pptx_path(pptx_format: str | None = None) -> str:
    """Return the path to the built-in default .pptx package.

    *pptx_format* selects the aspect-ratio preset. ``None`` resolves to the
    original 4:3 template. Valid string values are the keys of
    :data:`_PRESET_TEMPLATES` (case-insensitive). An unknown preset name
    raises :class:`ValueError`.
    """
    key = "4x3" if pptx_format is None else pptx_format.lower()
    if key not in _PRESET_TEMPLATES:
        valid = ", ".join(sorted(_PRESET_TEMPLATES))
        raise ValueError("unknown pptx_format %r; expected one of: %s" % (pptx_format, valid))

    _thisdir = os.path.split(__file__)[0]
    return os.path.join(_thisdir, "templates", _PRESET_TEMPLATES[key])


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
    )
    return prs_part.content_type in valid_content_types
