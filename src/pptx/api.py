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


def Presentation(pptx: str | IO[bytes] | None = None) -> presentation.Presentation:
    """
    Return a |Presentation| object loaded from *pptx*, where *pptx* can be
    either a path to a ``.pptx`` file (a string) or a file-like object. If
    *pptx* is missing or ``None``, the built-in default presentation
    "template" is loaded.
    """
    if pptx is None:
        pptx = _default_pptx_path()

    presentation_part = Package.open(pptx).main_document_part

    if not _is_pptx_package(presentation_part):
        tmpl = "file '%s' is not a PowerPoint file, content type is '%s'"
        raise ValueError(tmpl % (pptx, presentation_part.content_type))

    return presentation_part.presentation


def _default_pptx_path() -> str:
    """Return the path to the built-in default .pptx package."""
    _thisdir = os.path.split(__file__)[0]
    return os.path.join(_thisdir, "templates", "default.pptx")


def _is_pptx_package(prs_part: PresentationPart):
    """Return |True| if *prs_part* is a valid main document part, |False| otherwise."""
    valid_content_types = (CT.PML_PRESENTATION_MAIN, CT.PML_PRES_MACRO_MAIN)
    return prs_part.content_type in valid_content_types
