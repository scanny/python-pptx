"""Model3D proxy — read-only access to an embedded 3D model on a graphic frame.

PowerPoint 365 (Office 2016+) "Insert > 3D Models" produces a graphic-frame whose
`a:graphicData/@uri` is
:data:`pptx.spec.GRAPHIC_DATA_URI_MODEL_3D` and whose single child is an
`am3d:model3D` element. That element carries an `r:embed` attribute pointing at a
slide-relationship for the embedded 3D-model part (typically a binary glTF `.glb`,
sometimes `.obj` / `.fbx`), plus an `ext` attribute recording the file extension and
a tree of scene/camera/lighting descendants.

Authoring a 3D model from scratch requires understanding camera parameters, lighting
rigs, and the embedded model format; that is deliberately **out of scope** for the
issue #410 MVP. This module ships a thin read-only proxy:

* :class:`Model3D` — surfaced via :attr:`~pptx.shapes.graphfrm.GraphicFrame.model_3d`
  when :attr:`~pptx.shapes.graphfrm.GraphicFrame.has_model_3d` is |True|. Exposes the
  relationship id, the raw bytes of the embedded model file, and the `ext` discriminator.

See ``docs/dev/analysis/model-3d.rst`` for the full design rationale and the roadmap
to full authoring support.

.. versionadded:: 2026.05.0
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.shared import ParentedElementProxy

if TYPE_CHECKING:
    from pptx.oxml.shapes.graphfrm import CT_Model3D
    from pptx.parts.slide import BaseSlidePart
    from pptx.types import ProvidesPart


class Model3D(ParentedElementProxy):
    """Provides read-only access to an embedded 3D model.

    A 3D model on a slide is a ``p:graphicFrame`` whose ``a:graphicData/@uri`` matches
    :data:`~pptx.spec.GRAPHIC_DATA_URI_MODEL_3D`. The single ``am3d:model3D`` child
    carries the relationship to the embedded model file (``.glb`` / ``.obj`` /
    ``.fbx``) and the scene/camera parameters PowerPoint uses to render it.

    MVP scope (issue #410): :attr:`embedded_rel_id`, :attr:`media_blob`, and :attr:`ext`
    expose the embedded asset. Scene/camera introspection and authoring are deferred to
    a follow-up feature iteration; see ``docs/dev/analysis/model-3d.rst``.

    .. versionadded:: 2026.05.0
    """

    part: BaseSlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, model3D: CT_Model3D, parent: ProvidesPart):
        super().__init__(model3D, parent)
        self._model3D = model3D

    @property
    def embedded_rel_id(self) -> str | None:
        """str value of the `r:embed` attribute, or |None| when absent.

        This is the relationship id (against the slide part) that resolves to the
        embedded 3D-model part.

        .. versionadded:: 2026.05.0
        """
        return self._model3D.embed_rId

    @property
    def ext(self) -> str | None:
        """str file-extension of the embedded model, or |None| when absent.

        Typical values are ``"glb"`` (binary glTF, the format PowerPoint writes by
        default), ``"obj"``, and ``"fbx"``. The extension is recorded on the
        ``am3d:model3D`` element itself as an `ext` attribute and mirrors the file
        extension of the embedded part.

        .. versionadded:: 2026.05.0
        """
        return self._model3D.ext

    @property
    def media_blob(self) -> bytes | None:
        """Bytes of the embedded 3D-model file, or |None| when unresolvable.

        Returns |None| when the `r:embed` attribute is missing or when the
        relationship cannot be resolved to a part. Otherwise returns the full
        contents of the embedded part — suitable for writing to a file with the
        :attr:`ext` extension.

        .. versionadded:: 2026.05.0
        """
        rId = self._model3D.embed_rId
        if rId is None:
            return None
        try:
            return self.part.related_part(rId).blob
        except KeyError:
            return None
