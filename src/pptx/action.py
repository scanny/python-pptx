"""Objects related to mouse click and hover actions on a shape or text."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, cast

from pptx.enum.action import PP_ACTION
from pptx.media import Audio
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes import Subshape
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.action import CT_EmbeddedWAVAudioFile, CT_Hyperlink
    from pptx.oxml.shapes.shared import CT_NonVisualDrawingProps
    from pptx.oxml.text import CT_TextCharacterProperties
    from pptx.parts.slide import SlidePart
    from pptx.shapes.base import BaseShape
    from pptx.slide import Slide, Slides


class ActionSetting(Subshape):
    """Properties specifying how a shape or run reacts to mouse actions."""

    # -- The Subshape base class provides access to the Slide Part, which is needed to access
    # -- relationships, which is where hyperlinks live.

    def __init__(
        self,
        xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties,
        parent: BaseShape,
        hover: bool = False,
    ):
        super(ActionSetting, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover

    @property
    def action(self):
        """Member of :ref:`PpActionType` enumeration, such as `PP_ACTION.HYPERLINK`.

        The returned member indicates the type of action that will result when the
        specified shape or text is clicked or the mouse pointer is positioned over the
        shape during a slide show.

        If there is no click-action or the click-action value is not recognized (is not
        one of the official `MsoPpAction` values) then `PP_ACTION.NONE` is returned.
        """
        hlink = self._hlink

        if hlink is None:
            return PP_ACTION.NONE

        action_verb = hlink.action_verb

        if action_verb == "hlinkshowjump":
            relative_target = hlink.action_fields["jump"]
            return {
                "firstslide": PP_ACTION.FIRST_SLIDE,
                "lastslide": PP_ACTION.LAST_SLIDE,
                "lastslideviewed": PP_ACTION.LAST_SLIDE_VIEWED,
                "nextslide": PP_ACTION.NEXT_SLIDE,
                "previousslide": PP_ACTION.PREVIOUS_SLIDE,
                "endshow": PP_ACTION.END_SHOW,
            }[relative_target]

        return {
            None: PP_ACTION.HYPERLINK,
            "hlinksldjump": PP_ACTION.NAMED_SLIDE,
            "hlinkpres": PP_ACTION.PLAY,
            "hlinkfile": PP_ACTION.OPEN_FILE,
            "customshow": PP_ACTION.NAMED_SLIDE_SHOW,
            "ole": PP_ACTION.OLE_VERB,
            "macro": PP_ACTION.RUN_MACRO,
            "program": PP_ACTION.RUN_PROGRAM,
        }.get(action_verb, PP_ACTION.NONE)

    @lazyproperty
    def hyperlink(self) -> Hyperlink:
        """
        A |Hyperlink| object representing the hyperlink action defined on
        this click or hover mouse event. A |Hyperlink| object is always
        returned, even if no hyperlink or other click action is defined.
        """
        return Hyperlink(self._element, self._parent, self._hover)

    @property
    def screen_tip(self) -> str | None:
        """The ScreenTip (tooltip) shown on hover for this click/hover action.

        This corresponds to the ``tooltip`` attribute of the
        ``a:hlinkClick`` / ``a:hlinkHover`` element. Returns |None| when no
        hyperlink element is present or when the element carries no
        ``tooltip`` attribute.

        Assigning a string sets the tooltip (creating the hyperlink element
        if necessary). Assigning |None| or the empty string removes the
        ``tooltip`` attribute, leaving the hyperlink element itself
        untouched because it may still carry a URL, sound, or slide-jump
        action.
        """
        hlink = self._sound_hlink()
        if hlink is None:
            return None
        return hlink.tooltip

    @screen_tip.setter
    def screen_tip(self, value: str | None) -> None:
        if not value:
            hlink = self._sound_hlink()
            if hlink is None:
                return
            hlink.tooltip = None
            return
        if self._hover:
            hlink = cast(
                "CT_NonVisualDrawingProps", self._element
            ).get_or_add_hlinkHover()
        else:
            hlink = self._element.get_or_add_hlinkClick()
        hlink.tooltip = value

    @property
    def sound(self) -> Sound | None:
        """A |Sound| object for the WAV audio played on click/hover, or |None|.

        Returns |None| when no ``a:snd`` child element is present on the click/hover
        action. When present, the returned |Sound| provides read-only access to the
        sound name and the embedded audio blob.
        """
        hlink = self._sound_hlink()
        if hlink is None:
            return None
        snd = hlink.snd
        if snd is None:
            return None
        return Sound(snd, cast("SlidePart", self.part))

    def set_sound(
        self,
        audio_file: str | IO[bytes] | Audio,
        name: str | None = None,
        mime_type: str | None = None,
    ) -> Sound:
        """Attach a WAV audio clip to be played on this click/hover action.

        `audio_file` may be a filesystem path, a file-like binary stream, or a
        pre-built :class:`~pptx.media.Audio` instance. Any sound previously set on
        this action is removed (including its relationship when no other shape
        references the same audio part). The audio is embedded as a media part inside
        the package and an ``a:snd`` child element referencing it is added to the
        ``a:hlinkClick`` / ``a:hlinkHover`` element, creating that hyperlink element
        if necessary.

        `name` sets the ``name`` attribute on the ``a:snd`` element; when omitted it
        defaults to the base filename of the source file (or ``'sound.wav'`` for
        in-memory streams).

        Returns the newly-created |Sound| object.
        """
        if isinstance(audio_file, Audio):
            audio = audio_file
        else:
            audio = Audio.from_path_or_file_like(audio_file, mime_type)

        self.remove_sound()

        if self._hover:
            hlink = cast(
                "CT_NonVisualDrawingProps", self._element
            ).get_or_add_hlinkHover()
        else:
            hlink = self._element.get_or_add_hlinkClick()

        slide_part = cast("SlidePart", self.part)
        rId = slide_part.get_or_add_sound_media_part(audio)

        snd = hlink.get_or_add_snd()
        snd.rEmbed = rId
        snd.name = name if name is not None else audio.filename

        return Sound(snd, slide_part)

    def remove_sound(self) -> None:
        """Remove any ``a:snd`` sound from this click/hover action.

        Also drops the AUDIO relationship associated with the sound. Does nothing
        when no sound is currently set. The enclosing ``a:hlinkClick`` /
        ``a:hlinkHover`` element itself is preserved because it may still carry a
        separate hyperlink or action.
        """
        hlink = self._sound_hlink()
        if hlink is None:
            return
        snd = hlink.snd
        if snd is None:
            return
        rId = snd.rEmbed
        if rId:
            self.part.drop_rel(rId)
        hlink.remove(snd)

    def _sound_hlink(self) -> CT_Hyperlink | None:
        """`a:hlinkClick`/`a:hlinkHover` element for read-side access, or None.

        This parallels :attr:`_hlink` but skips the ``isinstance`` assert used
        there so callers (sound lookup, screen-tip lookup, etc.) can be tested
        and type-checked without materializing a full
        `CT_NonVisualDrawingProps` instance.
        """
        if self._hover:
            return getattr(self._element, "hlinkHover", None)
        return self._element.hlinkClick

    @property
    def target_slide(self) -> Slide | None:
        """
        A reference to the slide in this presentation that is the target of
        the slide jump action in this shape. Slide jump actions include
        `PP_ACTION.FIRST_SLIDE`, `LAST_SLIDE`, `NEXT_SLIDE`,
        `PREVIOUS_SLIDE`, and `NAMED_SLIDE`. Returns |None| for all other
        actions. In particular, the `LAST_SLIDE_VIEWED` action and the `PLAY`
        (start other presentation) actions are not supported.

        A slide object may be assigned to this property, which makes the
        shape an "internal hyperlink" to the assigened slide::

            slide, target_slide = prs.slides[0], prs.slides[1]
            shape = slide.shapes[0]
            shape.target_slide = target_slide

        Assigning |None| removes any slide jump action. Note that this is
        accomplished by removing any action present (such as a hyperlink),
        without first checking that it is a slide jump action.
        """
        slide_jump_actions = (
            PP_ACTION.FIRST_SLIDE,
            PP_ACTION.LAST_SLIDE,
            PP_ACTION.NEXT_SLIDE,
            PP_ACTION.PREVIOUS_SLIDE,
            PP_ACTION.NAMED_SLIDE,
        )

        if self.action not in slide_jump_actions:
            return None

        if self.action == PP_ACTION.FIRST_SLIDE:
            return self._slides[0]
        elif self.action == PP_ACTION.LAST_SLIDE:
            return self._slides[-1]
        elif self.action == PP_ACTION.NEXT_SLIDE:
            next_slide_idx = self._slide_index + 1
            if next_slide_idx >= len(self._slides):
                raise ValueError("no next slide")
            return self._slides[next_slide_idx]
        elif self.action == PP_ACTION.PREVIOUS_SLIDE:
            prev_slide_idx = self._slide_index - 1
            if prev_slide_idx < 0:
                raise ValueError("no previous slide")
            return self._slides[prev_slide_idx]
        elif self.action == PP_ACTION.NAMED_SLIDE:
            assert self._hlink is not None
            rId = self._hlink.rId
            slide_part = cast("SlidePart", self.part.related_part(rId))
            return slide_part.slide

    @target_slide.setter
    def target_slide(self, slide: Slide | None):
        self._clear_click_action()
        if slide is None:
            return
        hlink = self._element.get_or_add_hlinkClick()
        hlink.action = "ppaction://hlinksldjump"
        hlink.rId = self.part.relate_to(slide.part, RT.SLIDE)

    def _clear_click_action(self):
        """Remove any existing click action."""
        hlink = self._hlink
        if hlink is None:
            return
        rId = hlink.rId
        if rId:
            self.part.drop_rel(rId)
        # -- also drop any audio relationship carried by an `a:snd` child so we don't
        # -- leave a dangling AUDIO rel when the whole hyperlink element is removed
        snd = hlink.snd
        if snd is not None and snd.rEmbed:
            self.part.drop_rel(snd.rEmbed)
        self._element.remove(hlink)

    @property
    def _hlink(self) -> CT_Hyperlink | None:
        """
        Reference to the `a:hlinkClick` or `a:hlinkHover` element for this
        click action. Returns |None| if the element is not present.
        """
        if self._hover:
            assert isinstance(self._element, CT_NonVisualDrawingProps)
            return self._element.hlinkHover
        return self._element.hlinkClick

    @lazyproperty
    def _slide(self):
        """
        Reference to the slide containing the shape having this click action.
        """
        return self.part.slide

    @lazyproperty
    def _slide_index(self):
        """
        Position in the slide collection of the slide containing the shape
        having this click action.
        """
        return self._slides.index(self._slide)

    @lazyproperty
    def _slides(self) -> Slides:
        """
        Reference to the slide collection for this presentation.
        """
        return self.part.package.presentation_part.presentation.slides


class Hyperlink(Subshape):
    """Represents a hyperlink action on a shape or text run."""

    def __init__(
        self,
        xPr: CT_NonVisualDrawingProps | CT_TextCharacterProperties,
        parent: BaseShape,
        hover: bool = False,
    ):
        super(Hyperlink, self).__init__(parent)
        # xPr is either a cNvPr or rPr element
        self._element = xPr
        # _hover determines use of `a:hlinkClick` or `a:hlinkHover`
        self._hover = hover

    @property
    def address(self) -> str | None:
        """Read/write. The URL of the hyperlink.

        URL can be on http, https, mailto, or file scheme; others may work. Returns |None| if no
        hyperlink is defined, including when another action such as `RUN_MACRO` is defined on the
        object. Assigning |None| removes any action defined on the object, whether it is a hyperlink
        action or not.
        """
        hlink = self._hlink

        # there's no URL if there's no click action
        if hlink is None:
            return None

        # a click action without a relationship has no URL
        rId = hlink.rId
        if not rId:
            return None

        return self.part.target_ref(rId)

    @address.setter
    def address(self, url: str | None):
        # implements all three of add, change, and remove hyperlink
        self._remove_hlink()

        if url:
            rId = self.part.relate_to(url, RT.HYPERLINK, is_external=True)
            hlink = self._get_or_add_hlink()
            hlink.rId = rId

    def _get_or_add_hlink(self) -> CT_Hyperlink:
        """Get the `a:hlinkClick` or `a:hlinkHover` element for the Hyperlink object.

        The actual element depends on the value of `self._hover`. Create the element if not present.
        """
        if self._hover:
            return cast("CT_NonVisualDrawingProps", self._element).get_or_add_hlinkHover()
        return self._element.get_or_add_hlinkClick()

    @property
    def _hlink(self) -> CT_Hyperlink | None:
        """Reference to the `a:hlinkClick` or `h:hlinkHover` element for this click action.

        Returns |None| if the element is not present.
        """
        if self._hover:
            return cast("CT_NonVisualDrawingProps", self._element).hlinkHover
        return self._element.hlinkClick

    def _remove_hlink(self):
        """Remove the a:hlinkClick or a:hlinkHover element.

        Also drops any relationship it might have.
        """
        hlink = self._hlink
        if hlink is None:
            return
        rId = hlink.rId
        if rId:
            self.part.drop_rel(rId)
        # -- also drop any audio relationship carried by an `a:snd` child so we
        # -- don't leave a dangling AUDIO rel when the hyperlink is removed
        snd = hlink.snd
        if snd is not None and snd.rEmbed:
            self.part.drop_rel(snd.rEmbed)
        self._element.remove(hlink)


class Sound(object):
    """Represents an ``a:snd`` embedded WAV sound on a click or hover action.

    A |Sound| object provides read-only access to an existing click-action sound; use
    :meth:`ActionSetting.set_sound` to create or replace one and
    :meth:`ActionSetting.remove_sound` to delete it.
    """

    def __init__(self, snd: CT_EmbeddedWAVAudioFile, part: SlidePart):
        super(Sound, self).__init__()
        self._snd = snd
        self._part = part

    @property
    def blob(self) -> bytes:
        """The bytestream of the embedded audio "file"."""
        return self._media_part.blob

    @property
    def name(self) -> str:
        """The ``name`` attribute of the ``a:snd`` element (defaults to an empty string).

        This is the display name PowerPoint shows for the sound; it typically matches
        the original filename (for example ``'applause.wav'``).
        """
        return self._snd.name or ""

    @property
    def rId(self) -> str:
        """The ``r:embed`` relationship id pointing at the embedded audio part."""
        return self._snd.rEmbed

    @property
    def _media_part(self):
        """The |MediaPart| holding the audio binary referenced by this sound."""
        return self._part.related_part(self._snd.rEmbed)
