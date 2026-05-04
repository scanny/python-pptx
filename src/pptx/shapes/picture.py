"""Shapes based on the `p:pic` element, including Picture and Movie."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Literal, cast

from pptx.dml.line import LineFormat
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_MEDIA_TYPE
from pptx.media import Video
from pptx.oxml.ns import qn
from pptx.parts.media import MediaPart
from pptx.shapes.base import BaseShape
from pptx.shared import ParentedElementProxy
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.dml.fill import CT_AlphaModulateFixedEffect, CT_Blip
    from pptx.oxml.shapes.picture import CT_Picture
    from pptx.oxml.shapes.shared import CT_LineProperties
    from pptx.oxml.timing import CT_TLTimeCondition
    from pptx.parts.slide import SlidePart
    from pptx.types import ProvidesPart

MovieStartCondition = Literal["onClick", "withPrevious", "afterPrevious"]


class _BasePicture(BaseShape):
    """Base class for shapes based on a `p:pic` element."""

    def __init__(self, pic: CT_Picture, parent: ProvidesPart):
        super(_BasePicture, self).__init__(pic, parent)
        self._pic = pic

    @property
    def crop_bottom(self) -> float:
        """|float| representing relative portion cropped from shape bottom.

        Read/write. 1.0 represents 100%. For example, 25% is represented by 0.25. Negative values
        are valid as are values greater than 1.0.
        """
        return self._pic.srcRect_b

    @crop_bottom.setter
    def crop_bottom(self, value: float):
        self._pic.srcRect_b = value

    @property
    def crop_left(self) -> float:
        """|float| representing relative portion cropped from left of shape.

        Read/write. 1.0 represents 100%. A negative value extends the side beyond the image
        boundary.
        """
        return self._pic.srcRect_l

    @crop_left.setter
    def crop_left(self, value: float):
        self._pic.srcRect_l = value

    @property
    def crop_right(self) -> float:
        """|float| representing relative portion cropped from right of shape.

        Read/write. 1.0 represents 100%.
        """
        return self._pic.srcRect_r

    @crop_right.setter
    def crop_right(self, value: float):
        self._pic.srcRect_r = value

    @property
    def crop_top(self) -> float:
        """|float| representing relative portion cropped from shape top.

        Read/write. 1.0 represents 100%.
        """
        return self._pic.srcRect_t

    @crop_top.setter
    def crop_top(self, value: float):
        self._pic.srcRect_t = value

    def get_or_add_ln(self):
        """Return the `a:ln` element for this `p:pic`-based image.

        The `a:ln` element contains the line format properties XML.
        """
        return self._pic.get_or_add_ln()

    @lazyproperty
    def line(self) -> LineFormat:
        """Provides access to properties of the picture outline, such as its color and width."""
        return LineFormat(self)

    @property
    def ln(self) -> CT_LineProperties | None:
        """The `a:ln` element for this `p:pic`.

        Contains the line format properties such as line color and width. |None| if no `a:ln`
        element is present.
        """
        return self._pic.ln


class Movie(_BasePicture):
    """A movie shape, one that places a video on a slide.

    Like |Picture|, a movie shape is based on the `p:pic` element. A movie is composed of a video
    and a *poster frame*, the placeholder image that represents the video before it is played.
    """

    def delete(self) -> None:
        """Remove this movie from the slide and clean up its associated parts and timing node.

        A movie shape is backed by up to three slide-part relationships: the video / audio part
        (``a:videoFile/@r:link`` or ``a:audioFile/@r:link``), the ``p14:media`` part
        (``@r:embed``), and the poster-frame image part (``a:blip/@r:embed``). The matching
        relationships are dropped so unreferenced parts are garbage-collected on save (see
        :meth:`Picture.delete` for the rId-reference-count semantics).

        A movie also owns a ``p:video`` entry in the slide's ``p:timing`` sub-tree (added by
        :meth:`SlideShapes.add_movie`) — leaving it in place after the ``p:pic`` is removed
        produces a dangling ``p:spTgt/@spid`` reference that causes PowerPoint to flag the file
        as corrupt on open (issue #974). The matching ``p:video`` element is located by the
        shape's id and removed.

        .. versionadded:: 2026.05.0
        """
        # -- capture shape_id BEFORE removing the p:pic from the tree --
        shape_id = self.shape_id
        # -- drop the three media-related slide-part rels (video, media, poster-frame) --
        for rId in self._media_rIds:
            self.part.drop_rel(rId)
        # -- drop the p:video timing-tree entry that targets this shape --
        self._remove_video_timing(shape_id)
        super().delete()

    @property
    def _media_part(self) -> MediaPart | None:
        """The |MediaPart| holding this movie's embedded audio / video bytes.

        Located via the ``p14:media`` ``@r:embed`` (``RT.MEDIA`` rel) when
        present, falling back to the ``a:videoFile`` / ``a:audioFile``
        ``@r:link`` (``RT.VIDEO`` / ``RT.AUDIO`` rel). In a well-formed
        ``add_movie``-authored shape both rIds point at the same
        |MediaPart|. Returns |None| when neither descriptor carries an
        rId — e.g. a malformed pic with no media references at all.
        """
        pic = self._pic
        rId = pic.media_embed_rId or pic.media_video_rId
        if rId is None:
            return None
        part = self.part.related_part(rId)
        if not isinstance(part, MediaPart):
            return None
        return part

    @property
    def _media_rIds(self) -> list[str]:
        """The rIds referenced by this movie's media sub-elements.

        Returns a list containing the ``a:videoFile``/``a:audioFile`` ``@r:link``, the
        ``p14:media`` ``@r:embed``, and the poster-frame ``a:blip`` ``@r:embed`` — in
        whatever subset is actually present on this particular movie. Duplicates are
        preserved so :meth:`pptx.opc.package.XmlPart.drop_rel` sees the correct total
        reference count when the same rId appears in more than one of these slots.
        """
        rIds: list[str] = []
        rIds.extend(
            self._pic.xpath(
                "./p:nvPicPr/p:nvPr/a:videoFile/@r:link" " | ./p:nvPicPr/p:nvPr/a:audioFile/@r:link"
            )
        )
        rIds.extend(self._pic.xpath("./p:nvPicPr/p:nvPr/p:extLst/p:ext/p14:media/@r:embed"))
        blip_rId = self._pic.blip_rId
        if blip_rId is not None:
            rIds.append(blip_rId)
        return rIds

    def _remove_video_timing(self, shape_id: int) -> None:
        """Remove the `p:video` timing-tree entry targeting *shape_id*, if present.

        Searches the entire `p:sld` document for any ``p:video`` whose
        ``p:cMediaNode/p:tgtEl/p:spTgt/@spid`` matches *shape_id*. Using ``//p:video``
        (rather than the schema-canonical
        ``/p:sld/p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:video``) tolerates the
        ``mc:AlternateContent``-wrapped timing layout emitted by PowerPoint 2010+
        (issue #954) without duplicating the wrapper-aware traversal logic that
        already lives on ``CT_Slide``. When no matching ``p:video`` is found (e.g. the
        movie pre-existed in a file without a timing entry) this is a no-op.
        """
        videos = self._pic.xpath(
            "/p:sld//p:video[p:cMediaNode/p:tgtEl/p:spTgt/@spid='%d']" % shape_id
        )
        for video in videos:
            video.getparent().remove(video)

    @lazyproperty
    def media_format(self) -> _MediaFormat:
        """The |_MediaFormat| object for this movie.

        The |_MediaFormat| object provides access to formatting properties for the movie.
        """
        return _MediaFormat(self._pic, self)

    @property
    def media_type(self) -> PP_MEDIA_TYPE:
        """Member of :ref:`PpMediaType` describing this shape.

        The return value is unconditionally `PP_MEDIA_TYPE.MOVIE` in this case.
        """
        return PP_MEDIA_TYPE.MOVIE

    @property
    def blob(self) -> bytes | None:
        """Bytes of the embedded media (video or audio) or |None| when absent.

        This is the raw binary payload of the underlying :class:`.MediaPart`
        referenced by the shape — for example the MP4 / MP3 / WAV bytes
        written into ``ppt/media/mediaN.ext`` inside the package. Returns
        |None| when no media part is associated with this shape (e.g. the
        shape was loaded from a malformed file whose media relationships
        are missing).

        Mirrors :attr:`Picture.image` + :attr:`Image.blob` for the movie
        case; useful for extracting embedded audio / video from an existing
        presentation without going through the package internals.

        .. versionadded:: 2026.05.0
        """
        media_part = self._media_part
        if media_part is None:
            return None
        return media_part.blob

    @property
    def content_type(self) -> str | None:
        """MIME type of the embedded media, e.g. ``"video/mp4"`` or |None|.

        Read from the underlying :class:`.MediaPart`'s ``Content-Type``
        (the value recorded in ``[Content_Types].xml`` for the media
        partname). Returns |None| when no media part is associated with
        this shape.

        .. versionadded:: 2026.05.0
        """
        media_part = self._media_part
        if media_part is None:
            return None
        return media_part.content_type

    @property
    def ext(self) -> str | None:
        """File extension of the embedded media, e.g. ``"mp4"`` or |None|.

        The returned extension does not include the leading period and is
        read from the media part's partname (``ppt/media/mediaN.<ext>``) —
        i.e. the extension PowerPoint actually wrote into the package, which
        is chosen by :meth:`Video.ext` / :meth:`Audio.ext` at add-time based
        on the supplied filename or MIME type. Returns |None| when no media
        part is associated with this shape.

        .. versionadded:: 2026.05.0
        """
        media_part = self._media_part
        if media_part is None:
            return None
        return media_part.partname.ext

    @property
    def poster_frame(self):
        """Return |Image| object containing poster frame for this movie.

        Returns |None| if this movie has no poster frame (uncommon).
        """
        slide_part, rId = self.part, self._pic.blip_rId
        if rId is None:
            return None
        return slide_part.get_image(rId)

    def replace_media(
        self,
        new_path_or_file: str | IO[bytes],
        mime_type: str | None = None,
    ) -> None:
        """Swap the underlying audio/video binary, keeping this shape in place.

        Loads the media at `new_path_or_file` (a filesystem path or a binary
        file-like object) into a new |MediaPart| and rewrites this shape's
        ``a:videoFile`` / ``a:audioFile`` ``@r:link`` and ``p14:media``
        ``@r:embed`` rIds to point at it. Position, size, cropping, poster
        frame, hyperlink, and ``p:timing`` entries on the shape are
        unchanged — only the media bytes behind the shape are replaced.

        Parameters
        ----------
        new_path_or_file
            Either a filesystem path (``str``) or a binary file-like object
            containing the replacement audio or video. Binary file-like
            objects are read to EOF.
        mime_type
            Optional hint like ``"audio/mpeg"`` or ``"video/mp4"``. When
            omitted, the package falls back to ``video/unknown``; providing
            the correct MIME type helps downstream viewers pick the right
            codec.

        Notes
        -----
        The shape's media-element tag (``a:videoFile`` vs ``a:audioFile``) is
        preserved. If you replace audio media with video media (or vice
        versa) the tag will *not* change — callers that need to switch
        modality should use :meth:`SlideShapes.add_movie` to create a fresh
        shape. The old media part becomes eligible for garbage-collection on
        save once nothing else references its rIds.

        .. versionadded:: 2026.05.0
        """
        # -- read the new media file into a Video value object (handles both
        #    audio and video payloads via the same MediaPart pipeline) --
        new_video = Video.from_path_or_file_like(new_path_or_file, mime_type)

        pic = self._pic
        old_video_rId = pic.media_video_rId
        old_media_rId = pic.media_embed_rId
        if old_video_rId is None or old_media_rId is None:
            raise ValueError(
                "cannot replace_media: this shape is not a media pic (no "
                "a:videoFile/a:audioFile or p14:media descriptor found)"
            )

        # -- add the new media part to the package and get its rIds --
        slide_part: SlidePart = self.part  # pyright: ignore[reportAssignmentType]
        new_media_rId, new_video_rId = slide_part.get_or_add_video_media_part(new_video)

        # -- rewrite the pic XML to point at the new rIds --
        pic.media_video_rId = new_video_rId
        pic.media_embed_rId = new_media_rId

        # -- drop old rels; drop_rel is a no-op when the rId is still referenced
        #    (e.g. same media shared by two shapes, or new == old) --
        if old_video_rId != new_video_rId:
            slide_part.drop_rel(old_video_rId)
        if old_media_rId != new_media_rId:
            slide_part.drop_rel(old_media_rId)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Return member of :ref:`MsoShapeType` describing this shape.

        The return value is unconditionally `MSO_SHAPE_TYPE.MEDIA` in this
        case.
        """
        return MSO_SHAPE_TYPE.MEDIA

    @property
    def start_condition(self) -> MovieStartCondition:
        """The trigger that causes this movie to begin playing.

        Read/write. One of ``"onClick"`` (default), ``"withPrevious"``, or
        ``"afterPrevious"``. The value is stored on the first
        ``p:stCondLst/p:cond`` of the ``p:video`` time-node in the slide's
        ``p:timing`` tree — the node added by :meth:`SlideShapes.add_movie`.

        Returns ``"onClick"`` when no timing node is found (e.g. the movie
        was loaded from a malformed slide) or when the condition encodes
        PowerPoint's default click-triggered playback (``delay="indefinite"``
        with no ``evt``).
        """
        cond = self._video_cond
        if cond is None:
            return "onClick"
        evt = cond.evt
        delay = cond.delay
        if evt == "onClick" or (evt is None and delay == "indefinite"):
            return "onClick"
        if evt == "onEnd":
            return "afterPrevious"
        # -- evt is None and delay is numeric (or absent) -> with previous --
        return "withPrevious"

    @start_condition.setter
    def start_condition(self, value: MovieStartCondition):
        if value not in ("onClick", "withPrevious", "afterPrevious"):
            raise ValueError(
                "start_condition must be 'onClick', 'withPrevious', or 'afterPrevious', "
                "got %r" % (value,)
            )
        cond = self._get_or_add_video_cond()
        if cond is None:
            raise ValueError(
                "cannot set start_condition: this movie has no associated p:video "
                "timing node; was it created with SlideShapes.add_movie?"
            )
        # -- preserve existing numeric delay; replace "indefinite" with None so
        #    the setter logic below re-computes a sensible default --
        existing_delay = cond.delay if isinstance(cond.delay, int) else None
        if value == "onClick":
            cond.evt = None
            cond.delay = "indefinite" if existing_delay in (None, 0) else existing_delay
        elif value == "withPrevious":
            cond.evt = None
            cond.delay = existing_delay if existing_delay is not None else 0
        else:  # afterPrevious
            cond.evt = "onEnd"
            cond.delay = existing_delay if existing_delay is not None else 0

    @property
    def start_time(self) -> float | None:
        """Delay in seconds between the start trigger and the video beginning to play.

        Read/write. Returns |None| when the delay is ``"indefinite"`` (the
        default, meaning "wait for user click") or when no timing node is
        present; otherwise a ``float`` number of seconds (``p:cond/@delay``
        divided by 1000).

        Assigning a ``float`` or ``int`` writes the value to ``p:cond/@delay``
        in milliseconds. Assigning |None| writes ``"indefinite"`` — which
        PowerPoint interprets as "pause here until clicked". Note that the
        PowerPoint UI typically pairs a non-``None`` ``start_time`` with a
        non-``"onClick"`` ``start_condition``; the two properties are
        orthogonal in the XML but the combination ``start_condition="onClick"``
        + numeric ``start_time`` is unusual.
        """
        cond = self._video_cond
        if cond is None:
            return None
        delay = cond.delay
        if delay is None or delay == "indefinite":
            return None
        return delay / 1000.0

    @start_time.setter
    def start_time(self, value: float | int | None):
        cond = self._get_or_add_video_cond()
        if cond is None:
            raise ValueError(
                "cannot set start_time: this movie has no associated p:video "
                "timing node; was it created with SlideShapes.add_movie?"
            )
        if value is None:
            cond.delay = "indefinite"
            return
        if not isinstance(value, (int, float)) or value < 0:
            raise ValueError(
                "start_time must be a non-negative number of seconds or None, got %r" % (value,)
            )
        cond.delay = int(round(value * 1000))

    @property
    def _video_cond(self) -> CT_TLTimeCondition | None:
        """The first `p:cond` under this movie's `p:video/p:cMediaNode/p:cTn/p:stCondLst`.

        Returns |None| when the movie has no timing node (pre-existing file
        without a ``p:video`` entry, or a movie not added via
        :meth:`SlideShapes.add_movie`).
        """
        # -- locate by spTgt/@spid match; the p:video element contains a
        #    p:tgtEl/p:spTgt whose @spid equals this movie's shape_id --
        shape_id = self.shape_id
        conds = self._pic.xpath(
            "/p:sld/p:timing//p:video"
            "[p:cMediaNode/p:tgtEl/p:spTgt/@spid='%d']"
            "/p:cMediaNode/p:cTn/p:stCondLst/p:cond" % shape_id
        )
        return conds[0] if conds else None

    def _get_or_add_video_cond(self) -> CT_TLTimeCondition | None:
        """Get or add the first `p:cond` child of this movie's stCondLst.

        Returns |None| if no ``p:video`` node exists for this movie — i.e.
        the stCondLst's parent ``p:cTn`` is missing. We don't synthesize a
        whole ``p:video`` subtree on demand because allocating a fresh
        ``p:cTn`` id requires slide-level context that isn't cleanly
        available from a bare ``Movie`` proxy.
        """
        shape_id = self.shape_id
        cTns = self._pic.xpath(
            "/p:sld/p:timing//p:video"
            "[p:cMediaNode/p:tgtEl/p:spTgt/@spid='%d']"
            "/p:cMediaNode/p:cTn" % shape_id
        )
        if not cTns:
            return None
        cTn = cTns[0]
        stCondLst = cTn.get_or_add_stCondLst()
        conds = stCondLst.findall(qn("p:cond"))
        if conds:
            return conds[0]
        return stCondLst.add_cond()


class Picture(_BasePicture):
    """A picture shape, one that places an image on a slide.

    Based on the `p:pic` element.
    """

    def delete(self) -> None:
        """Remove this picture from the slide and drop its image part if unreferenced.

        Removes the `p:pic` element from its parent `p:spTree`. When the embedded image is not
        referenced elsewhere in the slide part (reference count < 2), the matching relationship
        is also removed, which allows the image part to be garbage-collected on save if no other
        part refers to it.

        .. versionadded:: 2026.05.0
        """
        rId = self._pic.blip_rId
        if rId is not None:
            self.part.drop_rel(rId)
        super().delete()

    @property
    def auto_shape_type(self) -> MSO_SHAPE | None:
        """Member of MSO_SHAPE indicating masking shape.

        A picture can be masked by any of the so-called "auto-shapes" available in PowerPoint,
        such as an ellipse or triangle. When a picture is masked by a shape, the shape assumes the
        same dimensions as the picture and the portion of the picture outside the shape boundaries
        does not appear. Note the default value for a newly-inserted picture is
        `MSO_AUTO_SHAPE_TYPE.RECTANGLE`, which performs no cropping because the extents of the
        rectangle exactly correspond to the extents of the picture.

        The available shapes correspond to the members of :ref:`MsoAutoShapeType`.

        The return value can also be |None|, indicating the picture either has no geometry (not
        expected) or has custom geometry, like a freeform shape. A picture with no geometry will
        have no visible representation on the slide, although it can be selected. This is because
        without geometry, there is no "inside-the-shape" for it to appear in.
        """
        prstGeom = self._pic.spPr.prstGeom
        if prstGeom is None:  # ---generally means cropped with freeform---
            return None
        return prstGeom.prst

    @auto_shape_type.setter
    def auto_shape_type(self, member: MSO_SHAPE):
        MSO_SHAPE.validate(member)
        spPr = self._pic.spPr
        prstGeom = spPr.prstGeom
        if prstGeom is None:
            spPr._remove_custGeom()  # pyright: ignore[reportPrivateUsage]
            prstGeom = spPr._add_prstGeom()  # pyright: ignore[reportPrivateUsage]
        prstGeom.prst = member

    @property
    def image(self):
        """The |Image| object for this picture, or |None| when no image is embedded.

        Provides access to the properties and bytes of the image in this picture shape.
        Returns |None| when the shape has no embedded image — either because the
        ``a:blip`` element lacks an ``r:embed`` reference, or because the
        ``p:blipFill`` sub-element itself is absent (issue #434 — graceful
        handling of shapes whose image data was stripped by another tool).
        """
        rId = self._pic.blip_rId
        if rId is None:
            return None
        return self.part.get_image(rId)

    def replace_image(self, image_file: str | IO[bytes]) -> None:
        """Swap the embedded image with the one in `image_file`.

        `image_file` can be either a path to a file (a string) or a file-like
        object. Position, size, rotation, cropping, masking shape, outline,
        and any other shape-level formatting are preserved — only the pixel
        bytes of the image are replaced.

        The new image is added to the package (reusing an existing image
        part when one with identical content is already present). The
        relationship to the previous image part is dropped if no other
        reference to it remains in this slide part, allowing the old image
        part to be garbage-collected on save when it is otherwise orphaned.

        Raises |ValueError| if this picture has no embedded image (i.e.
        its ``a:blip`` element lacks an ``r:embed`` reference); such a
        picture is malformed and has no "current" image to replace.

        .. versionadded:: 2026.05.0
        """
        blipFill = self._pic.blipFill
        blip = blipFill.blip if blipFill is not None else None
        if blip is None or blip.rEmbed is None:
            raise ValueError("no embedded image to replace")
        old_rId = blip.rEmbed

        slide_part = self.part
        _, new_rId = slide_part.get_or_add_image_part(image_file)

        # -- rebind a:blip/@r:embed first so the XML reference count seen by --
        # -- drop_rel reflects the replacement (post-swap the old rId's count --
        # -- drops by one; drop_rel keeps the rel iff any references remain). --
        blip.rEmbed = new_rId

        if new_rId != old_rId:
            slide_part.drop_rel(old_rId)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        """Unconditionally `MSO_SHAPE_TYPE.PICTURE` in this case."""
        return MSO_SHAPE_TYPE.PICTURE

    @property
    def transparency(self) -> float:
        """Picture transparency as a percentage in ``[0.0, 100.0]``.

        Corresponds to the ``Picture Transparency`` slider in PowerPoint's
        Format Picture pane. ``0.0`` is fully opaque (the default) and
        ``100.0`` is fully transparent.

        Read/write. Returns ``0.0`` when no ``a:alphaModFix`` child is
        present on the picture's ``a:blip`` element — the XML "no effect
        applied" encoding for full opacity.

        Assigning a value in ``[0.0, 100.0]`` writes
        ``p:blipFill/a:blip/a:alphaModFix@amt`` with the complementary
        alpha-amount (``(100 - transparency) * 1000`` in the XML, since
        ``amt`` encodes *remaining* opacity). Assigning ``0`` or ``0.0``
        removes any existing ``a:alphaModFix`` child.

        Raises ``ValueError`` when the shape has no ``a:blip`` element
        (malformed picture whose image data has been stripped) or when
        the value is outside ``[0.0, 100.0]``.

        .. versionadded:: 2026.05.0
        """
        blipFill = self._pic.blipFill
        blip = cast("CT_Blip | None", blipFill.blip if blipFill is not None else None)
        alphaModFix = blip.alphaModFix if blip is not None else None
        if alphaModFix is None:
            return 0.0
        # -- amt is the remaining-opacity fraction (1.0 == fully opaque) --
        return (1.0 - cast(float, alphaModFix.amt)) * 100.0

    @transparency.setter
    def transparency(self, value: float) -> None:
        if not isinstance(value, (int, float)):  # pyright: ignore[reportUnnecessaryIsInstance]
            raise ValueError(
                "transparency must be a number in range 0.0 to 100.0, got %r" % (value,)
            )
        if value < 0.0 or value > 100.0:
            raise ValueError("transparency must be in range 0.0 to 100.0, got %r" % (value,))
        blipFill = self._pic.blipFill
        if blipFill is None:
            raise ValueError("cannot set transparency on a picture with no blipFill")
        blip = cast("CT_Blip | None", blipFill.blip)
        if blip is None:
            raise ValueError("cannot set transparency on a picture with no a:blip element")
        # -- 0 (or 0.0) removes any existing a:alphaModFix --
        if value == 0:
            # -- xmlchemy generates `_remove_alphaModFix` / `get_or_add_alphaModFix`
            # -- on the ZeroOrOne; those dynamic methods aren't visible to pyright --
            blip._remove_alphaModFix()  # pyright: ignore[reportAttributeAccessIssue, reportUnknownMemberType]
            return
        # -- amt encodes remaining opacity as a fraction in [0.0, 1.0]. --
        alphaModFix = cast(
            "CT_AlphaModulateFixedEffect",
            blip.get_or_add_alphaModFix(),  # pyright: ignore[reportAttributeAccessIssue, reportUnknownMemberType]
        )
        alphaModFix.amt = (100.0 - float(value)) / 100.0


class _MediaFormat(ParentedElementProxy):
    """Provides access to formatting properties for a Media object.

    Media format properties are things like start point, volume, and
    compression type.
    """
