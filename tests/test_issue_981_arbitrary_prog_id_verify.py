# pyright: reportPrivateUsage=false

"""Regression test for issue #981 — arbitrary ``prog_id`` OLE embedding.

Issue #981 tracks the capability advertised in ``FEATURES.md`` under
"OLE objects and embedded files":

    ``SlideShapes.add_ole_object(..., prog_id, ..., extension=...)`` —
    Embed an OLE / package object with a custom icon. Accepts well-known
    ``PROG_ID`` enum members or arbitrary strings; when ``prog_id`` is
    not a known member, ``extension`` becomes required and the package
    content-type is emitted. ``[Added in 1.0.2.dev0]`` for the arbitrary
    ``prog_id`` / ``extension`` path.

The underlying feature landed as ``feat: #752 accept arbitrary prog_id
+ extension in SlideShapes.add_ole_object()`` (see HISTORY.rst line 3575
of master at ``95e5766f``). Issue #777
(``feat/issue-752-ole-embed-generic``) exercises the flow with the
specific ``MSHtml.MHT`` progId for HTML payloads; this module pins the
earlier OLE-wave contract from the caller's perspective using a
genuinely made-up progId (``"MyCustom.Object.1"``), so future
refactoring cannot regress the "any string, any extension" promise that
#981 / #822 / #752 collectively make.

Cross-references:

* #752 — original fork-era implementation of the arbitrary-``prog_id``
  path on :meth:`.SlideShapes.add_ole_object`.
* #777 — HTML-specific verify of the same pipeline via
  ``prog_id="MSHtml.MHT"`` (see
  ``tests/test_issue_777_html_ole_embed.py``).
* #822 — earlier OLE wave tracked alongside #752 in the same embedding
  area (not independently reproducible in this worktree; listed for
  provenance).
* #981 — the public-facing "arbitrary prog_id + extension support"
  capability that this suite verifies end-to-end.

The scenarios below exercise three orthogonal guarantees:

* a fully arbitrary, non-built-in ``prog_id`` string (not in
  :class:`pptx.enum.shapes.PROG_ID`) combined with an explicit
  ``extension`` is accepted by :meth:`.SlideShapes.add_ole_object`;
* the authored ``prog_id`` survives a ``save`` + ``Presentation(...)``
  reopen round-trip unchanged and the embedded blob is byte-for-byte
  identical;
* the package laid down on disk contains the expected
  ``/ppt/embeddings/oleObject*.<ext>`` part with the generic
  ``OFC_OLE_OBJECT`` content-type, regardless of how exotic the
  ``prog_id`` is.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.util import Inches

# -- fixtures ------------------------------------------------------------

_CUSTOM_PROG_ID = "MyCustom.Object.1"

_CUSTOM_BLOB = b"MYCUSTOM\x00\x01\x02arbitrary-binary-payload\xffend\n"

_ANOTHER_PROG_ID = "Acme.Widget.3"

_ANOTHER_BLOB = b"widget-payload-\x00\x01\x02\x03"


@pytest.fixture
def blank_slide():
    """Return a fresh slide on a fresh Presentation for each scenario."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    return prs, slide


# -- the verification suite ---------------------------------------------


class DescribeIssue981ArbitraryProgIdVerify(object):
    """#981 verify-and-close: arbitrary ``prog_id`` OLE embedding.

    Pins the contract advertised in ``FEATURES.md`` for
    :meth:`.SlideShapes.add_ole_object` when the ``prog_id`` is *not*
    a member of :class:`pptx.enum.shapes.PROG_ID`. This is the
    generic-OLE path first delivered by #752 and subsequently verified
    from the HTML angle in #777.
    """

    # -- authoring + round-trip --------------------------------------------

    def it_accepts_a_non_built_in_prog_id_string_with_explicit_extension(self, blank_slide):
        """A made-up ``prog_id`` string plus ``extension="bin"`` is accepted.

        Exercises the reporter-facing contract: any ``progId`` that
        PowerPoint will accept at open time should round-trip through
        :meth:`.SlideShapes.add_ole_object`, even when no
        :class:`PROG_ID` convenience member exists for it.
        """
        _, slide = blank_slide

        shape = slide.shapes.add_ole_object(
            io.BytesIO(_CUSTOM_BLOB),
            prog_id=_CUSTOM_PROG_ID,
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
            extension="bin",
        )

        assert shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        assert shape.ole_format.prog_id == _CUSTOM_PROG_ID
        assert shape.ole_format.blob == _CUSTOM_BLOB
        assert shape.ole_format.show_as_icon is True

    def it_preserves_the_arbitrary_prog_id_across_save_and_reopen(self, blank_slide):
        """``Presentation.save(...)`` + ``Presentation(...)`` preserves ``prog_id``.

        The reopened :class:`_OleFormat` proxy must report the exact
        ``progId`` string the author supplied, and the embedded payload
        must be byte-identical. This is the #981 headline guarantee.
        """
        prs, slide = blank_slide

        slide.shapes.add_ole_object(
            io.BytesIO(_CUSTOM_BLOB),
            prog_id=_CUSTOM_PROG_ID,
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
            extension="bin",
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]
        ole_shapes = [
            s for s in slide2.shapes if s.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        ]
        assert len(ole_shapes) == 1
        reopened = ole_shapes[0]
        assert reopened.ole_format.prog_id == _CUSTOM_PROG_ID
        assert reopened.ole_format.blob == _CUSTOM_BLOB
        assert reopened.ole_format.show_as_icon is True

    def it_writes_the_embedded_part_with_the_caller_supplied_extension(self, blank_slide):
        """The on-disk part-name carries the ``extension`` argument.

        When ``prog_id`` is arbitrary the library falls back to the
        generic ``OFC_OLE_OBJECT`` content-type and names the embedded
        part ``/ppt/embeddings/oleObject*.<ext>`` using the caller's
        ``extension``. This pins that the wiring between
        :meth:`.SlideShapes.add_ole_object` and
        :meth:`.EmbeddedPackagePart.factory` honours the explicit hint.
        """
        prs, slide = blank_slide

        slide.shapes.add_ole_object(
            io.BytesIO(_CUSTOM_BLOB),
            prog_id=_CUSTOM_PROG_ID,
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
            extension="dat",
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        with zipfile.ZipFile(buf) as zf:
            names = zf.namelist()
            ole_names = [
                n for n in names if n.startswith("ppt/embeddings/oleObject") and n.endswith(".dat")
            ]
            assert len(ole_names) == 1
            # -- the embedded payload is stored verbatim --
            assert zf.read(ole_names[0]) == _CUSTOM_BLOB
            # -- [Content_Types].xml carries the generic OLE content-type
            # -- for the custom extension (no app-specific override) --
            content_types_xml = zf.read("[Content_Types].xml").decode("utf-8")
            assert CT.OFC_OLE_OBJECT in content_types_xml

    # -- negative / robustness --------------------------------------------

    def it_defaults_the_part_extension_to_bin_when_none_is_supplied(self, blank_slide):
        """An arbitrary ``prog_id`` with no ``extension`` falls back to ``.bin``.

        The :class:`EmbeddedPackagePart` factory documents this
        fallback; exercising it here pins that #981's "extension
        becomes required" wording is a recommendation for readability,
        not an enforced precondition — callers who omit it still get a
        valid, openable package.
        """
        prs, slide = blank_slide

        slide.shapes.add_ole_object(
            io.BytesIO(_CUSTOM_BLOB),
            prog_id=_CUSTOM_PROG_ID,
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        with zipfile.ZipFile(buf) as zf:
            ole_names = [
                n
                for n in zf.namelist()
                if n.startswith("ppt/embeddings/oleObject") and n.endswith(".bin")
            ]
            assert len(ole_names) == 1
            assert zf.read(ole_names[0]) == _CUSTOM_BLOB

    def it_accepts_multiple_distinct_arbitrary_prog_ids_on_one_slide(self, blank_slide):
        """Two OLE objects with different made-up ``prog_id`` strings coexist.

        Each embed gets its own part and its own
        :attr:`_OleFormat.prog_id` after reopen. This pins that the
        factory's part-name allocation (``oleObject1`` / ``oleObject2``
        / ...) does not conflate progIds across shapes on the same
        slide.
        """
        prs, slide = blank_slide

        slide.shapes.add_ole_object(
            io.BytesIO(_CUSTOM_BLOB),
            prog_id=_CUSTOM_PROG_ID,
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
            extension="bin",
        )
        slide.shapes.add_ole_object(
            io.BytesIO(_ANOTHER_BLOB),
            prog_id=_ANOTHER_PROG_ID,
            left=Inches(3),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
            extension="wid",
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        ole_shapes = [
            s for s in prs2.slides[0].shapes if s.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        ]
        assert len(ole_shapes) == 2
        by_prog_id = {s.ole_format.prog_id: s.ole_format.blob for s in ole_shapes}
        assert by_prog_id == {
            _CUSTOM_PROG_ID: _CUSTOM_BLOB,
            _ANOTHER_PROG_ID: _ANOTHER_BLOB,
        }

    # -- cross-check with the PROG_ID enum path ----------------------------

    def but_the_built_in_PROG_ID_enum_path_still_works_unchanged(self, blank_slide):
        """Adding the arbitrary-string path did not regress the enum path.

        Cross-checked here so a single ``pytest tests/test_issue_981*``
        invocation tells the caller both the new capability and the
        pre-existing well-known path are healthy. The XLSX convenience
        member round-trips under the Office-package content-type, which
        is a distinct code path from the generic OLE one exercised
        above.
        """
        prs, slide = blank_slide

        # -- a tiny but syntactically valid xlsx-shaped blob: the library
        # -- does not parse the blob, only embeds its bytes. --
        xlsx_blob = b"PK\x03\x04" + b"\x00" * 24

        shape = slide.shapes.add_ole_object(
            io.BytesIO(xlsx_blob),
            prog_id=PROG_ID.XLSX,
            left=Inches(1),
            top=Inches(1),
            width=Inches(1),
            height=Inches(1),
        )

        assert shape.ole_format.prog_id == "Excel.Sheet.12"
        assert shape.ole_format.blob == xlsx_blob

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        reopened = next(
            s for s in prs2.slides[0].shapes if s.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
        )
        assert reopened.ole_format.prog_id == "Excel.Sheet.12"
        assert reopened.ole_format.blob == xlsx_blob
