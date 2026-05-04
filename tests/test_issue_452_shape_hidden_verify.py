# pyright: reportPrivateUsage=false

"""Verify-close regression test for issue #452 — hide/show a shape.

Issue #452 (https://github.com/scanny/python-pptx/issues/452) asked for a
supported way to toggle a shape's visibility without deleting it — the
PowerPoint "Selection Pane > eye-icon" idiom. The reporter wanted to
flip a shape on or off programmatically so the same slide master could
drive several output variants (e.g. an internal vs. external deck) by
hiding annotation shapes on the internal-only variant.

The feature is resolved by the #971 wave which shipped
``BaseShape.is_hidden`` — a read/write :class:`bool` mapping the
``hidden`` attribute on the shape's ``cNvPr`` element. Every shape type
(``p:sp``, ``p:pic``, ``p:cxnSp``, ``p:graphicFrame``, ``p:grpSp``)
surfaces the same property because ``cNvPr`` is the non-visual drawing
props element shared by all of them. Assigning ``True`` writes
``hidden="1"``; assigning ``False`` removes the attribute so the XML
round-trips to the schema default (``hidden="false"``, i.e. visible).

This suite is the verify-close pass. It exercises the real public API
against a fresh :class:`~pptx.presentation.Presentation` — no mocks —
and covers the reporter's end-to-end workflow:

* Default visible state (attribute absent, reads ``False``).
* Hide a shape and pin the XML surface that PowerPoint consumes
  (``cNvPr/@hidden="1"``) on each ``nv*Pr`` parent variant.
* Unhide clears the attribute so the XML matches the default.
* :meth:`Presentation.save` + reopen preserves the hidden flag.
* Multiple shapes are independent — hiding one does not affect
  siblings in the same shape tree.
* Every shape kind (``Shape``, ``Picture``, ``Connector``,
  ``GroupShape``, ``GraphicFrame``) supports the toggle.
"""

from __future__ import annotations

import io

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches


class DescribeIssue452ShapeHiddenVerify:
    """Verify-close regression suite for ``BaseShape.is_hidden`` (issue #452)."""

    def it_reports_False_as_the_default_visible_state(self):
        # -- a freshly added shape has no cNvPr/@hidden attribute, so
        # -- is_hidden reads as False. This pins the default against a
        # -- regression that wrote hidden="0" or hidden="false" eagerly.
        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
        )

        assert rect.is_hidden is False
        # -- attribute is absent; not merely set to "0" --
        cNvPr = rect._element.xpath("./p:nvSpPr/p:cNvPr")[0]
        assert cNvPr.get("hidden") is None

    def it_sets_hidden_true_and_emits_the_cNvPr_hidden_attribute(self):
        # -- the #452 reporter's canonical flow: flip is_hidden to True,
        # -- confirm the underlying XML attribute PowerPoint consumes is
        # -- set to "1" on p:nvSpPr/p:cNvPr.
        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
        )

        rect.is_hidden = True

        assert rect.is_hidden is True
        cNvPr = rect._element.xpath("./p:nvSpPr/p:cNvPr")[0]
        assert cNvPr.get("hidden") == "1"

    def it_clears_hidden_when_reset_to_False(self):
        # -- unhiding a previously-hidden shape removes the attribute so
        # -- the XML matches the schema default rather than writing an
        # -- explicit hidden="0" (PowerPoint reads both, but this pins
        # -- that python-pptx keeps the serialization minimal).
        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
        )
        rect.is_hidden = True
        assert rect.is_hidden is True

        rect.is_hidden = False

        assert rect.is_hidden is False
        cNvPr = rect._element.xpath("./p:nvSpPr/p:cNvPr")[0]
        assert cNvPr.get("hidden") is None

    def it_preserves_the_hidden_flag_across_save_and_reopen(self):
        # -- end-to-end: hiding a shape must survive a Presentation.save
        # -- + reopen round-trip. This is the pin that guards against a
        # -- serializer that drops the cNvPr/@hidden attribute or reads
        # -- it from the wrong nv*Pr parent after reload.
        prs, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
        )
        rect.is_hidden = True

        reloaded = _roundtrip(prs)
        reloaded_shape = reloaded.slides[0].shapes[0]

        assert reloaded_shape.is_hidden is True
        cNvPr = reloaded_shape._element.xpath("./p:nvSpPr/p:cNvPr")[0]
        assert cNvPr.get("hidden") in ("1", "true")

    def it_hides_one_shape_without_affecting_its_siblings(self):
        # -- hiding is per-shape; two sibling shapes in the same shape
        # -- tree must remain independent. This pins against a
        # -- regression where the setter leaked the attribute onto the
        # -- wrong shape (e.g. via a shared cNvPr reference bug).
        _, slide = _fresh_slide()
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
        )
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4), Inches(1), Inches(2), Inches(2))
        triangle = slide.shapes.add_shape(
            MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(4), Inches(2), Inches(2)
        )

        oval.is_hidden = True

        assert rect.is_hidden is False
        assert oval.is_hidden is True
        assert triangle.is_hidden is False
        # -- the attribute is set only on the oval's cNvPr --
        assert rect._element.xpath("./p:nvSpPr/p:cNvPr")[0].get("hidden") is None
        assert oval._element.xpath("./p:nvSpPr/p:cNvPr")[0].get("hidden") == "1"
        assert triangle._element.xpath("./p:nvSpPr/p:cNvPr")[0].get("hidden") is None

    def it_toggles_hidden_on_every_shape_kind(self):
        # -- the #971 feature promises coverage on every shape kind that
        # -- carries a cNvPr (p:sp, p:pic, p:cxnSp, p:grpSp,
        # -- p:graphicFrame). This pins one representative call per
        # -- shape kind — the authoring site that #452 callers will hit.
        _, slide = _fresh_slide()

        sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))
        pic = slide.shapes.add_picture(
            "tests/test_files/python-powered.png",
            Inches(3),
            Inches(1),
            Inches(1),
            Inches(1),
        )
        cxn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(3), Inches(3), Inches(3)
        )
        # -- must wrap existing sp / pic before recalculation or the
        # -- group computes an empty bbox. Create distinct children so
        # -- the group has shapes to wrap.
        g_child_a = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(5), Inches(1), Inches(1), Inches(1)
        )
        g_child_b = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(5), Inches(3), Inches(1), Inches(1)
        )
        grp = slide.shapes.add_group_shape([g_child_a, g_child_b])
        tbl = slide.shapes.add_table(2, 2, Inches(1), Inches(5), Inches(3), Inches(2))

        # -- default: everything visible --
        for shape in (sp, pic, cxn, grp, tbl):
            assert shape.is_hidden is False

        # -- hide each one in turn, confirm toggle takes and the
        # -- expected cNvPr attribute is present under the shape's
        # -- native nv*Pr parent. BaseShape is unhashable, so carry the
        # -- parent-tag pair as a list of tuples.
        nv_parent = [
            (sp, "p:nvSpPr"),
            (pic, "p:nvPicPr"),
            (cxn, "p:nvCxnSpPr"),
            (grp, "p:nvGrpSpPr"),
            (tbl, "p:nvGraphicFramePr"),
        ]
        for shape, parent_tag in nv_parent:
            shape.is_hidden = True
            assert shape.is_hidden is True
            cNvPr_nodes = shape._element.xpath("./" + parent_tag + "/p:cNvPr")
            assert len(cNvPr_nodes) == 1
            assert cNvPr_nodes[0].get("hidden") == "1"

            # -- unhide reverts cleanly --
            shape.is_hidden = False
            assert shape.is_hidden is False
            cNvPr_after = shape._element.xpath("./" + parent_tag + "/p:cNvPr")
            assert cNvPr_after[0].get("hidden") is None


# -- helpers --------------------------------------------------------------


def _fresh_slide():
    """Return ``(prs, slide)`` — a new ``Presentation`` with one blank slide."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    return prs, slide


def _roundtrip(prs):
    """Serialize `prs` to a ``BytesIO`` and reopen it.

    Returns the reloaded ``Presentation``.
    """
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)
