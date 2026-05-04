from __future__ import annotations

import tempfile
from pathlib import Path

from behave import given, then, when
from pptx import Presentation

REAL_FIXTURE = Path("/mnt/data/Temp/microsoft/sample.pptx")


@given("a PowerPoint-authored pptx fixture")
def given_fixture(context):
    if not REAL_FIXTURE.exists():
        context.scenario.skip("fixture not available at " + str(REAL_FIXTURE))
        return
    context.original_path = REAL_FIXTURE
    context.tmpdir = Path(tempfile.mkdtemp())
    context.roundtrip_path = context.tmpdir / "rt.pptx"


@when("I round-trip it through python-pptx")
def when_roundtrip(context):
    Presentation(str(context.original_path)).save(str(context.roundtrip_path))


@then("ooxml-validate reports zero new issues")
def then_zero_new_issues(context):
    try:
        from ooxml_validate import OoxmlValidateToolNotFound, validate_roundtrip
    except ImportError:
        context.scenario.skip("ooxml-validate not installed")
        return
    try:
        new_issues = validate_roundtrip(context.original_path, context.roundtrip_path)
    except OoxmlValidateToolNotFound as e:
        context.scenario.skip("ooxml-validate tool not available: " + str(e).splitlines()[0])
        return
    assert len(new_issues) == 0, f"{len(new_issues)} new issues: " + ", ".join(
        f"{i.rule_id}" for i in new_issues[:5]
    )


@then("the LibreOffice PDF render is visually identical")
def then_visually_identical(context):
    try:
        from ooxml_validate import libreoffice_pdf_diff
    except ImportError:
        context.scenario.skip("ooxml-validate not installed")
        return
    try:
        r = libreoffice_pdf_diff(context.original_path, context.roundtrip_path)
    except FileNotFoundError as e:
        context.scenario.skip("LibreOffice not available: " + str(e).splitlines()[0])
        return
    assert r.ok, f"diff failed: pages_differing={r.pages_differing} max_diff={r.max_difference}"
