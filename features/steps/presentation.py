"""Gherkin step implementations for presentation-level features."""

from __future__ import annotations

import io
import os
import zipfile
from typing import TYPE_CHECKING, cast

from behave import given, then, when
from behave.runner import Context
from helpers import saved_pptx_path, test_file, test_potx, test_pptx

from pptx import Presentation
from pptx.exc import EncryptedPackageError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches

if TYPE_CHECKING:
    from pptx import presentation
    from pptx.shapes.picture import Picture

# given ===================================================


@given("a clean working directory")
def given_clean_working_dir(context: Context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)


@given("a presentation")
def given_a_presentation(context: Context):
    context.presentation = Presentation(test_pptx("prs-properties"))


@given("a presentation having a notes master")
def given_a_presentation_having_a_notes_master(context: Context):
    context.prs = Presentation(test_pptx("prs-notes"))


@given("a presentation having no notes master")
def given_a_presentation_having_no_notes_master(context: Context):
    context.prs = Presentation(test_pptx("prs-properties"))


@given("a presentation with an image/jpg MIME-type")
def given_prs_with_image_jpg_MIME_type(context):
    context.prs = Presentation(test_pptx("test-image-jpg-mime"))


@given("a presentation with external relationships")
def given_prs_with_ext_rels(context: Context):
    context.prs = Presentation(test_pptx("ext-rels"))


@given("a password-protected presentation on disk")
def given_password_protected_presentation_on_disk(context: Context):
    # -- encrypt a small in-memory .pptx and leave the bytes in context for the
    # -- subsequent Then steps; using a stream avoids writing to disk.
    prs = Presentation(test_pptx("test"))
    buf = io.BytesIO()
    prs.save(buf, password="right-password")
    buf.seek(0)
    context.encrypted_pptx_bytes = buf.getvalue()


@given("an initialized pptx environment")
def given_initialized_pptx_env(context: Context):
    pass


# when ====================================================


@when("I change the slide width and height")
def when_change_slide_width_and_height(context: Context):
    presentation = context.presentation
    presentation.slide_width = Inches(4)
    presentation.slide_height = Inches(3)


@when("I construct a Presentation instance with no path argument")
def when_construct_default_prs(context: Context):
    context.prs = Presentation()


@when('I construct a Presentation with pptx_format="{preset}"')
def when_construct_prs_with_pptx_format(context: Context, preset: str):
    context.prs = Presentation(pptx_format=preset)


@when("I open a basic PowerPoint presentation")
def when_open_basic_pptx(context: Context):
    context.prs = Presentation(test_pptx("test"))


@when("I open a presentation extracted into a directory")
def when_I_open_a_presentation_extracted_into_a_directory(context: Context):
    context.prs = Presentation(test_file("extracted-pptx"))


@when("I open a PowerPoint template file")
def when_open_powerpoint_template_file(context: Context):
    context.prs = Presentation(test_potx("minimal"))


@when("I open a presentation contained in a stream")
def when_open_presentation_stream(context: Context):
    with open(test_pptx("test"), "rb") as f:
        stream = io.BytesIO(f.read())
    context.prs = Presentation(stream)
    stream.close()


@when("I save and reload the presentation")
def when_save_and_reload_prs(context: Context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)
    context.prs = Presentation(saved_pptx_path)


@when("I save that stream to a file")
def when_save_stream_to_a_file(context: Context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.stream.seek(0)
    with open(saved_pptx_path, "wb") as f:
        f.write(context.stream.read())


@when("I save the presentation")
def when_save_presentation(context: Context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)


@when("I save the presentation to a stream")
def when_save_presentation_to_stream(context: Context):
    context.stream = io.BytesIO()
    context.prs.save(context.stream)


@when("I save it twice with zip_date_time fixed to 2020-01-01")
def when_save_twice_with_fixed_zip_date_time(context: Context):
    context.stream_a = io.BytesIO()
    context.stream_b = io.BytesIO()
    context.prs.save(context.stream_a, zip_date_time=(2020, 1, 1, 0, 0, 0))
    context.prs.save(context.stream_b, zip_date_time=(2020, 1, 1, 0, 0, 0))


@when("I save a password-protected presentation")
def when_save_password_protected_presentation(context: Context):
    prs = Presentation(test_pptx("test"))
    prs.save(saved_pptx_path, password="pw4behave")


@when("I open the password-protected presentation with the correct password")
def when_open_password_protected_with_correct_pw(context: Context):
    context.prs = Presentation(saved_pptx_path, password="pw4behave")


@when("I save a password-protected presentation with a fixed zip date-time")
def when_save_password_protected_with_fixed_zip_date_time(context: Context):
    prs = Presentation(test_pptx("test"))
    prs.save(
        saved_pptx_path,
        zip_date_time=(2020, 1, 1, 0, 0, 0),
        password="pw4behave",
    )


# then ====================================================


@then("I receive a presentation based on the default template")
def then_receive_prs_based_on_def_tmpl(context: Context):
    prs = context.prs
    assert prs is not None
    slide_masters = prs.slide_masters
    assert slide_masters is not None
    assert len(slide_masters) == 1
    slide_layouts = slide_masters[0].slide_layouts
    assert slide_layouts is not None
    assert len(slide_layouts) == 11


@then("its slide height matches its known value")
def then_slide_height_matches_known_value(context: Context):
    presentation = context.presentation
    assert presentation.slide_height == 6858000


@then("its slide width matches its known value")
def then_slide_width_matches_known_value(context: Context):
    presentation = context.presentation
    assert presentation.slide_width == 9144000


@then("I see the pptx file in the working directory")
def then_see_pptx_file_in_working_dir(context: Context):
    assert os.path.isfile(saved_pptx_path)
    minimum = 30000
    actual = os.path.getsize(saved_pptx_path)
    assert actual > minimum


@then("the presentation is loaded")
def then_the_presentation_is_loaded(context: Context):
    prs = context.prs
    assert prs is not None
    # -- a loaded .potx still yields a working Presentation graph --
    assert len(prs.slide_masters) >= 1


@then("I see the pptx file in the working directory after saving")
def then_see_pptx_after_saving(context: Context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)
    assert os.path.isfile(saved_pptx_path)
    # -- a .potx round-tripped to a saved file is non-trivial in size --
    assert os.path.getsize(saved_pptx_path) > 1000


@then("len(notes_master.shapes) is {shape_count}")
def then_len_notes_master_shapes_is_shape_count(context: Context, shape_count: str):
    notes_master = context.prs.notes_master
    expected = int(shape_count)
    actual = len(notes_master.shapes)
    assert actual == expected, "got %s" % actual


@then("prs.notes_master is a NotesMaster object")
def then_prs_notes_master_is_a_NotesMaster_object(context: Context):
    prs = context.prs
    assert type(prs.notes_master).__name__ == "NotesMaster"


@then("prs.slides is a Slides object")
def then_prs_slides_is_a_Slides_object(context: Context):
    prs = context.presentation
    assert type(prs.slides).__name__ == "Slides"


@then("prs.slide_masters is a SlideMasters object")
def then_prs_slide_masters_is_a_SlideMasters_object(context: Context):
    prs = context.presentation
    assert type(prs.slide_masters).__name__ == "SlideMasters"


@then("the external relationships are still there")
def then_ext_rels_are_preserved(context: Context):
    prs = context.prs
    sld = prs.slides[0]
    rel = sld.part._rels["rId2"]
    assert rel.is_external
    assert rel.reltype == RT.HYPERLINK
    assert rel.target_ref == "https://github.com/scanny/python-pptx"


@then("the package has the expected number of .rels parts")
def then_the_package_has_the_expected_number_of_rels_parts(context: Context):
    with zipfile.ZipFile(saved_pptx_path, "r") as z:
        member_count = len(z.namelist())
    assert member_count == 18, "expected 18, got %d" % member_count


@then("I can access the JPEG image")
def then_I_can_access_the_JPEG_image(context):
    prs = cast("presentation.Presentation", context.prs)
    slide = prs.slides[0]
    picture = cast("Picture", slide.shapes[0])
    try:
        picture.image
    except AttributeError:
        raise AssertionError("JPEG image not recognized")


@then("the slide height matches the new value")
def then_slide_height_matches_new_value(context: Context):
    presentation = context.presentation
    assert presentation.slide_height == Inches(3)


@then("the slide width matches the new value")
def then_slide_width_matches_new_value(context: Context):
    presentation = context.presentation
    assert presentation.slide_width == Inches(4)


@then("the new presentation has slide width {cx:d} and height {cy:d}")
def then_new_prs_has_slide_width_and_height(context: Context, cx: int, cy: int):
    prs = context.prs
    assert prs.slide_width == cx, "expected %d, got %s" % (cx, prs.slide_width)
    assert prs.slide_height == cy, "expected %d, got %s" % (cy, prs.slide_height)


@then("len(prs.slide_masters) is {count:d}")
def then_len_prs_slide_masters_is_count(context: Context, count: int):
    actual = len(context.prs.slide_masters)
    assert actual == count, "expected %d slide masters, got %d" % (count, actual)


@then("len(prs.slide_layouts) is {count:d}")
def then_len_prs_slide_layouts_is_count(context: Context, count: int):
    actual = len(context.prs.slide_layouts)
    assert actual == count, "expected %d slide layouts, got %d" % (count, actual)


@then("both saved streams are byte-for-byte identical")
def then_both_streams_byte_identical(context: Context):
    a = context.stream_a.getvalue()
    b = context.stream_b.getvalue()
    assert a == b, "streams differ (lengths %d vs %d)" % (len(a), len(b))


@then("every zip member carries the 2020-01-01 00:00:00 last-modified stamp")
def then_every_zip_member_has_fixed_date(context: Context):
    context.stream_a.seek(0)
    with zipfile.ZipFile(context.stream_a, "r") as z:
        for info in z.infolist():
            assert info.date_time == (2020, 1, 1, 0, 0, 0), (
                "member %r has date_time %r" % (info.filename, info.date_time)
            )


@then("the saved .pptx starts with the OLE2 magic signature")
def then_saved_pptx_starts_with_ole2_magic(context: Context):
    with open(saved_pptx_path, "rb") as f:
        header = f.read(8)
    assert header == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", (
        "expected OLE2 magic for an encrypted .pptx, got %r" % header
    )


@then("opening it without a password raises EncryptedPackageError")
def then_opening_without_password_raises(context: Context):
    try:
        Presentation(io.BytesIO(context.encrypted_pptx_bytes))
    except EncryptedPackageError:
        return
    raise AssertionError("expected EncryptedPackageError was not raised")


@then("opening it with the wrong password raises EncryptedPackageError")
def then_opening_with_wrong_password_raises(context: Context):
    try:
        Presentation(
            io.BytesIO(context.encrypted_pptx_bytes), password="definitely-wrong"
        )
    except EncryptedPackageError:
        return
    raise AssertionError("expected EncryptedPackageError was not raised")


@then("every decrypted zip member carries the 2020-01-01 00:00:00 last-modified stamp")
def then_every_decrypted_zip_member_has_fixed_date(context: Context):
    # -- decrypt the on-disk password-protected file and verify every member
    # -- of the resulting (inner plaintext) zip carries the fixed stamp.
    from pptx.opc._crypto import decrypt_stream

    with open(saved_pptx_path, "rb") as f:
        plain_bytes = decrypt_stream(f, "pw4behave")
    with zipfile.ZipFile(io.BytesIO(plain_bytes), "r") as z:
        for info in z.infolist():
            assert info.date_time == (2020, 1, 1, 0, 0, 0), (
                "member %r has date_time %r" % (info.filename, info.date_time)
            )
