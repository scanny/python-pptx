import logging
import os

from behave   import given, when, then
from hamcrest import (assert_that, has_item, is_, is_not, equal_to,
                      greater_than)

from pptx import packaging
from pptx.presentation import Package
from pptx.util import Inches

def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir       = os.path.split(__file__)[0]
scratch_dir   = absjoin(thisdir, '../_scratch')
test_file_dir = absjoin(thisdir, '../../test/test_files')
basic_pptx_path = absjoin(test_file_dir, 'test.pptx')
empty_pptx_path = absjoin(test_file_dir, 'no-slides.pptx')
saved_pptx_path = absjoin(scratch_dir,   'test_out.pptx')
test_image_path = absjoin(test_file_dir, 'python-powered.png')

# logging.debug("saved_pptx_path is ==> '%s'\n", saved_pptx_path)

@given('a clean working directory')
def step(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)


@given('I have an empty presentation open')
def step(context):
    context.pkg = Package().open(empty_pptx_path)
    context.prs = context.pkg.presentation


@given('I have a reference to a slide')
def step(context):
    context.pkg = Package().open(empty_pptx_path)
    context.prs = context.pkg.presentation
    slidelayout = context.prs.slidemasters[0].slidelayouts[0]
    context.sld = context.prs.slides.add_slide(slidelayout)


@when('I open a basic PowerPoint presentation')
def step(context):
    context.pkg = Package().open(basic_pptx_path)


@when('I add a new slide')
def step(context):
    slidelayout = context.prs.slidemasters[0].slidelayouts[0]
    context.prs.slides.add_slide(slidelayout)


@when("I add a picture to the slide's shape collection")
def step(context):
    shapes = context.sld.shapes
    x, y = (Inches(1.25), Inches(1.25))
    shapes.add_picture(test_image_path, x, y)


@when('I save the presentation')
def step(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    packaging.Package().marshal(context.pkg).save(saved_pptx_path)


@then('I see the pptx file in the working directory')
def step(context):
    assert_that(os.path.isfile(saved_pptx_path))
    minimum = 30000
    actual = os.path.getsize(saved_pptx_path)
    assert_that(actual, is_(greater_than(minimum)))


@then('the pptx file contains a single slide')
def step(context):
    prs = Package().open(saved_pptx_path).presentation
    assert_that(len(prs.slides), is_(equal_to(1)))


@then('the image is saved in the pptx file')
def step(context):
    pkgng_pkg = packaging.Package().open(saved_pptx_path)
    partnames = [part.partname for part in pkgng_pkg.parts
                 if part.partname.startswith('/ppt/media/')]
    assert_that(partnames, has_item('/ppt/media/image1.png'))


@then('the picture appears in the slide')
def step(context):
    prs = Package().open(saved_pptx_path).presentation
    sld = prs.slides[0]
    shapes = sld.shapes
    classnames = [sp.__class__.__name__ for sp in shapes]
    assert_that(classnames, has_item('Picture'))


