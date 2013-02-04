import logging
import os

from behave   import given, when, then
from hamcrest import (assert_that, has_item, is_, is_not, equal_to,
                      greater_than)

from pptx import packaging
from pptx import Presentation
from pptx.util import Inches

def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir       = os.path.split(__file__)[0]
scratch_dir   = absjoin(thisdir, '../_scratch')
test_file_dir = absjoin(thisdir, '../../test/test_files')
basic_pptx_path = absjoin(test_file_dir, 'test.pptx')
saved_pptx_path = absjoin(scratch_dir,   'test_out.pptx')
test_image_path = absjoin(test_file_dir, 'python-powered.png')

test_text = "python-pptx was here!"

# logging.debug("saved_pptx_path is ==> '%s'\n", saved_pptx_path)

# given ---------------------------------------------------

@given('a clean working directory')
def step(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)


@given('an initialized pptx environment')
def step(context):
    pass


@given('I have an empty presentation open')
def step(context):
    context.prs = Presentation()


@given('I have a reference to a blank slide')
def step(context):
    context.prs = Presentation()
    slidelayout = context.prs.slidelayouts[6]
    context.sld = context.prs.slides.add_slide(slidelayout)


@given('I have a reference to a slide')
def step(context):
    context.prs = Presentation()
    slidelayout = context.prs.slidelayouts[0]
    context.sld = context.prs.slides.add_slide(slidelayout)


# when ----------------------------------------------------

@when('I add a new slide')
def step(context):
    slidelayout = context.prs.slidemasters[0].slidelayouts[0]
    context.prs.slides.add_slide(slidelayout)


@when("I add a picture to the slide's shape collection")
def step(context):
    shapes = context.sld.shapes
    x, y = (Inches(1.25), Inches(1.25))
    shapes.add_picture(test_image_path, x, y)


@when("I add a text box to the slide's shape collection")
def step(context):
    shapes = context.sld.shapes
    x, y = (Inches(1.00), Inches(2.00))
    cx, cy = (Inches(3.00), Inches(1.00))
    sp = shapes.add_textbox(x, y, cx, cy)
    sp.text = test_text


@when('I construct a Presentation instance with no path argument')
def step(context):
    context.prs = Presentation()


@when('I open a basic PowerPoint presentation')
def step(context):
    context.prs = Presentation(basic_pptx_path)


@when('I save the presentation')
def step(context):
    if os.path.isfile(saved_pptx_path):
        os.remove(saved_pptx_path)
    context.prs.save(saved_pptx_path)


@when("I set the title text of the slide")
def step(context):
    context.sld.shapes.title.text = test_text


# then ----------------------------------------------------

@then('I receive a presentation based on the default template')
def step(context):
    prs = context.prs
    assert_that(prs, is_not(None))
    slidemasters = prs.slidemasters
    assert_that(slidemasters, is_not(None))
    assert_that(len(slidemasters), is_(1))
    slidelayouts = slidemasters[0].slidelayouts
    assert_that(slidelayouts, is_not(None))
    assert_that(len(slidelayouts), is_(11))


@then('I see the pptx file in the working directory')
def step(context):
    assert_that(os.path.isfile(saved_pptx_path))
    minimum = 30000
    actual = os.path.getsize(saved_pptx_path)
    assert_that(actual, is_(greater_than(minimum)))


@then('the image is saved in the pptx file')
def step(context):
    pkgng_pkg = packaging.Package().open(saved_pptx_path)
    partnames = [part.partname for part in pkgng_pkg.parts
                 if part.partname.startswith('/ppt/media/')]
    assert_that(partnames, has_item('/ppt/media/image1.png'))


@then('the picture appears in the slide')
def step(context):
    prs = Presentation(saved_pptx_path)
    sld = prs.slides[0]
    shapes = sld.shapes
    classnames = [sp.__class__.__name__ for sp in shapes]
    assert_that(classnames, has_item('Picture'))


@then('the text box appears in the slide')
def step(context):
    prs = Presentation(saved_pptx_path)
    textbox = prs.slides[0].shapes[0]
    textbox_text = textbox.textframe.paragraphs[0].runs[0].text
    assert_that(textbox_text, is_(equal_to(test_text)))


@then('the pptx file contains a single slide')
def step(context):
    prs = Presentation(saved_pptx_path)
    assert_that(len(prs.slides), is_(equal_to(1)))


@then('the text appears in the title placeholder')
def step(context):
    prs = Presentation(saved_pptx_path)
    title_shape = prs.slides[0].shapes.title
    title_text = title_shape.textframe.paragraphs[0].runs[0].text
    assert_that(title_text, is_(equal_to(test_text)))

