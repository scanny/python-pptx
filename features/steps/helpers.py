# encoding: utf-8

"""
Helper methods and variables for acceptance tests.
"""

import os


def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
scratch_dir = absjoin(thisdir, '../_scratch')
# legacy acceptance test pptx files are in unit test file dir
test_file_dir = absjoin(thisdir, '../../tests/test_files')
# new ones should go here instead, others should be moved over
test_pptx_dir = absjoin(thisdir, 'test_files')

# new test pptx files ------------------
shp_pos_and_size_pptx_path = absjoin(test_pptx_dir, 'shp-pos-and-size.pptx')

# legacy test pptx files ---------------
italics_pptx_path = absjoin(test_file_dir, 'italic-runs.pptx')
no_core_props_pptx_path = absjoin(test_file_dir, 'no-core-props.pptx')
test_image_path = absjoin(test_file_dir, 'python-powered.png')

# scratch test pptx file ---------------
saved_pptx_path = absjoin(scratch_dir, 'test_out.pptx')

test_text = "python-pptx was here!"


def test_pptx(name):
    """
    Return the absolute path to test .pptx file with root name *name*.
    """
    return absjoin(thisdir, 'test_files', '%s.pptx' % name)
