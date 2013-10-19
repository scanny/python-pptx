# encoding: utf-8

"""
Helper methods and variables for acceptance tests.
"""

import os


def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))

thisdir = os.path.split(__file__)[0]
scratch_dir = absjoin(thisdir, '../_scratch')
test_file_dir = absjoin(thisdir, '../../tests/test_files')
basic_pptx_path = absjoin(test_file_dir, 'test.pptx')
no_core_props_pptx_path = absjoin(test_file_dir, 'no-core-props.pptx')
saved_pptx_path = absjoin(scratch_dir, 'test_out.pptx')
test_image_path = absjoin(test_file_dir, 'python-powered.png')

test_text = "python-pptx was here!"
