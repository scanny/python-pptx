# __init__.py
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

import inspect
import sys

import pptx.exc as exceptions
sys.modules['pptx.exceptions'] = exceptions

from pptx.presentation import\
    ( Presentation
    )

__all__ = sorted(name for name, obj in locals().items() if not (name.startswith('_') or inspect.ismodule(obj)))
                 
__version__ = '0.0.1'

del inspect, sys

# # Model part constructors might be specified something like this:
# import pptx.presentation
# from pptx.package import PartClassMap
# 
# PartClassMap.register(
#     { 'presentation' : pptx.presentation.Presentation
#     , 'slideMaster'  : pptx.presentation.SlideMaster
#     , 'slideLayout'  : pptx.presentation.SlideLayout
#     , 'slide'        : pptx.presentation.Slide
#     }
# )

# import pptx.presentation
# from pptx.packaging import PartTypeSpec
# 
# PartTypeSpec.register(
#     # {
#     # , 'application/vnd.openxmlformats-officedocument.presentationml'\
#     #   '.comments+xml': pptx.presentation.Comments
#     # , 'application/vnd.openxmlformats-officedocument.presentationml'\
#     #   '.commentAuthors+xml': pptx.presentation.CommentAuthors
#     { 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.handoutMaster+xml': pptx.presentation.HandoutMaster
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.notesMaster+xml': pptx.presentation.NotesMaster
#     # , 'application/vnd.openxmlformats-officedocument.presentationml'\
#     #   '.notesSlide+xml': pptx.presentation.NotesSlide
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.presentation.main+xml': pptx.presentation.Presentation
#     # , 'application/vnd.openxmlformats-officedocument.presentationml'\
#     #   '.template.main+xml': pptx.presentation.Template
#     # , 'application/vnd.openxmlformats-officedocument.presentationml'\
#     #   '.slideshow.main+xml': pptx.presentation.SlideShow
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.presProps+xml': pptx.presentation.PresProps
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.slide+xml': pptx.presentation.Slide
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.slideLayout+xml': pptx.presentation.SlideLayout
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.slideMaster+xml': pptx.presentation.SlideMaster
#     # , 'application/vnd.openxmlformats-officedocument.presentationml'\
#     #   '.tags+xml': pptx.presentation.Tags
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.viewProps+xml': pptx.presentation.ViewProps
#     , 'application/vnd.openxmlformats-officedocument'\
#       '.theme+xml': pptx.presentation.Theme
#     , 'application/vnd.openxmlformats-officedocument.presentationml'\
#       '.tableStyles+xml': pptx.presentation.TableStyles
#     # , 'application/vnd.openxmlformats-package'\
#     #   '.core-properties+xml': pptx.presentation.CoreProps
#     # , 'application/vnd.openxmlformats-officedocument'\
#     #   '.custom-properties+xml': pptx.presentation.CustomProps
#     # , 'application/vnd.openxmlformats-officedocument'\
#     #   '.extended-properties+xml': pptx.presentation.ExtendedProps
#     ,  'image/gif': pptx.presentation.Image
#     ,  'image/jpeg': pptx.presentation.Image
#     ,  'image/png': pptx.presentation.Image
#     # ,  'application/vnd.openxmlformats-officedocument.presentationml'\
#     #    '.printerSettings': pptx.presentation.PrinterSettings
#     }
# )
