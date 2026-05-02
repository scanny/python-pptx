.. _comments_api:

Comments
========

Python-pptx supports the legacy PowerPoint review-comments feature (ECMA-376
Part 1 ``p:cmLst`` / ``p:cmAuthorLst``). Each slide can carry zero or more
comments, each of which references an author in the package-level author
registry.

The 2018 "modern comments" format (which uses a Microsoft extension namespace
and separate package parts) is not yet supported.


|Comments| objects
------------------

A |Comments| instance is obtained via the
:attr:`~pptx.slide.Slide.comments` property of a |Slide|:

.. code-block:: python

    >>> from pptx import Presentation
    >>> prs = Presentation()
    >>> slide = prs.slides.add_slide(prs.slide_layouts[5])
    >>> comments = slide.comments

On first access the package-level comment-authors part and the slide-level
comments part are created automatically if they do not yet exist.

.. autoclass:: pptx.comments.Comments()
   :members:


|Comment| objects
-----------------

An individual |Comment| is obtained by indexed access, iteration, or as the
return value of :meth:`Comments.add_comment`.

.. autoclass:: pptx.comments.Comment()
   :members:


|CommentAuthors| objects
------------------------

The |CommentAuthors| collection contains every comment-author registered at
the package level. Add a new author via
:meth:`CommentAuthors.add_author` or reuse one via
:meth:`CommentAuthors.get_or_add`.

.. autoclass:: pptx.comments.CommentAuthors()
   :members:


|CommentAuthor| objects
-----------------------

.. autoclass:: pptx.comments.CommentAuthor()
   :members:
