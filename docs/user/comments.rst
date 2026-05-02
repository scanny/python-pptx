Working with Review Comments
============================

PowerPoint users can attach review comments to individual slides. |pp|
supports reading and writing the *legacy* comment format defined by
ECMA-376 Part 1 (the ``p:cmLst`` / ``p:cmAuthorLst`` schema). This is the
format used by PowerPoint up through 2018 and is still round-tripped
correctly by current versions.

.. note::

    PowerPoint 2018 and later also emit a newer "modern comments" format
    that lives in a Microsoft extension namespace. That format is **not**
    supported yet — reading or writing it is a follow-up item. Files that
    contain only legacy comments round-trip faithfully.


Reading comments
----------------

Every slide has a :attr:`~pptx.slide.Slide.comments` property that returns a
:class:`~pptx.comments.Comments` collection:

.. code-block:: python

    from pptx import Presentation

    prs = Presentation("reviewed.pptx")
    slide = prs.slides[0]
    if slide.has_comments:
        for comment in slide.comments:
            author = comment.author.name if comment.author else "<unknown>"
            print(f"{author}: {comment.text}")

:attr:`~pptx.slide.Slide.has_comments` returns ``True`` only if a comments
part already exists on the slide — accessing
:attr:`~pptx.slide.Slide.comments` itself creates one lazily as a side effect
of adding an author or a comment. Use ``has_comments`` if you want to check
without triggering creation.


Writing comments
----------------

Adding a comment requires a :class:`~pptx.comments.CommentAuthor` reference,
so the underlying XML can link the comment back to the package-level author
registry:

.. code-block:: python

    import datetime

    from pptx import Presentation

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # -- Author registry lives on the Comments collection (package-scoped).
    authors = slide.comments._authors
    alice = authors.get_or_add("Alice", "A.")

    slide.comments.add_comment(
        alice,
        "Can we tighten the wording here?",
        position=(914400, 457200),           # -- EMU; (x, y) on the slide
        datetime_value=datetime.datetime.utcnow(),
    )

    prs.save("reviewed.pptx")

The same author registry is shared across every slide in the presentation.
Calling :meth:`~pptx.comments.CommentAuthors.get_or_add` reuses an existing
author whose ``name`` and ``initials`` match.


Comment attributes
------------------

Each :class:`~pptx.comments.Comment` exposes:

- ``text`` — the body text of the comment
- ``author`` / ``author_id`` — the originating author (or ``None`` if the
  author id no longer resolves)
- ``position`` — ``(x, y)`` anchor on the slide in English Metric Units
- ``datetime`` — creation timestamp, or ``None`` if the file didn't record
  one
- ``idx`` — per-author sequence index, assigned automatically

All of these properties are read/write on newly added comments.
