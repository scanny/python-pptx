
Working with Presentations
==========================

|pp| allows you to create new presentations as well as make changes to
existing ones. Actually, it only lets you make changes to existing
presentations; it's just that if you start with a presentation that doesn't
have any slides, it feels at first like you're creating one from scratch.

However, a lot of how a presentation looks is determined by the parts that are
left when you delete all the slides, specifically the theme, the slide master,
and the slide layouts that derive from the master. Let's walk through it a step
at a time using examples, starting with the two things you can do with
a presentation, open it and save it.


Opening a presentation
----------------------

The simplest way to get started is to open a new presentation without
specifying a file to open::

    from pptx import Presentation

    prs = Presentation()
    prs.save('test.pptx')

This creates a new presentation from the built-in default template and saves it
unchanged to a file named 'test.pptx'. A couple things to note:

* The so-called "default template" is actually just a PowerPoint file that
  doesn't have any slides in it, stored with the installed |pp| package. It's
  the same as what you would get if you created a new presentation from a fresh
  PowerPoint install, a 4x3 aspect ratio presentation based on the "White"
  template. Well, except it won't contain any slides. PowerPoint always adds
  a blank first slide by default.

* You don't need to do anything to it before you save it. If you want to see
  exactly what that template contains, just look in the 'test.pptx' file this
  creates.

* We've called it a *template*, but in fact it's just a regular PowerPoint file
  with all the slides removed. Actual PowerPoint template files (.potx files)
  are something a bit different. More on those later maybe, but you won't need
  them to work with |pp|.


REALLY opening a presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Okay, so if you want any control at all to speak of over the final
presentation, or if you want to change an existing presentation, you need to
open one with a filename::

    prs = Presentation('existing-prs-file.pptx')
    prs.save('new-file-name.pptx')

Things to note:

* You can open any PowerPoint 2007 or later file this way (.ppt files from
  PowerPoint 2003 and earlier won't work). While you might not be able to
  manipulate all the contents yet, whatever is already in there will load and
  save just fine. The feature set is still being built out, so you can't add or
  change things like Notes Pages yet, but if the presentation has them |pp| is
  polite enough to leave them alone and smart enough to save them without
  actually understanding what they are.

* If you use the same filename to open and save the file, |pp| will obediently
  overwrite the original file without a peep. You'll want to make sure that's
  what you intend.


Opening a 'file-like' presentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

|pp| can open a presentation from a so-called *file-like* object. It can also
save to a file-like object. This can be handy when you want to get the source
or target presentation over a network connection or from a database and don't
want to (or aren't allowed to) fuss with interacting with the file system. In
practice this means you can pass an open file or StringIO/BytesIO stream object
to open or save a presentation like so::

    f = open('foobar.pptx')
    prs = Presentation(f)
    f.close()

    # or

    with open('foobar.pptx') as f:
        source_stream = StringIO(f.read())
    prs = Presentation(source_stream)
    source_stream.close()
    ...
    target_stream = StringIO()
    prs.save(target_stream)


Okay, so you've got a presentation open and are pretty sure you can save it
somewhere later. Next step is to get a slide in there ...
