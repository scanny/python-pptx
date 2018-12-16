Feature: Access an individual slide
  In order to interact with an individual slide
  As a developer using python-pptx
  I need ways to create, access, or delete a slide


  Scenario: Slides.__getitem__()
    Given a Slides object containing 3 slides
     Then slides[2] is a Slide object


  Scenario: Slides.__iter__()
    Given a Slides object containing 3 slides
     Then iterating slides produces 3 Slide objects


  Scenario: Slides.__len__()
    Given a Slides object containing 3 slides
     Then len(slides) is 3


  Scenario: Slides.add_slide()
    Given a Slides object containing 3 slides
     When I call slides.add_slide()
     Then len(slides) is 4
      And slide.slide_layout is the one passed in the call


  Scenario: Slides.get()
    Given a Slides object containing 3 slides
     Then slides.get(256) is slides[0]
      And slides.get(666, default=slides[2]) is slides[2]


  Scenario: SlideLayouts.__getitem__()
    Given a SlideLayouts object containing 2 layouts as slide_layouts
     Then slide_layouts[1] is a SlideLayout object


  Scenario: SlideLayouts.__iter__()
    Given a SlideLayouts object containing 2 layouts as slide_layouts
     Then iterating slide_layouts produces 2 SlideLayout objects


  Scenario: SlideLayouts.__len__()
    Given a SlideLayouts object containing 2 layouts as slide_layouts
     Then len(slide_layouts) is 2


  Scenario: SlideLayouts.get_by_name()
    Given a SlideLayouts object containing 2 layouts as slide_layouts
     Then slide_layouts.get_by_name(slide_layouts[1].name) is slide_layouts[1]


  Scenario: SlideLayouts.index()
    Given a SlideLayouts object containing 2 layouts as slide_layouts
     Then slide_layouts.index(slide_layouts[1]) == 1


  Scenario: SlideLayouts.remove()
    Given a SlideLayouts object containing 2 layouts as slide_layouts
     When I call slide_layouts.remove(slide_layouts[1])
     Then len(slide_layouts) is 1


  Scenario: SlideMasters.__getitem__()
    Given a SlideMasters object containing 2 masters
     Then slide_masters[1] is a SlideMaster object


  Scenario: SlideMasters.__iter__()
    Given a SlideMasters object containing 2 masters
     Then iterating slide_masters produces 2 SlideMaster objects


  Scenario: SlideMasters.__len__()
    Given a SlideMasters object containing 2 masters
     Then len(slide_masters) is 2
