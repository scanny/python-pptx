Feature: Access an individual slide
  In order to interact with an individual slide
  As a developer using python-pptx
  I need ways to create, access, or delete a slide


  Scenario: Slides.__len__()
    Given a Slides object containing 3 slides
     Then len(slides) is 3


  Scenario: Slides indexed access
    Given a Slides object containing 3 slides
     Then slides[2] is a Slide object


  Scenario: Iterate Slides
    Given a Slides object containing 3 slides
     Then iterating slides produces 3 Slide objects


  Scenario: Slides.add_slide()
     Given a Slides object containing 3 slides
      When I call slides.add_slide()
      Then len(slides) is 4
       And slide.slide_layout is the one passed in the call
