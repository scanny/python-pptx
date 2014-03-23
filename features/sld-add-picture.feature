Feature: Add a picture to a slide
  In order to produce a presentation that includes images
  As a python developer using python-pptx
  I need a way to place an image on a slide

  @wip
  Scenario: Add a picture to a slide
     Given a slide
      When I add a picture to the slide's shape collection
       And I save the presentation
      Then the image is saved in the pptx file
       And the picture appears in the slide

  @wip
  Scenario: Add a picture stream to a slide
     Given a slide
      When I add a picture stream to the slide's shape collection
       And I save the presentation
      Then the image is saved in the pptx file
       And the picture appears in the slide
