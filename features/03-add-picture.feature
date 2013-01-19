Feature: Add a picture to a slide
  In order to generate a presentation that includes images
  As a python developer using python-pptx
  I want to place an image on a slide

  Scenario: Add a picture to a slide
     Given I have a reference to a slide
      When I add a picture to the slide's shape collection
       And I save the presentation
      Then the image is saved in the pptx file
       And the picture appears in the slide




# Feature: Create a presentation from the default template
