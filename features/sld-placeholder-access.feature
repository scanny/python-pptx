Feature: Access slide placeholders
  In order to operate on the placeholder shapes of a slide
  As a developer using python-pptx
  I need a placeholder collection on the slide

  Scenario: Access placeholder collection of a slide
     Given a slide having two placeholders
      Then I can access the placeholder collection of the slide
       And the length of the slide placeholder collection is 2

  Scenario: Access placeholder in slide placeholder collection
     Given a slide placeholder collection
      Then I can iterate over the slide placeholders
       And I can access a slide placeholder by index
