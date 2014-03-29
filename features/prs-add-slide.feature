Feature: Add a new slide
  In order to extend a presentation
  As a python developer using python-pptx
  I need to add a new slide to a presentation

  Scenario: Add a new slide to a presentation
     Given I have an empty presentation open
      When I add a new slide
       And I save the presentation
      Then the pptx file contains a single slide
       And the layout was applied to the slide
