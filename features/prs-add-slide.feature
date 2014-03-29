Feature: Add a new slide
  In order to extend a presentation
  As a python developer using python-pptx
  I need to add a new slide to a presentation

  Scenario: Add a new slide to a presentation
     Given an empty presentation
      When I add a slide based on a layout
       And I save the presentation
      Then the pptx file contains a single slide
       And the layout has been applied to the slide

  Scenario: DELETEME - Temporary to drive clone_layout_placeholder
     Given an empty slide shape collection
      When I clone layout placeholders
      Then corresponding slide placeholders are added to the collection
