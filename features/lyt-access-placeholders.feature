Feature: Access layout placeholders
  In order to operate on the placeholder shapes of a slide layout
  As a developer using python-pptx
  I need a placeholder collection on the slide layout

  Scenario: Access placeholder collection of a slide layout
     Given a slide layout having two placeholders
      Then I can access the placeholder collection of the slide layout
       And the length of the layout placeholder collection is 2

  Scenario: Access placeholder in layout placeholder collection
     Given a layout placeholder collection
      Then I can iterate over the layout placeholders
       And I can access a layout placeholder by index
       And I can access a layout placeholder by idx value
