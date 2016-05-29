Feature: Access shapes on a slide layout
  In order to operate on the shapes of a slide layout
  As a developer using python-pptx
  I need access to the shape collection of a slide layout

  Scenario: Access shape collection of a slide layout
     Given a slide layout having three shapes
      Then I can access the shape collection of the slide layout
       And the length of the layout shape collection counts all its shapes

  Scenario: Access shape in layout shape collection
     Given a layout shape collection
      Then I can iterate over the layout shapes
       And I can access a layout shape by index
       And each shape is of the appropriate type
