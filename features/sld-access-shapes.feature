Feature: Access shapes on a slide
  In order to operate on a shape in a slide
  As a developer using python-pptx
  I need a shape collection to provide access to shapes on a slide

  Scenario: Access shape collection
     Given a slide having three shapes
      Then I can access the shape collection of the slide
       And the length of the shape collection is 3

  Scenario: Access shape in shape collection
     Given a slide shape collection
      Then I can iterate over the shapes
       And I can access a shape by index
       And each slide shape is of the appropriate type

  @wip
  Scenario: Get index of shape in shape collection sequence
     Given a slide shape collection
      Then the index of each shape matches its position in the sequence
