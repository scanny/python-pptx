Feature: Get and set AutoShape properties
  In order to identify and adjust AutoShapes
  As a developer using python-pptx
  I need a set of properties on Shape


  Scenario: Get AutoShape text
     Given an autoshape having text
      Then shape.text is the text in the shape


  Scenario: Set AutoShape text
     Given an autoshape having text
      When I assign a string to shape.text
      Then shape.text is the string I assigned
