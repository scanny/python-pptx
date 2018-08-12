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


  Scenario: Set AutoShape adjustment value
     Given a chevron shape
      When I assign 0.15 to shape.adjustments[0]
      Then shape.adjustments[0] is 0.15


  Scenario: Shape.line
     Given a Shape object as shape
      Then shape.line is a LineFormat object
