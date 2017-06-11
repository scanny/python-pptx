Feature: Get and change shape line properties
  In order to format the outline of an auto shape
  As a developer using python-pptx
  I need access to the line properties of the shape


  Scenario: Access AutoShape line format
     Given an autoshape
      Then shape.line is a LineFormat object


  Scenario: Access Picture line format
     Given a picture
      Then shape.line is a LineFormat object
