Feature: Adjust auto shape
  In order to customize the path of an auto shape to my needs
  As a developer using python-pptx
  I need to set the adjustment values of an auto shape


  Scenario: Set AutoShape adjustment value
     Given a chevron shape
      When I assign 0.15 to shape.adjustments[0]
      Then shape.adjustments[0] is 0.15
