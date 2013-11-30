Feature: Adjust auto shape
  In order to customize the path of an auto shape to my needs
  As a developer using python-pptx
  I need to set the adjustment values of an auto shape

  Scenario: set shape adjustment value
     Given I have a reference to a chevron shape
      When I set the first adjustment value to 0.15
       And I save the presentation
      Then the chevron shape appears with a less acute arrow head
