Feature: Change properties of text in shapes
  In order to change the formatting of text to my needs
  As a developer using python-pptx
  I need to set the properties of text in a shape

  Scenario: Set paragraph alignment
     Given I have a reference to a paragraph
      When I set the paragraph alignment to centered
       And I save the presentation
      Then the paragraph is aligned centered


