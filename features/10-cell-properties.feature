Feature: Change properties of table cell
  In order to change the formatting of a cell in a table to my needs
  As a developer using python-pptx
  I need to set the properties of a table cell

  Scenario: Set cell vertical anchor
     Given I have a reference to a table cell
      When I set the cell vertical anchor to middle
       And I save the presentation
      Then the cell contents are vertically centered

  Scenario: Set cell margins
     Given I have a reference to a table cell
      When I set the cell margins
       And I save the presentation
      Then the cell contents are inset by the margins
