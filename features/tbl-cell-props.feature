Feature: Change properties of table cell
  In order to change the formatting of a cell in a table to my needs
  As a developer using python-pptx
  I need to set the properties of a table cell

  Scenario: Set cell text
     Given a 2x2 table
      When I set the text of the first cell
       And I save the presentation
      Then the text appears in the first cell of the table


  Scenario: set cell vertical anchor
     Given a table cell
      When I set the cell vertical anchor to middle
       And I save the presentation
      Then the cell contents are vertically centered


  Scenario: set cell margins
     Given a table cell
      When I set the cell margins
       And I save the presentation
      Then the cell contents are inset by the margins


  Scenario: set cell fill
     Given a table cell
      When I set the cell fill type to solid
       And I set the cell fill foreground color to an RGB value
      Then the foreground color of the cell is the RGB value I set
