Feature: Change properties of table
  In order to change the formatting of a table to meet my needs
  As a developer using python-pptx
  I need to set the properties of a table

  Scenario: Set cell first_col property
     Given I have a reference to a table
      When I set the first_col property to True
       And I save the presentation
      Then the first column of the table has special formatting

  Scenario: Set cell last_row property
     Given I have a reference to a table
      When I set the last_row property to True
       And I save the presentation
      Then the last row of the table has special formatting
