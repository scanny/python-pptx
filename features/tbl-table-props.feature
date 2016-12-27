Feature: Change properties of table
  In order to change the formatting of a table to meet my needs
  As a developer using python-pptx
  I need to set the properties of a table


  Scenario: Set column widths
     Given a 2x2 table
      When I set the width of the table's columns
       And I save the presentation
      Then the table appears with the new column widths


  Scenario: Set table first_row property
     Given a 2x2 table
      When I set the first_row property to True
       And I save the presentation
      Then the first row of the table has special formatting


  Scenario: Set table first_col property
     Given a 2x2 table
      When I set the first_col property to True
       And I save the presentation
      Then the first column of the table has special formatting


  Scenario: Set table last_row property
     Given a 2x2 table
      When I set the last_row property to True
       And I save the presentation
      Then the last row of the table has special formatting


  Scenario: Set table last_col property
     Given a 2x2 table
      When I set the last_col property to True
       And I save the presentation
      Then the last column of the table has special formatting


  Scenario: Set table horz_banding property
     Given a 2x2 table
      When I set the horz_banding property to True
       And I save the presentation
      Then the rows of the table have alternating shading


  Scenario: Set table vert_banding property
     Given a 2x2 table
      When I set the vert_banding property to True
       And I save the presentation
      Then the columns of the table have alternating shading
