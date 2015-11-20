Feature: Insert a table into a placeholder
  In order to add a table to a slide at a pre-defined location and size
  As a developer using python-pptx
  I need a way to insert a table into a placeholder


  Scenario: Insert a table into a table placeholder
     Given an unpopulated table placeholder shape
      When I call placeholder.insert_table(rows=2, cols=3)
      Then the return value is a PlaceholderGraphicFrame object
       And the placeholder contains the table
       And the table has 2 rows and 3 columns
