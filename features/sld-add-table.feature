Feature: Add a table to a slide
  In order to present tabular content
  As a presentation developer
  I need the ability to place a table on a slide

  Scenario: Add a table to a slide
     Given a blank slide
      When I add a table to the slide's shape collection
       And I save the presentation
      Then the table appears in the slide

  Scenario: Set column widths
     Given a 2x2 table
      When I set the width of the table's columns
       And I save the presentation
      Then the table appears with the new column widths

  Scenario: Set cell text
     Given a 2x2 table
      When I set the text of the first cell
       And I save the presentation
      Then the text appears in the first cell of the table
