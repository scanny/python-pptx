Feature: Add a table to a slide
  In order to present tabular content
  As a presentation developer
  I need the ability to place a table on a slide

  Scenario: Add a table to a slide
     Given I have a reference to a blank slide
      When I add a table to the slide's shape collection
       And I save the presentation
      Then the table appears in the slide
