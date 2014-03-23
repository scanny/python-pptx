Feature: Add a shape to a slide
  In order to produce a slide that features geometric shapes
  As a presentation developer
  I need a way to place a shape on a slide

  Scenario: Add an auto shape to a slide
     Given a blank slide
      When I add an auto shape to the slide's shape collection
       And I save the presentation
      Then the auto shape appears in the slide

