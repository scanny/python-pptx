Feature: Add a shape to a slide
  In order to accommodate a requirement for geometric shapes on a slide
  As a presentation developer
  I need the ability to place a shape on a slide

  Scenario: Add an auto shape to a slide
     Given I have a reference to a blank slide
      When I add an auto shape to the slide's shape collection
       And I save the presentation
      Then the auto shape appears in the slide

