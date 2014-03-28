Feature: Add a text box to a slide
  In order to have free-form text on a slide
  As a presentation developer
  I need the ability to place a text box on a slide

  Scenario: Add a text box to a slide
     Given a blank slide
      When I add a text box to the slide's shape collection
       And I save the presentation
      Then the text box appears in the slide
