Feature: Round-trip a presentation
  In order to satisfy myself that python-pptx might work
  As a pptx developer
  I want to see it pass a basic sanity-check

  Scenario: Round-trip a basic presentation
     Given a clean working directory
      When I open a basic PowerPoint presentation
       And I save the presentation
      Then I see the pptx file in the working directory
