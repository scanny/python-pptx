Feature: Round-trip a presentation
  In order to satisfy myself that python-pptx might work
  As a pptx developer
  I want to see it pass a basic sanity-check

  Scenario: Round-trip a basic presentation
     Given a clean working directory
      When I open a basic PowerPoint presentation
       And I save the presentation
      Then I see the pptx file in the working directory

  Scenario: Start presentation from package stream
     Given a clean working directory
      When I open a presentation contained in a stream
       And I save the presentation
      Then I see the pptx file in the working directory

  Scenario: Save presentation to package stream
     Given a clean working directory
      When I open a basic PowerPoint presentation
       And I save the presentation to a stream
       And I save that stream to a file
      Then I see the pptx file in the working directory

  Scenario: Round-trip external relationships
     Given a presentation with external relationships
      When I save and reload the presentation
      Then the external relationships are still there
