Feature: Default presentation template is provided
  In order to get started on a presentation with a minimum of preparation
  As a developer using python-pptx
  I would like the option to start from a default template

  Scenario: Create a minimal presentation from the default template
     Given an initialized pptx environment
      When I construct a Presentation instance with no path argument
      Then I receive a presentation based on the default template
