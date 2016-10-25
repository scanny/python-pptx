Feature: Placeholder inherits properties
  In order to conveniently access effective placeholder property values
  As a developer using python-pptx
  I need a placeholder to provide property values that reflect inheritance

  Scenario: Master placeholder property values
     Given a master placeholder
      Then I can get the placeholder position
       And I can get the placeholder dimensions

  Scenario: Layout placeholder direct property settings
     Given a layout placeholder having directly set position and size
      Then I get the direct settings when I query position and size

  Scenario: Layout placeholder inherited property settings
     Given a layout placeholder having no direct position or size settings
      Then I get inherited settings when I query position and size

  Scenario: NotesSlide placeholder direct property settings
     Given a notes slide placeholder having directly set position and size
      Then I get the direct settings when I query position and size

  Scenario: NotesSlide placeholder inherited property settings
     Given a notes slide placeholder having no direct position or size settings
      Then I get inherited settings when I query position and size

  Scenario: Slide placeholder direct property settings
     Given a slide placeholder having directly set position and size
      Then I get the direct settings when I query position and size

  Scenario: Slide placeholder inherited property settings
     Given a slide placeholder having no direct position or size settings
      Then I get inherited settings when I query position and size
