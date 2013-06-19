Feature: Read and write core document properties
  In order to find documents and make them manageable by digital means
  As a developer using python-pptx
  I need to access and assign Dublin Core metadata for a presentation

  Scenario: round-trip Dublin Core metadata describing a presentation
     Given I have a reference to the core properties of a presentation
      When I set the core properties to valid values
       And I save the presentation
      Then the core properties of the presentation have the values I set
