Feature: Keep extended document properties in sync on save
  In order to have presentations preview correctly in Gmail and other downstream tools
  As a developer using python-pptx
  I need the extended-properties part (docProps/app.xml) to reflect the real slide count

  Scenario: slide count in app.xml is updated on save
     Given I have a presentation with three slides added
      When I save the presentation
      Then the <Slides> element in docProps/app.xml equals 3
