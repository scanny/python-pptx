Feature: Read and write legacy PowerPoint review comments
  In order to exchange review feedback with PowerPoint users
  As a developer using python-pptx
  I need to read and write the legacy PowerPoint comment parts on a slide

  Scenario: add a comment to a slide and round-trip it
     Given a presentation with one slide
      When I add a comment authored by 'Alice' to the slide
       And I save the presentation
      Then the reloaded slide reports it has a comments part
       And the reloaded slide has one comment with the expected text and author

  Scenario: read an empty comments collection on a slide with no comments
     Given a presentation with one slide
      Then the slide reports it has no comments part
       And the slide's Comments collection has length zero
