Feature: Organize a presentation into sections
  In order to group related slides into named, navigable sections
  As a developer using python-pptx
  I need to create, rename, and delete presentation sections, and assign
  slides to them


  Scenario: Add a section and assign slides to it
    Given a presentation with three slides
     When I add a section named "Intro" with slides 1 and 2
     Then Presentation.sections has 1 section
      And the section is named "Intro"
      And the section has two assigned slides
      And the presentation.xml has a p14:sectionLst under p:extLst


  Scenario: Round-trip sections through save and reload
    Given a presentation with three slides
     When I add a section named "Intro" with slides 1 and 2
      And I save and reload the presentation
     Then Presentation.sections has 1 section
      And the section is named "Intro"
      And the section has two assigned slides


  Scenario: Rename and remove a section
    Given a presentation with three slides
     When I add a section named "Intro" with slides 1 and 2
      And I rename the first section to "Overview"
     Then the section is named "Overview"
     When I remove the first section
     Then Presentation.sections has 0 sections
      And the presentation has no p:extLst element


  Scenario: Reorder sections with move_before and move_after
    Given a presentation with three slides
     When I add three sections "Intro", "Body", "Outro"
      And I move the "Outro" section before the "Intro" section
     Then the section names are "Outro", "Intro", "Body"
     When I move the "Intro" section after the "Body" section
     Then the section names are "Outro", "Body", "Intro"


  Scenario: Locate the section that owns a slide
    Given a presentation with three slides
     When I add a section named "Intro" with slides 1 and 2
      And I add a section named "Outro" with slide 3
     Then find_containing for slide 1 returns the "Intro" section
      And find_containing for slide 3 returns the "Outro" section


  Scenario: Adding a slide already in another section raises
    Given a presentation with three slides
     When I add a section named "Intro" with slides 1 and 2
      And I add a section named "Outro" with slide 3
     Then adding slide 1 to the "Outro" section raises a ValueError mentioning "Intro"


  Scenario: Move a slide between sections
    Given a presentation with three slides
     When I add a section named "Intro" with slides 1 and 2
      And I add a section named "Outro" with slide 3
      And I move slide 2 into the "Outro" section
     Then the "Intro" section contains slides 1
      And the "Outro" section contains slides 3 and 2
