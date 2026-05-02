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
