Feature: Presentation properties
  In order to interact with a presentation
  As a developer using python-pptx
  I need read/write properties on the presentation object


  Scenario: Get Presentation.slide_width, .slide_height
    Given a presentation
     Then its slide width matches its known value
      And its slide height matches its known value


  Scenario: Set Presentation.slide_width, .slide_height
    Given a presentation
     When I change the slide width and height
     Then the slide width matches the new value
      And the slide height matches the new value


  Scenario Outline: Presentation.notes_master
    Given a presentation having <a-or-no> notes master
     Then prs.notes_master is a NotesMaster object
      And len(notes_master.shapes) is <shape-count>

    Examples: Notes master states
      | a-or-no | shape-count |
      | a       |      7      |
      | no      |      6      |


  Scenario: Presentation.slides
    Given a presentation
     Then prs.slides is a Slides object


  Scenario: Presentation.slide_masters
    Given a presentation
     Then prs.slide_masters is a SlideMasters object
