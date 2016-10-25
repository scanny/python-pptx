Feature: Notes slide properties
  In order to interact with a notes slide
  As a developer using python-pptx
  I need properties and methods on NotesSlide


  Scenario: NotesSlide.placeholders
    Given a notes slide
     Then notes_slide.placeholders is a NotesSlidePlaceholders object
      And iterating produces 3 NotesSlidePlaceholder objects


  Scenario: NotesSlide.notes_placeholder
     Given a notes slide
      Then notes_slide.notes_placeholder is a NotesSlidePlaceholder object


  Scenario: NotesSlide.notes_text_frame
     Given a notes slide
      Then notes_slide.notes_text_frame is a TextFrame object
