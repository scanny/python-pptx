Feature: slide properties
  In order to interact with a slide, layout, master, or notes slide
  As a developer using python-pptx
  I need properties and methods on the slide object


  Scenario Outline: Slide.background
    Given a Slide object having <default-or-overridden> background as slide
     Then slide.background is a _Background object

    Examples: Slide.background cases
      | default-or-overridden |
      | the default           |
      | an overridden         |


  Scenario Outline: Slide.has_notes_slide
    Given a slide having <a-or-no> notes slide
     Then slide.has_notes_slide is <value>

    Examples: Slide.has_notes_slide states
      | a-or-no | value |
      | a       | True  |
      | no      | False |


  Scenario Outline: Slide.name
    Given a slide having name <name>
     Then slide.name is <value>

    Examples: name scenarios
      | name                 | value            |
      | Overview             | Overview         |
      | of no explicit value | the empty string |


  Scenario Outline: Slide.notes_slide
    Given a slide having <a-or-no> notes slide
     Then slide.notes_slide is a NotesSlide object
      And len(notes_slide.shapes) is <shape-count>

    Examples: Slide.notes_slide states
      | a-or-no | shape-count |
      | a       |      3      |
      | no      |      3      |


  Scenario: Slide.shapes
    Given a slide
     Then slide.shapes is a SlideShapes object


  Scenario: Slide.placeholders
    Given a slide
     Then slide.placeholders is a SlidePlaceholders object


  Scenario: Slide.slide_id
    Given a slide having slide id 256
     Then slide.slide_id is 256


  Scenario Outline: SlideLayout.name
    Given a SlideLayout object having name <name> as slide
     Then slide.name is <value>

    Examples: name scenarios
      | name                 | value            |
      | Bullet Layout        | Bullet Layout    |
      | of no explicit value | the empty string |


  Scenario: SlideLayout.shapes
    Given a SlideLayout object as slide
     Then slide.shapes is a LayoutShapes object


  Scenario: SlideLayout.placeholders
    Given a SlideLayout object as slide
     Then slide.placeholders is a LayoutPlaceholders object


  Scenario: SlideLayout.slide_master
    Given a SlideLayout object as slide_layout
     Then slide_layout.slide_master is a SlideMaster object


  Scenario: SlideMaster.background
    Given a SlideMaster object as slide
     Then slide.background is a _Background object


  Scenario: SlideMaster.shapes
    Given a SlideMaster object as slide
     Then slide.shapes is a MasterShapes object


  Scenario: SlideMaster.placeholders
    Given a SlideMaster object as slide
     Then slide.placeholders is a MasterPlaceholders object


  Scenario: SlideMaster.slide_layouts
    Given a SlideMaster object as slide_master
     Then slide_master.slide_layouts is a SlideLayouts object


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
