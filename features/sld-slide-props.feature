Feature: slide properties
  In order to interact with a slide, layout, or master
  As a developer using python-pptx
  I need properties and methods on Slide, SlideLayout, and SlideMaster


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
    Given a slide layout having name <name>
     Then slide_layout.name is <value>

    Examples: name scenarios
      | name                 | value            |
      | Bullet Layout        | Bullet Layout    |
      | of no explicit value | the empty string |


  Scenario: SlideLayout.shapes
    Given a slide layout
     Then slide_layout.shapes is a LayoutShapes object


  Scenario: SlideLayout.placeholders
    Given a slide layout
     Then slide_layout.placeholders is a LayoutPlaceholders object


  Scenario: SlideLayout.slide_master
    Given a slide layout
     Then slide_layout.slide_master is a SlideMaster object


  Scenario: SlideMaster.shapes
    Given a slide master
     Then slide_master.shapes is a MasterShapes object


  Scenario: SlideMaster.placeholders
    Given a slide master
     Then slide_master.placeholders is a MasterPlaceholders object


  Scenario: SlideMaster.slide_layouts
    Given a slide master
     Then slide_master.slide_layouts is a SlideLayouts object
