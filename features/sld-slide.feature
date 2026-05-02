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


  Scenario Outline: Slide.follow_master_background
    Given a Slide object having <default-or-overridden> background as slide
     Then slide.follow_master_background is <value>

    Examples: Slide.follow_master_background cases
      | default-or-overridden | value |
      | the default           | True  |
      | an overridden         | False |


  Scenario: Slide.follow_master_background() resets an overridden background
    Given a Slide object having an overridden background as slide
     Then slide.follow_master_background is False
     When I call slide.follow_master_background()
     Then slide.follow_master_background is True
      And the call returned the slide itself


  Scenario: Slide.follow_master_background() is a no-op when already following master
    Given a Slide object having the default background as slide
     Then slide.follow_master_background is True
     When I call slide.follow_master_background()
     Then slide.follow_master_background is True
      And the call returned the slide itself


  Scenario Outline: Slide.has_notes_slide
    Given a slide having <a-or-no> notes slide
     Then slide.has_notes_slide is <value>

    Examples: Slide.has_notes_slide states
      | a-or-no | value |
      | a       | True  |
      | no      | False |


  Scenario: Slide.is_hidden defaults to False on a fresh slide
    Given a slide
     Then slide.is_hidden is False


  Scenario: Slide.is_hidden can be set and round-trips
    Given a slide
     When I set slide.is_hidden to True
     Then slide.is_hidden is True
      And after a save/load round-trip slide.is_hidden is True


  Scenario: Slide.is_hidden can be cleared and round-trips
    Given a slide
     When I set slide.is_hidden to True
      And I set slide.is_hidden to False
     Then slide.is_hidden is False
      And after a save/load round-trip slide.is_hidden is False


  Scenario: Slide.show_master_shapes defaults to True on a fresh slide
    Given a slide
     Then slide.show_master_shapes is True


  Scenario: Slide.show_master_shapes can be set to False and round-trips
    Given a slide
     When I set slide.show_master_shapes to False
     Then slide.show_master_shapes is False
      And after a save/load round-trip slide.show_master_shapes is False


  Scenario: Slide.show_master_shapes can be restored and round-trips
    Given a slide
     When I set slide.show_master_shapes to False
      And I set slide.show_master_shapes to True
     Then slide.show_master_shapes is True
      And after a save/load round-trip slide.show_master_shapes is True


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
      | of no explicit value | Layout 1         |


  Scenario: SlideLayout.shapes
    Given a SlideLayout object as slide
     Then slide.shapes is a LayoutShapes object


  Scenario: SlideLayout.placeholders
    Given a SlideLayout object as slide
     Then slide.placeholders is a LayoutPlaceholders object


  Scenario: SlideLayout.slide_master
    Given a SlideLayout object as slide_layout
     Then slide_layout.slide_master is a SlideMaster object


  Scenario: SlideLayout.used_by_slides
    Given a SlideLayout object used by a slide as slide_layout
      And a Slide object based on slide_layout as slide
     Then slide in slide_layout.used_by_slides is True


  Scenario: SlideLayout.used_by_slides
    Given a SlideLayout object used by no slides as slide_layout
     Then slide_layout.used_by_slides == ()


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


  Scenario Outline: SlideMaster.name falls back to positional name
    Given a presentation with <n-masters> slide masters and no explicit master name
     Then prs.slide_masters[<idx>].name is <expected-name>

    Examples: positional-fallback cases (issue #679)
      | n-masters | idx | expected-name |
      | 1         | 0   | Master 1      |
      | 2         | 0   | Master 1      |
      | 2         | 1   | Master 2      |


  Scenario: SlideMaster.name can be set and round-trips
    Given a presentation with 1 slide master and no explicit master name
     When I set prs.slide_masters[0].name to "Office Theme"
     Then prs.slide_masters[0].name is Office Theme
      And after a save/load round-trip prs.slide_masters[0].name is Office Theme


  Scenario: SlideMaster.name clears back to positional fallback
    Given a presentation with 1 slide master and no explicit master name
     When I set prs.slide_masters[0].name to "Office Theme"
      And I set prs.slide_masters[0].name to None
     Then prs.slide_masters[0].name is Master 1


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
