Feature: slide-transition round-trip (Foundation F8 MVP)
  In order to preserve slide-transition configuration across save/load
  As a developer using python-pptx
  I need Slide.transition to round-trip the `p:transition` element


  Scenario: Slide.transition.type round-trips a fade transition
    Given a blank slide
     When I set slide.transition.type to FADE
      And I set slide.transition.duration to 800
      And I set slide.transition.advance_on_click to False
      And I set slide.transition.advance_after_time to 3000
      And I save and reopen the presentation
     Then slide.transition.type is PP_TRANSITION_TYPE.FADE
      And slide.transition.duration is 800
      And slide.transition.advance_on_click is False
      And slide.transition.advance_after_time is 3000


  Scenario: Slide.has_animations is False for a clean slide
    Given a blank slide
     Then slide.has_animations is False
      And slide.timing_xml is None


  Scenario: Slide.transition defaults to NONE on a fresh slide
    Given a blank slide
     Then slide.transition.type is PP_TRANSITION_TYPE.NONE
      And slide.transition.advance_on_click is True
      And slide.transition.advance_after_time is None


  Scenario: Slide.transition.type round-trips a MORPH transition (#942)
    Given a blank slide
     When I set slide.transition.type to MORPH
      And I set slide.transition.morph_option to byWord
      And I set slide.transition.duration to 2000
      And I save and reopen the presentation
     Then slide.transition.type is PP_TRANSITION_TYPE.MORPH
      And slide.transition.morph_option is byWord
      And slide.transition.duration is 2000
      And slide.transition is wrapped in mc:AlternateContent


  Scenario: Switching from MORPH back to a plain variant unwraps (#942)
    Given a blank slide
     When I set slide.transition.type to MORPH
      And I set slide.transition.type to FADE
      And I save and reopen the presentation
     Then slide.transition.type is PP_TRANSITION_TYPE.FADE
      And slide.transition is not wrapped in mc:AlternateContent
