Feature: Round-trip fidelity coverage round 2
  In order to widen round-trip confidence across shape/section/animation/props areas
  As a developer using python-pptx
  I need additional end-to-end scenarios that exercise save-and-reload across
  group shapes, sections, transitions, animations, extended properties,
  password-protected packages, chartex/smart_art passthrough, and cross-
  presentation operations


  # -- Cross-shape / group round-trips ---------------------------------

  Scenario: Group containing mixed content survives save and reload
    Given a fresh presentation with a group of an oval, a textbox, and a rectangle
     When I round-trip the presentation
     Then the reloaded group contains 3 child shapes
      And the reloaded group's textbox has text "label"


  Scenario: Nested group survives save and reload
    Given a fresh presentation with a two-level nested group
     When I round-trip the presentation
     Then the reloaded outer group has 1 group child and 1 non-group child
      And the reloaded inner group has 2 autoshape children


  Scenario: Group with named shapes preserves names on reload
    Given a fresh presentation with a group of three rectangles named A B C
     When I round-trip the presentation
     Then the reloaded group's child shape names are "A" "B" "C"


  Scenario: Group bounding box recomputes correctly after reload
    Given a fresh presentation with a group of two autoshapes at known coords
     When I round-trip the presentation
     Then the reloaded group.left and group.top equal the minimum child offsets


  Scenario: Ungroup after reload hoists children to the slide
    Given a fresh presentation with a group of two autoshapes
     When I round-trip the presentation
      And I call ungroup on the reloaded group
     Then the reloaded slide has 2 top-level autoshape children
      And the reloaded slide has 0 groups


  # -- Sections round-trip --------------------------------------------

  Scenario: Rename + remove survives save and reload
    Given a fresh presentation with a section renamed "Main"
     When I round-trip the presentation
     Then the reloaded presentation has 1 section named "Main"
     When I remove the only section from the reloaded presentation
      And I round-trip the presentation
     Then the reloaded presentation has 0 sections


  Scenario: Reorder sections survives save and reload
    Given a fresh presentation with three sections "A" "B" "C"
     When I reorder the sections to "C" "A" "B"
      And I round-trip the presentation
     Then the reloaded section names are "C", "A", "B"


  Scenario: Explicit section id survives save and reload
    Given a fresh presentation with a section authored with a fixed GUID
     When I round-trip the presentation
     Then the reloaded first section id equals the fixed GUID


  Scenario: Move slide between sections and round-trip
    Given a fresh presentation with two sections splitting four slides 2/2
     When I move slide 3 into the first section
      And I round-trip the presentation
     Then the reloaded first section has 3 slides
      And the reloaded second section has 1 slide


  Scenario: find_containing returns None for unassigned slide after reload
    Given a fresh presentation with sections that leave slide 3 unassigned
     When I round-trip the presentation
     Then find_containing for slide 3 is None on the reloaded presentation


  # -- Transition round-trip coverage ---------------------------------

  Scenario: Transition speed survives save and reload
    Given a fresh presentation with transition speed SLOW on slide 0
     When I round-trip the presentation
     Then the reloaded transition.speed is PP_TRANSITION_SPEED.SLOW


  Scenario: Transition wipe direction survives save and reload
    Given a fresh presentation with a wipe transition direction RIGHT on slide 0
     When I round-trip the presentation
     Then the reloaded transition.wipe_direction is PP_TRANSITION_SIDE_DIRECTION.RIGHT


  Scenario: Transition advance_after_time cleared then round-trips
    Given a fresh presentation with transition advance_after_time set to 5000
     When I clear advance_after_time on that slide
      And I round-trip the presentation
     Then the reloaded transition.advance_after_time is None


  Scenario: Per-slide distinct transitions round-trip independently
    Given a fresh presentation with slide 0 FADE and slide 1 MORPH
     When I round-trip the presentation
     Then the reloaded slide 0 transition.type is PP_TRANSITION_TYPE.FADE
      And the reloaded slide 1 transition.type is PP_TRANSITION_TYPE.MORPH


  Scenario: Transition cleared to NONE removes the element and round-trips
    Given a fresh presentation with a FADE transition then cleared
     When I round-trip the presentation
     Then the reloaded transition.type is PP_TRANSITION_TYPE.NONE


  # -- Animation round-trip coverage ----------------------------------

  Scenario: Shape animation type FADE_IN round-trips the delay 0 default
    Given a fresh presentation with a shape carrying a FADE_IN animation
     When I round-trip the presentation
     Then the reloaded shape.animation.type is MSO_ANIMATION_TYPE.FADE_IN
      And the reloaded shape.animation.delay is 0


  Scenario: Shape animation PULSE emphasis round-trips
    Given a fresh presentation with a shape carrying a PULSE emphasis animation
     When I round-trip the presentation
     Then the reloaded shape.animation.type is MSO_ANIMATION_TYPE.PULSE


  Scenario: Shape animation FADE_OUT exit round-trips
    Given a fresh presentation with a shape carrying a FADE_OUT exit animation
     When I round-trip the presentation
     Then the reloaded shape.animation.type is MSO_ANIMATION_TYPE.FADE_OUT


  Scenario: Cleared animation does not resurrect on reload
    Given a fresh presentation with a shape carrying a FADE_IN animation
     When I clear the shape's animation
      And I round-trip the presentation
     Then the reloaded shape.animation is None


  Scenario: Multiple animated shapes round-trip their sequence order
    Given a fresh presentation with two shapes authored FADE_IN then PULSE
     When I round-trip the presentation
     Then the reloaded animation_sequence preset_classes begin with "entr" then "emph"


  # -- Extended properties round-trip ---------------------------------

  Scenario: ExtendedProperties.company round-trips via save and reload
    Given a fresh presentation with extended company "Acme Corp"
     When I round-trip the presentation
     Then the reloaded extended_properties.company is "Acme Corp"


  Scenario: ExtendedProperties.manager round-trips via save and reload
    Given a fresh presentation with extended manager "Jane Doe"
     When I round-trip the presentation
     Then the reloaded extended_properties.manager is "Jane Doe"


  Scenario: ExtendedProperties application / app_version / hyperlink_base round-trip
    Given a fresh presentation with authored app / app_version / hyperlink_base
     When I round-trip the presentation
     Then the reloaded extended_properties.application is "python-pptx"
      And the reloaded extended_properties.app_version is "1.0"
      And the reloaded extended_properties.hyperlink_base is "https://example.com/"


  Scenario: ExtendedProperties.slide_count tracks additions through round-trip
    Given a fresh presentation with two additional blank slides
     When I round-trip the presentation
     Then the reloaded extended_properties.slide_count equals 3


  # -- Password-protected save/reload ---------------------------------

  Scenario: Password-protected round-trip via save(password=...)
    Given a fresh presentation with a textbox "secret"
     When I save the presentation encrypted with password "s3cr3t"
      And I reopen the encrypted package with the correct password
     Then the reloaded textbox text is "secret"


  Scenario: Encrypted package without password raises EncryptedPackageError
    Given a fresh presentation with a textbox "secret"
     When I save the presentation encrypted with password "s3cr3t"
     Then opening the package without a password raises EncryptedPackageError


  Scenario: Encrypted package with wrong password raises EncryptedPackageError
    Given a fresh presentation with a textbox "secret"
     When I save the presentation encrypted with password "s3cr3t"
     Then opening the package with the wrong password raises EncryptedPackageError


  Scenario: Password-protected save preserves core property
    Given a fresh presentation whose core title is "Encrypted Deck"
     When I save the presentation encrypted with password "pw"
      And I reopen the encrypted package with password "pw"
     Then the reloaded presentation core_properties.title is "Encrypted Deck"


  # -- chartex passthrough deep round-trip ----------------------------

  Scenario: Chartex part survives two save/reload cycles
    Given a presentation containing a chartex chart
     When I save and reload the chartex presentation twice
     Then the chartex shapes still number 2 after the second reload


  Scenario: Chartex slide still reports MSO_SHAPE_TYPE.CHART on reload
    Given a presentation containing a chartex chart
     When I save the presentation and reload it
     Then every chartex shape on the reloaded slide has shape_type CHART


  # -- Cross-presentation / merge -------------------------------------

  Scenario: add_slide_from_external carries a shape across presentations
    Given a source presentation with a slide that has a textbox "hello"
      And a target presentation with the default layout
     When I call add_slide_from_external on the target from the source slide
      And I round-trip the target presentation
     Then the reloaded target has a slide with a textbox "hello"


  Scenario: Presentation.merge appends source slides and round-trips
    Given a source presentation with two distinct slides A and B
      And a fresh target presentation
     When I call prs.merge(other) from target onto source
      And I round-trip the target presentation
     Then the reloaded target presentation has 2 slides
      And the reloaded target's first slide has a textbox "A"
      And the reloaded target's second slide has a textbox "B"


  Scenario: Merge into self raises ValueError
    Given a fresh presentation with a single blank slide
     Then calling presentation.merge with itself raises ValueError


  Scenario: add_slide_from_external across-presentation preserves named shape
    Given a source presentation with a slide whose shape is named "Hero"
      And a target presentation with the default layout
     When I call add_slide_from_external on the target from the source slide
      And I round-trip the target presentation
     Then the reloaded target's cloned slide has a shape named "Hero"
