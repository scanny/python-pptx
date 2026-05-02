Feature: shape-animation round-trip (issue #102 MVP)
  In order to bind entrance / emphasis / exit animations to shapes
  As a developer using python-pptx
  I need Shape.set_animation and Shape.animation to round-trip through save/load


  Scenario: Shape.animation is None on a plain rectangle
    Given a blank slide
     When I add a rectangle shape
     Then shape.animation is None


  Scenario: Shape.set_animation applies a fade-in entrance
    Given a blank slide
     When I add a rectangle shape
      And I set shape.animation to FADE_IN with delay 500
     Then shape.animation.type is FADE_IN
      And shape.animation.delay is 500
      And shape.animation.trigger is ON_CLICK


  Scenario: Shape animation round-trips through save and reopen
    Given a blank slide
     When I add a rectangle shape
      And I set shape.animation to FLY_IN with delay 250
      And I save and reopen the presentation preserving the shape
     Then shape.animation.type is FLY_IN
      And shape.animation.delay is 250


  Scenario: set_animation(None) removes any existing animation
    Given a blank slide
     When I add a rectangle shape
      And I set shape.animation to PULSE with delay 0
      And I clear shape.animation
     Then shape.animation is None
