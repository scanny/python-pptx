Feature: SlideShapes.add_picture_from
  In order to copy a picture between slides with one call
  As a developer using python-pptx
  I need `SlideShapes.add_picture_from(other_picture, left=None, top=None,
  width=None, height=None)` — an ergonomic alias for CLO-2's
  `Picture.clone_onto` that re-embeds the image bytes on the destination
  slide's package and preserves crop, outline, and effects formatting


  Scenario: Clone a picture onto another slide with one call
    Given a slide with a picture
      And a second slide in the same presentation
     When I add the picture_from the source onto the second slide
     Then the clone is a Picture
      And the clone's image bytes match the source


  Scenario: Left and top override the clone's position
    Given a slide with a picture
      And a second slide in the same presentation
     When I add_picture_from the source at (3in, 4in)
     Then the clone's position is (3in, 4in)


  Scenario: Width and height override the clone's size
    Given a slide with a picture
      And a second slide in the same presentation
     When I add_picture_from the source with size (2in, 1in)
     Then the clone's size is (2in, 1in)


  Scenario: Clone a picture into a different presentation
    Given a source presentation with a picture on slide 1
      And a target presentation with one empty slide
     When I add_picture_from the source onto the target presentation's slide
     Then the clone is a Picture
      And the clone's image part lives in the target presentation's package
