Feature: Compute the text-rendering rectangle of a shape
  In order to pick a font size that will not overflow a shape
  As a developer using python-pptx
  I need a read-only property that returns the (left, top, width, height)
  rectangle that PowerPoint allocates for rendering text inside a shape,
  accounting for the shape's position, size, and text-frame insets.


  Scenario: BaseShape.text_frame_rect on a text-box with default margins
    Given a text-box sized 4" x 1" positioned at (1", 2") with default insets
     When I read its text_frame_rect
     Then text_frame_rect.left.inches == 1.1
      And text_frame_rect.top.inches == 2.05
      And text_frame_rect.width.inches == 3.8
      And text_frame_rect.height.inches == 0.9


  Scenario: BaseShape.text_frame_rect raises on a shape with no text frame
    Given a connector shape
     Then reading text_frame_rect raises ValueError
