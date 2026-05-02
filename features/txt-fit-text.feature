Feature: Resize text to fit shape
  In order to reliably fit text into a shape of fixed size
  As a developer using python-pptx
  I need a way to reduce the point size of text to fit within its shape

  Scenario: Set text size to best-fit point size
    Given a text frame with more text than will fit
     When I call TextFrame.fit_text()
     Then text_frame.auto_size is MSO_AUTO_SIZE.NONE
      And text_frame.word_wrap is True
      And the size of the text is 10pt or 11pt

  Scenario: Raise TextLayoutError when no font size fits the shape
    Given a shape too narrow to fit any word at any considered font size
     When I call TextFitter.best_fit_font_size() on that shape
     Then a TextLayoutError is raised

  Scenario: Fit text whose longest word overflows the shape at max_size
    Given a shape whose first word overflows at max size but fits at a smaller size
     When I call TextFitter.best_fit_font_size() on that shape
     Then a smaller fitting point size is returned

  Scenario: Set font_scale and line_space_reduction on a placeholder text frame
    Given a placeholder text frame as text_frame
     When I assign 85.0 to text_frame.font_scale
      And I assign 20.0 to text_frame.line_space_reduction
     Then text_frame.font_scale is 85.0
      And text_frame.line_space_reduction is 20.0
