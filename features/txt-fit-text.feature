Feature: Resize text to fit shape
  In order to reliably fit text into a shape of fixed size
  As a developer using python-pptx
  I need a way to reduce the point size of text to fit within its shape

  Scenario: Set text size to best-fit point size
    Given a text frame with more text than will fit
     When I call TextFrame.fit_text()
     Then text_frame.auto_size is MSO_AUTO_SIZE.NONE
      And text_frame.word_wrap is True
      And the size of the text is 10pt
