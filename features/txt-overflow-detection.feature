Feature: Detect text overflow
  In order to prevent content from being cut off
  As a developer using python-pptx
  I need a way to detect when text will overflow its shape

  Scenario: Detect when text will overflow
    Given a text frame with text that will overflow at 18pt
     When I call text_frame.will_overflow(font_size=18)
     Then it returns True

  Scenario: Detect when text will fit
    Given a text frame with text that will fit at 10pt
     When I call text_frame.will_overflow(font_size=10)
     Then it returns False

  Scenario: Get detailed overflow information
    Given a text frame with text that will overflow at 18pt
     When I call text_frame.overflow_info(font_size=18)
     Then info.will_overflow is True
      And info.required_height is greater than info.available_height
      And info.overflow_percentage is greater than 0
      And info.fits_at_font_size is less than 18

  Scenario: Get overflow info for fitting text
    Given a text frame with text that will fit at 10pt
     When I call text_frame.overflow_info(font_size=10)
     Then info.will_overflow is False
      And info.overflow_height is 0
      And info.fits_at_font_size is None

  Scenario: Handle empty text frame
    Given an empty text frame
     When I call text_frame.will_overflow()
     Then it returns False

  Scenario: Use effective font size
    Given a text frame with uniform 14pt text
     When I call text_frame.will_overflow() without specifying font_size
     Then it uses 14pt as the font size

  Scenario: Reject mixed font sizes
    Given a text frame with mixed font sizes
     When I call text_frame.will_overflow() without specifying font_size
     Then it raises ValueError
