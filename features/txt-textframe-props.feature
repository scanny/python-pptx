Feature: Get and change properties of shape text frame
  In order configure the text container of a shape to my needs
  As a developer using python-pptx
  I need a set of properties on the text frame of a shape


  Scenario Outline: Get text frame auto-size setting
    Given a text frame having auto-size set to <setting>
     Then text_frame.auto_size is <value>

    Examples: Auto-size settings
      | setting           | value                           |
      | None              | None                            |
      | no auto-size      | MSO_AUTO_SIZE.NONE              |
      | fit shape to text | MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT |
      | fit text to shape | MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE |


  Scenario Outline: Change text frame auto-size setting
    Given a text frame
     When I assign <value> to text_frame.auto_size
     Then text_frame.auto_size is <value>

    Examples: Auto-size settings
      | value                           |
      | None                            |
      | MSO_AUTO_SIZE.NONE              |
      | MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT |
      | MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE |


  Scenario Outline: Set text frame margins
    Given a text frame
     When I set the <side> margin to <value>"
      And I reload the presentation
     Then the text frame's <side> margin is <value>"

    Examples: Italics Settings
      | side   | value |
      | left   | 0.1   |
      | top    | 0.2   |
      | right  | 0.3   |
      | bottom | 0.4   |


  Scenario Outline: Set word wrap property of text frame
    Given a text frame
     When I assign <value> to text_frame.word_wrap
      And I reload the presentation
     Then text_frame.word_wrap is <value>

    Examples: Word-wrap Settings
      | value |
      | True  |
      | False |
      | None  |
