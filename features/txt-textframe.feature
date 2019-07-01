Feature: Get and change properties of shape text frame
  In order configure the text container of a shape to my needs
  As a developer using python-pptx
  I need a set of properties on the text frame of a shape


  Scenario Outline: TextFrame.auto_size getter
    Given a TextFrame object having auto-size of <setting> as text_frame
     Then text_frame.auto_size is <value>

    Examples: Auto-size settings
      | setting           | value                           |
      | None              | None                            |
      | no auto-size      | MSO_AUTO_SIZE.NONE              |
      | fit shape to text | MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT |
      | fit text to shape | MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE |


  Scenario Outline: TextFrame.auto_size setter
    Given a TextFrame object as text_frame
     When I assign <value> to text_frame.auto_size
     Then text_frame.auto_size is <value>

    Examples: Auto-size settings
      | value                           |
      | None                            |
      | MSO_AUTO_SIZE.NONE              |
      | MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT |
      | MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE |


  Scenario Outline: TextFrame.margin_x setter
    Given a TextFrame object as text_frame
     When I assign text_frame.margin_<side> = Inches(<value>)
     Then text_frame.margin_<side>.inches == <value>

    Examples: TextFrame.margin_{side} value cases
      | side   | value |
      | left   | 0.1   |
      | top    | 0.2   |
      | right  | 0.3   |
      | bottom | 0.4   |


  Scenario Outline: TextFrame.text getter
    Given a TextFrame object containing <value> as text_frame
     Then text_frame.text == <value>

    Examples: Textframe paragraph and line break combinations
      | value     |
      | "abc"     |
      | "a\nb\nc" |


  Scenario Outline: TextFrame.text setter
    Given a TextFrame object as text_frame
     When I assign text_frame.text = <value>
     Then text_frame.text == <value>

    Examples: Textframe paragraph and line break combinations
      | value     |
      | "abc"     |
      | "a\nb\nc" |


  Scenario Outline: TextFrame.word_wrap setter
    Given a TextFrame object as text_frame
     When I assign <value> to text_frame.word_wrap
     Then text_frame.word_wrap is <value>

    Examples: Word-wrap Settings
      | value |
      | True  |
      | False |
      | None  |
