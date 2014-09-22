Feature: Change appearance of font used to render text
  In order to fine-tune the appearance of text
  As a developer using python-pptx
  I need a set of properties on the font used to render text


  Scenario Outline: Get bold setting
    Given a font with bold set <bold-state>
     Then font.bold is <expected-value>

    Examples: font.bold states
      | bold-state | expected-value |
      | on         | True           |
      | off        | False          |
      | to inherit | None           |


  Scenario Outline: Change font.bold setting
    Given a font with bold set <initial-state>
     When I assign <new-value> to font.bold
     Then font.bold is <new-value>

    Examples: Expected results of changing font.bold setting
      | initial-state | new-value |
      | on            | True      |
      | off           | True      |
      | to inherit    | True      |
      | on            | False     |
      | off           | False     |
      | to inherit    | False     |
      | on            | None      |
      | off           | None      |
      | to inherit    | None      |


  Scenario Outline: Get italic setting
    Given a font with italic set <italic-state>
     Then font.italic is <expected-value>

    Examples: font.italic states
      | italic-state | expected-value |
      | on           | True           |
      | off          | False          |
      | to inherit   | None           |


  Scenario Outline: Change font.italic setting
    Given a font with italic set <initial-state>
     When I assign <new-value> to font.italic
     Then font.italic is <new-value>

    Examples: Expected results of changing font.italic setting
      | initial-state | new-value |
      | on            | True      |
      | off           | True      |
      | to inherit    | True      |
      | on            | False     |
      | off           | False     |
      | to inherit    | False     |
      | on            | None      |
      | off           | None      |
      | to inherit    | None      |


  Scenario Outline: Get underline setting
    Given a font with underline set <underline-state>
     Then font.underline is <expected-value>

    Examples: font.underline states
      | underline-state | expected-value |
      | on              | True           |
      | off             | False          |
      | to inherit      | None           |
      | to DOUBLE_LINE  | DOUBLE_LINE    |
      | to WAVY_LINE    | WAVY_LINE      |


  Scenario Outline: Change font.underline setting
    Given a font with underline set <initial-state>
     When I assign <new-value> to font.underline
     Then font.underline is <expected-value>

    Examples: Expected results of changing font.underline setting
      | initial-state  | new-value   | expected-value |
      | on             | True        | True           |
      | off            | SINGLE_LINE | True           |
      | to inherit     | True        | True           |
      | to WAVY_LINE   | False       | False          |
      | to inherit     | NONE        | False          |
      | to DOUBLE_LINE | None        | None           |
      | off            | DOUBLE_LINE | DOUBLE_LINE    |
      | to WAVY_LINE   | DOUBLE_LINE | DOUBLE_LINE    |


  Scenario Outline: Get font size
    Given a font having <applied-size>
     Then font.size is <reported-size>

    Examples: Font sizes
      | applied-size                    | reported-size |
      | a directly applied size of 42pt | 42.0 points   |
      | no directly applied size        | None          |


  Scenario: Change font typeface
    Given a font
     When I assign a typeface name to the font
     Then the font name matches the typeface I set


  Scenario: Add hyperlink
    Given a text run
     When I set the hyperlink address
     Then the text of the run is a hyperlink


  Scenario: Add hyperlink in table cell
    Given a text run in a table cell
     When I set the hyperlink address
     Then the text of the run is a hyperlink


  Scenario: Remove hyperlink
    Given a text run having a hyperlink
     When I set the hyperlink address to None
     Then the text run is not a hyperlink
