Feature: Change appearance of font used to render text
  In order to fine-tune the appearance of text
  As a developer using python-pptx
  I need a set of properties on the font used to render text


  Scenario Outline: Get Font.bold
    Given a font with bold set <bold-state>
     Then font.bold is <expected-value>

    Examples: font.bold states
      | bold-state | expected-value |
      | on         | True           |
      | off        | False          |
      | to inherit | None           |


  Scenario Outline: Set Font.bold
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


  Scenario Outline: Get Font.italic
    Given a font with italic set <italic-state>
     Then font.italic is <expected-value>

    Examples: font.italic states
      | italic-state | expected-value |
      | on           | True           |
      | off          | False          |
      | to inherit   | None           |


  Scenario Outline: Set Font.italic
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


  Scenario Outline: Get Font.language_id
    Given a font having language id <lang-id-state>
     Then font.language_id is MSO_LANGUAGE_ID.<member>

    Examples: font.language_id states
      | lang-id-state          | member |
      | of no explicit setting | NONE   |
      | MSO_LANGUAGE_ID.POLISH | POLISH |


  Scenario Outline: Set Font.language_id
    Given a font having language id <initial-state>
     When I assign <new-value> to font.language_id
     Then font.language_id is MSO_LANGUAGE_ID.<member>

    Examples: font.language_id assignment state changes
      | initial-state          | new-value              | member |
      | of no explicit setting | MSO_LANGUAGE_ID.FRENCH | FRENCH |
      | MSO_LANGUAGE_ID.FRENCH | MSO_LANGUAGE_ID.POLISH | POLISH |
      | MSO_LANGUAGE_ID.POLISH | MSO_LANGUAGE_ID.NONE   | NONE   |
      | MSO_LANGUAGE_ID.FRENCH | None                   | NONE   |


  Scenario Outline: Get Font.underline
    Given a font with underline set <underline-state>
     Then font.underline is <expected-value>

    Examples: font.underline states
      | underline-state | expected-value |
      | on              | True           |
      | off             | False          |
      | to inherit      | None           |
      | to DOUBLE_LINE  | DOUBLE_LINE    |
      | to WAVY_LINE    | WAVY_LINE      |


  Scenario Outline: Set Font.underline
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


  Scenario Outline: Get Font.size
    Given a font having size of <value>
     Then font.size is <reported-size>

    Examples: Font sizes
      | value             | reported-size |
      | no explicit value | None          |
      | 42pt              | 42.0 points   |


  Scenario: Set Font.name
    Given a font
     When I assign a typeface name to the font
     Then the font name matches the typeface I set


  Scenario: Add hyperlink
    Given a text run
     When I set the hyperlink address
     Then run.text is a hyperlink


  Scenario: Add hyperlink in table cell
    Given a text run in a table cell
     When I set the hyperlink address
     Then run.text is a hyperlink


  Scenario: Remove hyperlink
    Given a text run having a hyperlink
     When I assign None to hyperlink.address
     Then run.text is not a hyperlink
