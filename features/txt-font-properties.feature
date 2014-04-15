Feature: Change appearance of font used to render text
  In order to fine-tune the appearance of text
  As a developer using python-pptx
  I need a set of properties on the font used to render text


  @wip
  Scenario Outline: Get font size
    Given a font having <applied size>
     Then the font size property value is <reported size>

    Examples: Font sizes
      | applied size                    | reported size |
      | a directly applied size of 42pt | 42.0 points   |
      | no directly applied size        | None          |


  Scenario: Change font typeface
    Given a font
     When I assign a typeface name to the font
     Then the font name matches the typeface I set


  Scenario Outline: Set italics property of text
    Given a run with italics set <initial>
     When I set italics <new>
      And I reload the presentation
     Then the run that had italics set <initial> now has it set <new>

    Examples: Italics Settings
      | initial | new     |
      | on      | on      |
      | on      | off     |
      | on      | to None |
      | off     | on      |
      | off     | off     |
      | off     | to None |
      | to None | on      |
      | to None | off     |
      | to None | to None |


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
