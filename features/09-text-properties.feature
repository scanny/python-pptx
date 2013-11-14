Feature: Change properties of text in shapes
  In order to change the formatting of text to my needs
  As a developer using python-pptx
  I need to set the properties of text in a shape

  Scenario: Set paragraph alignment
     Given I have a reference to a paragraph
      When I set the paragraph alignment to centered
       And I save the presentation
      Then the paragraph is aligned centered

  Scenario Outline: Set word wrap property of textframe
    Given a textframe
     When I set the textframe word wrap <value>
      And I save the presentation
     Then the textframe word wrap is set <value>

  Examples: Word-wrap Settings
    | value   |
    | on      |
    | off     |
    | to None |

  Scenario Outline: Set italics property of text
    Given a run with italics set <initial>
     When I set italics <new>
      And I save the presentation
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

  Scenario Outline: Set textframe margins
    Given a textframe
     When I set the <side> margin to <value>"
      And I save the presentation
     Then the textframe's <side> margin is <value>"

  Examples: Italics Settings
    | side   | value |
    | left   | 0.1   |
    | top    | 0.2   |
    | right  | 0.3   |
    | bottom | 0.4   |
