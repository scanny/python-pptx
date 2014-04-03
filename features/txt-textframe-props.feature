Feature: Get and change properties of shape text frame
  In order configure the text container of a shape to my needs
  As a developer using python-pptx
  I need a set of properties on the text frame of a shape


  Scenario Outline: Set textframe margins
    Given a textframe
     When I set the <side> margin to <value>"
      And I reload the presentation
     Then the textframe's <side> margin is <value>"

  Examples: Italics Settings
    | side   | value |
    | left   | 0.1   |
    | top    | 0.2   |
    | right  | 0.3   |
    | bottom | 0.4   |


  Scenario Outline: Set word wrap property of textframe
    Given a textframe
     When I set the textframe word wrap <value>
      And I reload the presentation
     Then the textframe word wrap is set <value>

  Examples: Word-wrap Settings
    | value   |
    | on      |
    | off     |
    | to None |

