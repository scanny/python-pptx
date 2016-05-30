Feature: Change paragraph properties
  In order to change the formatting of text to my needs
  As a developer using python-pptx
  I need a set of read/write properties on paragraph objects


  Scenario: Set Paragraph.alignment
     Given a paragraph
      When I set the paragraph alignment to centered
       And I reload the presentation
      Then the paragraph is aligned centered


  Scenario: Set Paragraph.indentation
     Given a paragraph
      When I indent the paragraph
       And I reload the presentation
      Then the paragraph is indented to the second level


  Scenario Outline: Get Paragraph.space_before/after
    Given a paragraph having space <before-after> of <setting>
     Then paragraph.space_<before-after> is <value>
      And paragraph.space_<before-after>.pt <pt-result>

    Examples: paragraph.space_before/after states
      | before-after | setting             | value | pt-result             |
      | before       | no explicit setting | None  | raises AttributeError |
      | before       | 6 pt                | 76200 | == 6.0                |
      | after        | no explicit setting | None  | raises AttributeError |
      | after        | 6 pt                | 76200 | == 6.0                |


  Scenario Outline: Set Paragraph.space_before/after
    Given a paragraph having space <before-after> of <setting>
     When I assign <new-value> to paragraph.space_<before-after>
     Then paragraph.space_<before-after> is <value>

    Examples: paragraph.space_before/after assignment results
      | before-after | setting             | new-value | value |
      | before       | no explicit setting |   76200   | 76200 |
      | before       | 6 pt                |   38100   | 38100 |
      | before       | 6 pt                |   None    | None  |
      | after        | no explicit setting |   76200   | 76200 |
      | after        | 6 pt                |   38100   | 38100 |
      | after        | 6 pt                |   None    | None  |


  Scenario Outline: Get Paragraph.line_spacing
    Given a paragraph having line spacing of <setting>
     Then paragraph.line_spacing is <value>
      And paragraph.line_spacing.pt <pt-result>

    Examples: paragraph.line_spacing states
      | setting             | value  | pt-result             |
      | no explicit setting |  None  | raises AttributeError |
      | 1.5 lines           |  1.5   | raises AttributeError |
      | 20 pt               | 254000 | is 20.0               |


  Scenario Outline: Set Paragraph.line_spacing
    Given a paragraph having line spacing of <setting>
     When I assign <new-value> to paragraph.line_spacing
     Then paragraph.line_spacing is <value>
      And paragraph.line_spacing.pt <pt-result>

    Examples: paragraph.line_spacing assignment results
     | setting             | new-value  | value  | pt-result             |
     | no explicit setting |    1.5     |  1.5   | raises AttributeError |
     | 1.5 lines           |    2.0     |  2.0   | raises AttributeError |
     | 1.5 lines           |    None    | None   | raises AttributeError |
     | no explicit setting |   254000   | 254000 | is 20.0               |
     | 20 pt               |   304800   | 304800 | is 24.0               |
     | 20 pt               |    None    | None   | raises AttributeError |
