Feature: Control fill
  In order to fine-tune the visual experience of filled areas
  As a developer using python-pptx
  I need properties and methods on FillFormat


  Scenario Outline: FillFormat type setters
    Given a FillFormat object as fill
     When I call fill.<type-setter>()
     Then fill.type is MSO_FILL.<type-name>

    Examples: Fill types
      | type-setter | type-name  |
      | background  | BACKGROUND |
      | patterned   | PATTERNED  |
      | solid       | SOLID      |


  Scenario: FillFormat.back_color
    Given a FillFormat object as fill
     When I call fill.patterned()
     Then fill.back_color is a ColorFormat object


  Scenario Outline: FillFormat.fore_color
    Given a FillFormat object as fill
     When I call fill.<type-setter>()
     Then fill.fore_color is a ColorFormat object

    Examples: Fill types
      | type-setter |
      | patterned   |
      | solid       |


  Scenario Outline: FillFormat.pattern getter
    Given a FillFormat object as fill having <pattern> fill
     Then fill.pattern is <value>

    Examples: Pattern fill types
      | pattern           | value             |
      | no pattern        | None              |
      | MSO_PATTERN.DIVOT | MSO_PATTERN.DIVOT |
      | MSO_PATTERN.WAVE  | MSO_PATTERN.WAVE  |


  Scenario: FillFormat.pattern setter
    Given a FillFormat object as fill
     When I call fill.patterned()
      And I assign MSO_PATTERN.CROSS to fill.pattern
     Then fill.pattern is MSO_PATTERN.CROSS
