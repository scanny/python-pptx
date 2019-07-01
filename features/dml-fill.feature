Feature: Control fill
  In order to fine-tune the visual experience of filled areas
  As a developer using python-pptx
  I need properties and methods on FillFormat


  Scenario Outline: FillFormat type getters
    Given <type> FillFormat object as fill
     Then fill.type == <value>

    Examples: Fill types
      | type          | value               |
      | an inheriting | None                |
      | a no-fill     | MSO_FILL.BACKGROUND |
      | a solid       | MSO_FILL.SOLID      |
      | a picture     | MSO_FILL.PICTURE    |
      | a gradient    | MSO_FILL.GRADIENT   |
      | a patterned   | MSO_FILL.PATTERNED  |


  Scenario Outline: FillFormat type setters
    Given a FillFormat object as fill
     When I call fill.<type-setter>()
     Then fill.type == <value>

    Examples: Fill types
      | type-setter | value               |
      | background  | MSO_FILL.BACKGROUND |
      | gradient    | MSO_FILL.GRADIENT   |
      | solid       | MSO_FILL.SOLID      |
      | patterned   | MSO_FILL.PATTERNED  |


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


  Scenario: FillFormat.gradient_angle getter
    Given a gradient FillFormat object as fill
     Then fill.gradient_angle == 90.0


  Scenario Outline: FillFormat.gradient_angle setter
    Given a gradient FillFormat object as fill
     When I assign <new-value> to fill.gradient_angle
     Then fill.gradient_angle == <value>

    Examples: angle value cases
      | new-value | value |
      | 42.42     | 42.42 |
      | 270.0     | 270.0 |
      | 480.0     | 120.0 |
      | -90.0     | 270.0 |
      | -942.4    | 137.6 |


  Scenario: FillFormat.gradient_stops
    Given a gradient FillFormat object as fill
     Then fill.gradient_stops is a _GradientStops object


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


  Scenario: _GradientStop.color
    Given a _GradientStop object as stop
     Then stop.color is a ColorFormat object


  Scenario: _GradientStop.position getter
    Given a _GradientStop object as stop
     Then stop.position == 0.20


  Scenario: _GradientStop.position setter
    Given a _GradientStop object as stop
     When I assign 0.42 to stop.position
     Then stop.position == 0.42
