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
      | solid       | SOLID      |


  Scenario Outline: FillFormat.fore_color
    Given a FillFormat object as fill
     When I call fill.<type-setter>()
     Then fill.fore_color is a ColorFormat object

    Examples: Fill types
      | type-setter |
      | solid       |
