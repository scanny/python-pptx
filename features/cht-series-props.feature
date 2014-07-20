Feature: Get and set chart series properties
  In order to customize the formatting of a series on a chart
  As a developer using python-pptx
  I need a way to get and set series properties

  Scenario Outline: Get series fill type
    Given a bar series having fill of <fill>
     Then the series has a fill type of <expected-fill-type>

    Examples: Fill types
      | fill      | expected-fill-type       |
      | Automatic | None                     |
      | No Fill   | MSO_FILL_TYPE.BACKGROUND |
      | Accent 1  | MSO_FILL_TYPE.SOLID      |
      | Orange    | MSO_FILL_TYPE.SOLID      |


  Scenario: Get series RGB color
    Given a bar series having fill of Orange
     Then the series fill RGB color is FF6600


  Scenario: Get series theme color
    Given a bar series having fill of Accent 1
     Then the series fill theme color is Accent 1


  @wip
  Scenario Outline: Get series line width
    Given a bar series having <width> line
     Then the series has a line width of <expected-width>

    Examples: line widths
      | width   | expected-width |
      | no      | 0              |
      | 1 point | 12700          |
