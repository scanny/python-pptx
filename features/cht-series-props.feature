Feature: Get and set chart series properties
  In order to customize the formatting of a series on a chart
  As a developer using python-pptx
  I need a way to get and set series properties

  @wip
  Scenario Outline: Get series fill type
    Given a bar series having fill of <fill>
     Then the series has a fill type of <expected-fill-type>

    Examples: Fill types
      | fill      | expected-fill-type       |
      | Automatic | None                     |
      | No Fill   | MSO_FILL_TYPE.BACKGROUND |
      | Accent 1  | MSO_FILL_TYPE.SOLID      |
      | Orange    | MSO_FILL_TYPE.SOLID      |


  @wip
  Scenario: Get series RGB color
    Given a bar series having fill of Orange
     Then the series fill RGB color is FF6600


  @wip
  Scenario: Get series theme color
    Given a bar series having fill of Accent 1
     Then the series fill theme color is Accent 1
