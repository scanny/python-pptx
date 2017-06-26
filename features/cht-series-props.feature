Feature: Get and set series properties
  In order to customize the formatting of a series on a chart
  As a developer using python-pptx
  I need a way to get and set series properties


  Scenario: Get Series.format
    Given a series
     Then series.format is a ChartFormat object


  Scenario Outline: Get Series.marker
    Given a series of type <type>
     Then series.marker is a Marker object

    Examples: Series types having .marker property
      | type  |
      | Line  |
      | XY    |
      | Radar |


  Scenario Outline: Get series fill type
    Given a bar series having fill of <fill>
     Then series.format.fill.type is <expected-fill-type>

    Examples: Fill types
      | fill      | expected-fill-type       |
      | Automatic | None                     |
      | No Fill   | MSO_FILL_TYPE.BACKGROUND |
      | Accent 1  | MSO_FILL_TYPE.SOLID      |
      | Orange    | MSO_FILL_TYPE.SOLID      |


  Scenario: Get series RGB color
    Given a bar series having fill of Orange
     Then series.format.fill.fore_color.rgb is FF6600


  Scenario: Get series theme color
    Given a bar series having fill of Accent 1
     Then series.format.fill.fore_color.theme_color is Accent 1


  Scenario Outline: Get series line width
    Given a bar series having <width> line
     Then series.format.line.width is <expected-width>

    Examples: line widths
      | width   | expected-width |
      | no      | 0              |
      | 1 point | 12700          |


  Scenario Outline: Get invert_if_negative value
    Given a bar series having invert_if_negative of <setting>
     Then series.invert_if_negative is <expected-value>

    Examples: invert_if_negative settings
      | setting             | expected-value |
      | no explicit setting | True           |
      | True                | True           |
      | False               | False          |


  Scenario Outline: Change invert_if_negative value
    Given a bar series having invert_if_negative of <setting>
     When I assign <new-value> to series.invert_if_negative
     Then series.invert_if_negative is <expected-value>

    Examples: invert_if_negative settings
      | setting             | new-value | expected-value |
      | no explicit setting | True      | True           |
      | no explicit setting | False     | False          |
      | True                | True      | True           |
      | False               | False     | False          |


  Scenario Outline: Get series.values
    Given a bar series with values <values>
     Then series.values is <expected-value>

    Examples: series values
      | values         | expected-value   |
      | 1.2, 2.3, 3.4  | (1.2, 2.3, 3.4)  |
      | 4.5, None, 6.7 | (4.5, None, 6.7) |


  Scenario Outline: Get series points
    Given a series of type <series-type>
     Then series.points is a <type-name> object

    Examples: series points classes
      | series-type | type-name      |
      | XY          | XyPoints       |
      | Bubble      | BubblePoints   |
      | Category    | CategoryPoints |
