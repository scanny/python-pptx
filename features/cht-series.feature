Feature: Chart series
  In order to customize the formatting of a series on a chart
  As a developer using python-pptx
  I need access to series objects and their properties


  Scenario Outline: SeriesCollection.__len__()
    Given a SeriesCollection object for a <container> having <count> series
     Then len(series_collection) is <count>

    Examples: series container types
      | container         | count |
      | single-plot chart |   3   |
      | multi-plot chart  |   5   |
      | plot              |   3   |


  Scenario Outline: SeriesCollection.__getitem__()
    Given a SeriesCollection object for a <container> having <count> series
     Then series_collection[2] is a Series object

    Examples: series container types
      | container         | count |
      | single-plot chart |   3   |
      | multi-plot chart  |   5   |
      | plot              |   3   |


  Scenario Outline: SeriesCollection.__iter__()
    Given a SeriesCollection object for a <container> having <count> series
     Then iterating series_collection produces <count> Series objects

    Examples: series container types
      | container         | count |
      | single-plot chart |   3   |
      | multi-plot chart  |   5   |
      | plot              |   3   |


  Scenario Outline: series.data_labels
    Given a <series-type>Series object as series
     Then series.data_labels is a DataLabels object

    Examples: Series types
      | series-type |
      | Area        |
      | Bar         |
      | Doughnut    |
      | Line        |
      | Pie         |
      | Radar       |


  Scenario Outline: BarSeries.invert_if_negative value
    Given a BarSeries object having invert_if_negative of <setting> as series
     Then series.invert_if_negative is <expected-value>

    Examples: invert_if_negative settings
      | setting             | expected-value |
      | no explicit setting | True           |
      | True                | True           |
      | False               | False          |


  Scenario Outline: BarSeries.invert_if_negative setter
    Given a BarSeries object having invert_if_negative of <setting> as series
     When I assign <new-value> to series.invert_if_negative
     Then series.invert_if_negative is <expected-value>

    Examples: invert_if_negative settings
      | setting             | new-value | expected-value |
      | no explicit setting | True      | True           |
      | no explicit setting | False     | False          |
      | True                | True      | True           |
      | False               | False     | False          |


  Scenario: series.format
    Given a series
     Then series.format is a ChartFormat object


  Scenario Outline: series.format.fill.type
    Given a BarSeries object having <fill> fill as series
     Then series.format.fill.type is <expected-fill-type>

    Examples: Fill types
      | fill      | expected-fill-type       |
      | Automatic | None                     |
      | No Fill   | MSO_FILL_TYPE.BACKGROUND |
      | Accent 1  | MSO_FILL_TYPE.SOLID      |
      | Orange    | MSO_FILL_TYPE.SOLID      |


  Scenario: series.format.fill.fore_color.rgb
    Given a BarSeries object having Orange fill as series
     Then series.format.fill.fore_color.rgb is FF6600


  Scenario: series.format.fill.fore_color.theme_color
    Given a BarSeries object having Accent 1 fill as series
     Then series.format.fill.fore_color.theme_color is Accent 1


  Scenario Outline: series.format.line.width
    Given a BarSeries object having <width> line as series
     Then series.format.line.width is <expected-width>

    Examples: line widths
      | width   | expected-width |
      | no      | 0              |
      | 1 point | 12700          |


  Scenario Outline: series.marker
    Given a <series-type>Series object as series
     Then series.marker is a Marker object

    Examples: Series types having .marker property
      | series-type |
      | Line        |
      | Xy          |
      | Radar       |


  Scenario Outline: series.points
    Given a <series-type>Series object as series
     Then series.points is a <type-name> object

    Examples: series points classes
      | series-type | type-name      |
      | Xy          | XyPoints       |
      | Bubble      | BubblePoints   |
      | Category    | CategoryPoints |


  Scenario Outline: series.values
    Given a BarSeries object having values <values> as series
     Then series.values is <expected-value>

    Examples: series values
      | values         | expected-value   |
      | 1.2, 2.3, 3.4  | (1.2, 2.3, 3.4)  |
      | 4.5, None, 6.7 | (4.5, None, 6.7) |
