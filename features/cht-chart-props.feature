Feature: Chart properties
  In order to customize the formatting of a chart
  As a developer using python-pptx
  I need read/write properties on Chart


  Scenario Outline: Chart.chart_title
    Given a chart having <a-or-no> title
     Then chart.chart_title is a ChartTitle object

    Examples: chart title presence cases
      | a-or-no |
      | a       |
      | no      |


  Scenario Outline: Get Chart.chart_type
    Given a chart of type <chart-type>
     Then chart.chart_type is <expected-enum-member>

    Examples: chart types
      | chart-type               | expected-enum-member         |
      | Area                     | AREA                         |
      | Stacked Area             | AREA_STACKED                 |
      | 100% Stacked Area        | AREA_STACKED_100             |
      | 3-D Area                 | THREE_D_AREA                 |
      | 3-D Stacked Area         | THREE_D_AREA_STACKED         |
      | 3-D 100% Stacked Area    | THREE_D_AREA_STACKED_100     |
      | Clustered Bar            | BAR_CLUSTERED                |
      | Stacked Bar              | BAR_STACKED                  |
      | 100% Stacked Bar         | BAR_STACKED_100              |
      | Clustered Column         | COLUMN_CLUSTERED             |
      | Stacked Column           | COLUMN_STACKED               |
      | 100% Stacked Column      | COLUMN_STACKED_100           |
      | Line                     | LINE                         |
      | Stacked Line             | LINE_STACKED                 |
      | 100% Stacked Line        | LINE_STACKED_100             |
      | Marked Line              | LINE_MARKERS                 |
      | Stacked Marked Line      | LINE_MARKERS_STACKED         |
      | 100% Stacked Marked Line | LINE_MARKERS_STACKED_100     |
      | Pie                      | PIE                          |
      | Exploded Pie             | PIE_EXPLODED                 |
      | XY (Scatter)             | XY_SCATTER                   |
      | XY Lines                 | XY_SCATTER_LINES             |
      | XY Lines No Markers      | XY_SCATTER_LINES_NO_MARKERS  |
      | XY Smooth Lines          | XY_SCATTER_SMOOTH            |
      | XY Smooth No Markers     | XY_SCATTER_SMOOTH_NO_MARKERS |
      | Bubble                   | BUBBLE                       |
      | 3D-Bubble                | BUBBLE_THREE_D_EFFECT        |
      | Radar                    | RADAR                        |
      | Marked Radar             | RADAR_MARKERS                |
      | Filled Radar             | RADAR_FILLED                 |


  Scenario Outline: Get Chart.category_axis
    Given a chart of type <chart-type>
     Then chart.category_axis is a <type-name> object

    Examples: category axis object types
      | chart-type                  | type-name    |
      | Stacked Bar                 | CategoryAxis |
      | Line (with date categories) | DateAxis     |
      | XY (Scatter)                | ValueAxis    |
      | Bubble                      | ValueAxis    |


  Scenario Outline: Get Chart.has_title
    Given a chart having <a-or-no> title
     Then chart.has_title is <expected-value>

    Examples: chart title presence cases
      | a-or-no | expected-value |
      | a       | True           |
      | no      | False          |


  Scenario Outline: Set Chart.has_title
    Given a chart having <a-or-no> title
     When I assign <value> to chart.has_title
     Then chart.has_title is <expected-value>

    Examples: chart title assignment cases
      | a-or-no | value | expected-value |
      | a       | True  | True           |
      | a       | False | False          |
      | no      | True  | True           |
      | no      | False | False          |


  Scenario Outline: Get Chart.value_axis
    Given a chart of type <chart-type>
     Then chart.value_axis is a ValueAxis object

    Examples: value axis object types
      | chart-type   |
      | Stacked Bar  |
      | XY (Scatter) |
      | Bubble       |


  Scenario: Get Chart.series
    Given a chart
     Then chart.series is a SeriesCollection object
