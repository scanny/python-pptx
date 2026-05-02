Feature: Chart properties
  In order to customize the formatting of a chart
  As a developer using python-pptx
  I need read/write properties and methods on Chart


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


  Scenario: Chart.font
    Given a Chart object as chart
     Then chart.font is a Font object


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


  Scenario: Chart.series
    Given a Chart object as chart
     Then chart.series is a SeriesCollection object


  Scenario Outline: Chart.has_secondary_value_axis
    Given a <axis-config>-axes chart
     Then chart.has_secondary_value_axis is <expected-value>

    Examples: secondary value axis presence cases
      | axis-config       | expected-value |
      | single-value      | False          |
      | primary-secondary | True           |


  Scenario: Chart.secondary_value_axis
    Given a primary-secondary-axes chart
     Then chart.secondary_value_axis is a ValueAxis object


  Scenario: Chart.secondary_value_axis raises when not present
    Given a single-value-axes chart
     Then accessing chart.secondary_value_axis raises ValueError
  Scenario: Chart.plot_area (issue #298)
    Given a Chart object as chart
     Then chart.plot_area is a PlotArea object
      And chart.plot_area.format is a ChartFormat object
      And chart.plot_area.format.fill is a FillFormat object
      And chart.plot_area.format.line is a LineFormat object


  Scenario: Chart.chart_style reads extended c14:style value (issue #516)
    Given a chart with an extended c14:style wrapped in mc:AlternateContent
     Then chart.chart_style is 118


  Scenario: Chart.chart_style writes an mc:AlternateContent wrapper for extended values
    Given a chart with no explicit chart style
     When I assign 118 to chart.chart_style
     Then chart.chart_style is 118
      And chartSpace has an mc:AlternateContent/mc:Choice/c14:style val=118
      And chartSpace has an mc:AlternateContent/mc:Fallback/c:style val=18


  Scenario: Chart.display_blanks_as reads the template default (issue #859)
    Given a chart with no explicit chart style
     Then chart.display_blanks_as is XL_DISPLAY_BLANKS_AS.GAPS


  Scenario Outline: Set Chart.display_blanks_as
    Given a chart with no explicit chart style
     When I assign XL_DISPLAY_BLANKS_AS.<member> to chart.display_blanks_as
     Then chart.display_blanks_as is XL_DISPLAY_BLANKS_AS.<member>

    Examples: display-blanks-as members
      | member       |
      | GAPS         |
      | ZERO         |
      | INTERPOLATED |
