Feature: Replace chart data
  In order to update the data for a chart while retaining its formatting
  As a developer using python-pptx
  I need a way to replace the data of a chart


  Scenario Outline: Replace category chart data
    Given a chart of size and type <spec>
     When I replace its data with <cats> categories and <sers> series
     Then len(plot.categories) is <cats>
      And len(chart.series) is <sers>
      And len(series.values) is <cats> for each series
      And each series has a new name
      And the chart has new chart data

    Examples: Replacement chart data
      | spec                 | cats | sers |
      | 2x2 Clustered Bar    |   3  |   3  |
      | 2x2 100% Stacked Bar |   3  |   3  |
      | 2x2 Clustered Column |   3  |   3  |
      | 4x3 Line             |   3  |   2  |
      | 3x1 Pie              |   5  |   1  |


  Scenario: Replace XY chart data
    Given a chart of size and type 3x2 XY
     When I replace its data with 3 series of 3 points each
     Then len(chart.series) is 3
      And len(series.values) is 3 for each series
      And each series has a new name
      And the chart has new chart data


  Scenario: Replace Bubble chart data
    Given a chart of size and type 3x2 Bubble
     When I replace its data with 3 series of 3 bubble points each
     Then len(chart.series) is 3
      And len(series.values) is 3 for each series
      And each series has a new name
      And the chart has new chart data


  Scenario: Cloned series use rotating theme accent colors (issue #529)
    Given a chart with an explicitly-colored series
     When I replace its data with 6 series that require 5 new cloned series
     Then each cloned series uses a distinct theme-accent schemeClr


  Scenario Outline: Chart.replace_data rejects mismatched ChartData type (issue #396)
    Given a chart of size and type <spec>
     Then replacing its data with <wrong-kind> raises ValueError

    Examples: incompatible chart-type / ChartData combinations
      | spec                 | wrong-kind     |
      | 3x2 XY               | CategoryData   |
      | 3x2 XY               | BubbleData     |
      | 3x2 Bubble           | CategoryData   |
      | 3x2 Bubble           | XyData         |
      | 2x2 Clustered Bar    | XyData         |
      | 4x3 Line             | BubbleData     |
