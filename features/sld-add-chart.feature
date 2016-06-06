Feature: Add a chart
  In order to include custom charts in a presentation
  As a developer using python-pptx
  I need a way to add a chart to a slide, specifying its type and data


  Scenario Outline: Add a category chart
    Given a blank slide
     When I add a <type> chart with <cats> categories and <sers> series
     Then chart.chart_type is <chart-type>
      And len(plot.categories) is <cats>
      And len(chart.series) is <sers>
      And len(series.values) is <cats> for each series
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | type             | chart-type       | cats | sers |
      | Clustered Bar    | BAR_CLUSTERED    |   3  |   3  |
      | 100% Stacked Bar | BAR_STACKED_100  |   3  |   3  |
      | Clustered Column | COLUMN_CLUSTERED |   3  |   3  |
      | Line             | LINE             |   3  |   2  |
      | Pie              | PIE              |   5  |   1  |
      | Radar            | RADAR            |   5  |   2  |


  Scenario Outline: Add an XY chart
    Given a blank slide
     When I add an <chart-type> chart having 2 series of 3 points each
     Then chart.chart_type is <chart-type>
      And len(chart.series) is 2
      And len(series.values) is 3 for each series
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | chart-type                   |
      | XY_SCATTER                   |
      | XY_SCATTER_LINES             |
      | XY_SCATTER_LINES_NO_MARKERS  |
      | XY_SCATTER_SMOOTH            |
      | XY_SCATTER_SMOOTH_NO_MARKERS |


  Scenario Outline: Add a bubble chart
    Given a blank slide
     When I add a <chart-type> chart having 2 series of 3 points each
     Then chart.chart_type is <chart-type>
      And len(chart.series) is 2
      And len(series.values) is 3 for each series
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | chart-type            |
      | BUBBLE                |
      | BUBBLE_THREE_D_EFFECT |
