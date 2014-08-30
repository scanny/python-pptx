Feature: Add a chart
  In order to include custom charts in a presentation
  As a developer using python-pptx
  I need a way to add a chart to a slide, specifying its type and data

  Scenario Outline: Add a chart
    Given a blank slide
     When I add a <type> chart with <cats> categories and <sers> series
     Then the chart type is <chart-type>
      And the chart has <cats> categories
      And the chart has <sers> series
      And each series has <cats> values
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | type             | chart-type       | cats | sers |
      | Clustered Bar    | BAR_CLUSTERED    |   3  |   3  |
      | 100% Stacked Bar | BAR_STACKED_100  |   3  |   3  |
      | Clustered Column | COLUMN_CLUSTERED |   3  |   3  |
      | Line             | LINE             |   3  |   2  |
      | Pie              | PIE              |   5  |   1  |
