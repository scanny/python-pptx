Feature: Get and set chart properties
  In order to customize the formatting of a chart
  As a developer using python-pptx
  I need a way to get and set chart properties


  @wip
  Scenario Outline: Get chart type
    Given a chart of type <chart-type>
     Then chart.chart_type is <expected-enum-member>

    Examples: chart types
      | chart-type        | expected-enum-member |
      | area              | AREA                 |
      | stacked area      | AREA_STACKED         |
      | 100% stacked area | AREA_STACKED_100     |
