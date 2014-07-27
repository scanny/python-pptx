Feature: Get and set chart properties
  In order to customize the formatting of a chart
  As a developer using python-pptx
  I need a way to get and set chart properties


  @wip
  Scenario Outline: Get chart type
    Given a chart of type <chart-type>
     Then chart.chart_type is <expected-enum-member>

    Examples: chart types
      | chart-type            | expected-enum-member     |
      | Area                  | AREA                     |
      | Stacked Area          | AREA_STACKED             |
      | 100% Stacked Area     | AREA_STACKED_100         |
      | 3-D Area              | THREE_D_AREA             |
      | 3-D Stacked Area      | THREE_D_AREA_STACKED     |
      | 3-D 100% Stacked Area | THREE_D_AREA_STACKED_100 |
