Feature: Access chart legend
  In order to customize the legend for a chart
  As a developer using python-pptx
  I need a way to access the chart legend

  Scenario Outline: Determine presence of legend
    Given a chart <having-or-not> a legend
     Then chart.has_legend is <expected-value>

    Examples: chart.has_legend states
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


  Scenario Outline: Adding a legend
    Given a chart <having-or-not> a legend
     When I assign <value> to chart.has_legend
     Then chart.has_legend is <expected-value>

    Examples: chart.has_legend states
      | having-or-not | value | expected-value |
      | having        | False | False          |
      | not having    | True  | True           |
      | having        | True  | True           |
      | not having    | False | False          |


  Scenario: Access legend object of chart
    Given a chart having a legend
     Then chart.legend is a legend object
