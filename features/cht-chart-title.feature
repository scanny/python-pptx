Feature: Access chart title
  In order to customize the title for a chart
  As a developer using python-pptx
  I need a way to access the chart title

  Scenario Outline: Determine presence of title
    Given a chart <having-or-not> a title
     Then chart.has_title is <expected-value>

    Examples: chart.has_title states
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


  Scenario Outline: Adding a title
    Given a chart <having-or-not> a title
     When I assign <value> to chart.has_title
     Then chart.has_title is <expected-value>

    Examples: chart.has_title states
      | having-or-not | value | expected-value |
      | having        | False | False          |
      | not having    | True  | True           |
      | having        | True  | True           |
      | not having    | False | False          |


  Scenario: Access title object of chart
    Given a chart having a title
     Then chart.title is a title object
