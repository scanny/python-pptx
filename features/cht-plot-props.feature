Feature: Get and set plot properties
  In order to customize the formatting of a plot on a chart
  As a developer using python-pptx
  I need a way to get and set plot properties


  @wip
  Scenario Outline: Determine whether a plot has data labels
    Given a bar plot <having-or-not> data labels
     Then the plot.has_data_labels property is <expected-value>

    Examples: having data labels or not
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


  @wip
  Scenario Outline: Change whether a plot has data labels
    Given a bar plot <having-or-not> data labels
     When I assign <value> to plot.has_data_labels
     Then the plot.has_data_labels property is <expected-value>

    Examples: expected results of changing has_data_labels
      | having-or-not | value | expected-value |
      | having        | False | False          |
      | having        | True  | True           |
      | not having    | False | False          |
      | not having    | True  | True           |
