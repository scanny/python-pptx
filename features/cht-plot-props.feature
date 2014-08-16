Feature: Get and set plot properties
  In order to customize the formatting of a plot on a chart
  As a developer using python-pptx
  I need a way to get and set plot properties


  Scenario Outline: Determine whether a plot has data labels
    Given a bar plot <having-or-not> data labels
     Then the plot.has_data_labels property is <expected-value>

    Examples: having data labels or not
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


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


  Scenario Outline: Get bar plot gap width
    Given a bar plot having gap width of <gap-width>
     Then the value of plot.gap_width is <expected-value>

    Examples: bar plot gap widths
      | gap-width         | expected-value |
      | no explicit value | 150            |
      | 300               | 300            |


  Scenario Outline: Set bar plot gap width
    Given a bar plot having gap width of <gap-width>
     When I assign <new-value> to plot.gap_width
     Then the value of plot.gap_width is <expected-value>

    Examples: expected results of changing gap width
      | gap-width         | new-value | expected-value |
      | no explicit value | 300       | 300            |
      | 300               | 275       | 275            |


  Scenario: Get plot categories
    Given a bar plot having known categories
     Then plot.categories contains the known category strings
