Feature: Get and set plot properties
  In order to customize the formatting of a plot on a chart
  As a developer using python-pptx
  I need a way to get and set plot properties


  Scenario Outline: Determine whether a plot has data labels
    Given a bar plot <having-or-not> data labels
     Then plot.has_data_labels is <expected-value>

    Examples: having data labels or not
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


  Scenario Outline: Change whether a plot has data labels
    Given a bar plot <having-or-not> data labels
     When I assign <value> to plot.has_data_labels
     Then plot.has_data_labels is <expected-value>

    Examples: expected results of changing has_data_labels
      | having-or-not | value | expected-value |
      | having        | False | False          |
      | having        | True  | True           |
      | not having    | False | False          |
      | not having    | True  | True           |


  Scenario Outline: Get bar plot gap width
    Given a bar plot having gap width of <gap-width>
     Then plot.gap_width is <expected-value>

    Examples: bar plot gap widths
      | gap-width         | expected-value |
      | no explicit value | 150            |
      | 300               | 300            |


  Scenario Outline: Set bar plot gap width
    Given a bar plot having gap width of <gap-width>
     When I assign <new-value> to plot.gap_width
     Then plot.gap_width is <expected-value>

    Examples: expected results of changing gap width
      | gap-width         | new-value | expected-value |
      | no explicit value | 300       | 300            |
      | 300               | 275       | 275            |


  Scenario Outline: Get bar plot overlap
    Given a bar plot having overlap of <overlap>
     Then plot.overlap is <expected-value>

    Examples: expected bar plot overlap values
      | overlap           | expected-value |
      | no explicit value | 0              |
      | 42                | 42             |
      | -42               | -42            |


  Scenario Outline: Set bar plot overlap
    Given a bar plot having overlap of <overlap>
     When I assign <new-value> to plot.overlap
     Then plot.overlap is <expected-value>

    Examples: expected results of changing overlap
      | overlap           | new-value | expected-value |
      | no explicit value | 42        | 42             |
      | 42                | -42       | -42            |


  Scenario: Get plot categories
    Given a bar plot having known categories
     Then plot.categories contains the known category strings


  Scenario Outline: Determine whether a plot varies color by category
    Given a bar plot having vary color by category set to <setting>
     Then plot.vary_by_categories is <expected-value>

    Examples: expected Plot.vary_by_categories values
      | setting             | expected-value |
      | no explicit setting | True           |
      | True                | True           |
      | False               | False          |


  Scenario Outline: Change whether a plot varies color by category
    Given a bar plot having vary color by category set to <setting>
     When I assign <value> to plot.vary_by_categories
     Then plot.vary_by_categories is <expected-value>

    Examples: expected Plot.vary_by_categories values
      | setting             | value | expected-value |
      | no explicit setting | False | False          |
      | True                | False | False          |
      | False               | True  | True           |
