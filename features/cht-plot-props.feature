Feature: Plot properties
  To customize the formatting of a plot on a chart
  As a developer using python-pptx
  I need read/write properties on plot objects


  Scenario Outline: Get bubble_plot.bubble_scale
    Given a bubble plot having bubble scale of <percent>
     Then bubble_plot.bubble_scale is <value>

    Examples: bubble_plot.bubble_scale values
      | percent           | value |
      | no explicit value |  100  |
      | 70%               |   70  |


  Scenario Outline: Set bubble_plot.bubble_scale
    Given a bubble plot having bubble scale of <percent>
     When I assign <new-value> to bubble_plot.bubble_scale
     Then bubble_plot.bubble_scale is <value>

    Examples: bubble_plot.bubble_scale assignment cases
      | percent           | new-value | value |
      | no explicit value | 70        |   70  |
      | 70%               | 150       |  150  |
      | 70%               | None      |  100  |
      | no explicit value | None      |  100  |


  Scenario: Get category_plot.categories
    Given a category plot
     Then plot.categories is a Categories object


  Scenario Outline: Get plot.has_data_labels
    Given a bar plot <having-or-not> data labels
     Then plot.has_data_labels is <expected-value>

    Examples: having data labels or not
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


  Scenario Outline: Set plot.has_data_labels
    Given a bar plot <having-or-not> data labels
     When I assign <value> to plot.has_data_labels
     Then plot.has_data_labels is <expected-value>

    Examples: expected results of changing has_data_labels
      | having-or-not | value | expected-value |
      | having        | False | False          |
      | having        | True  | True           |
      | not having    | False | False          |
      | not having    | True  | True           |


  Scenario Outline: Get bar_plot.gap_width
    Given a bar plot having gap width of <gap-width>
     Then plot.gap_width is <expected-value>

    Examples: bar plot gap widths
      | gap-width         | expected-value |
      | no explicit value | 150            |
      | 300               | 300            |


  Scenario Outline: Set bar_plot.gap_width
    Given a bar plot having gap width of <gap-width>
     When I assign <new-value> to plot.gap_width
     Then plot.gap_width is <expected-value>

    Examples: expected results of changing gap width
      | gap-width         | new-value | expected-value |
      | no explicit value | 300       | 300            |
      | 300               | 275       | 275            |


  Scenario Outline: Get bar_plot.overlap
    Given a bar plot having overlap of <overlap>
     Then plot.overlap is <expected-value>

    Examples: expected bar plot overlap values
      | overlap           | expected-value |
      | no explicit value | 0              |
      | 42                | 42             |
      | -42               | -42            |


  Scenario Outline: Set bar_plot.overlap
    Given a bar plot having overlap of <overlap>
     When I assign <new-value> to plot.overlap
     Then plot.overlap is <expected-value>

    Examples: expected results of changing overlap
      | overlap           | new-value | expected-value |
      | no explicit value | 42        | 42             |
      | 42                | -42       | -42            |


  Scenario Outline: Get plot.vary_by_categories
    Given a bar plot having vary color by category set to <setting>
     Then plot.vary_by_categories is <expected-value>

    Examples: expected Plot.vary_by_categories values
      | setting             | expected-value |
      | no explicit setting | True           |
      | True                | True           |
      | False               | False          |


  Scenario Outline: Set plot.vary_by_categories
    Given a bar plot having vary color by category set to <setting>
     When I assign <value> to plot.vary_by_categories
     Then plot.vary_by_categories is <expected-value>

    Examples: expected Plot.vary_by_categories values
      | setting             | value | expected-value |
      | no explicit setting | False | False          |
      | True                | False | False          |
      | False               | True  | True           |
