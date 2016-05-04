Feature: Access and modify data labels properties
  In order to customize the appearance of data labels on a chart
  As a developer using python-pptx
  I need a way to get and set data labels properties


  Scenario Outline: Determine position of data labels
    Given bar chart data labels positioned <relation-to> their data point
     Then data_labels.position is <expected-value>

    Examples: data_labels position values
      | relation-to                | expected-value |
      | in unspecified relation to | None           |
      | inside, at the base of     | INSIDE_BASE    |


  Scenario Outline: Change position of data labels
    Given bar chart data labels positioned <relation-to> their data point
     When I assign <new-value> to data_labels.position
     Then data_labels.position is <expected-value>

    Examples: expected results of assignment to data_labels.position
      | relation-to                | new-value   | expected-value |
      | in unspecified relation to | INSIDE_END  | INSIDE_END     |
      | inside, at the base of     | OUTSIDE_END | OUTSIDE_END    |
      | inside, at the base of     | None        | None           |
