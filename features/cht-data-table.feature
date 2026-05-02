Feature: Chart data-table
  In order to display the tabular source values beneath a chart
  As a developer using python-pptx
  I need read/write access to a chart's data-table properties


  Scenario Outline: Determine presence of data-table
    Given a chart <having-or-not> a data-table
     Then chart.has_data_table is <expected-value>
     And chart.data_table is <data-table-value>

    Examples: chart.has_data_table states
      | having-or-not | expected-value | data-table-value  |
      | having        | True           | a _DataTable      |
      | not having    | False          | None              |


  Scenario Outline: Add or remove a data-table
    Given a chart <having-or-not> a data-table
     When I assign <value> to chart.has_data_table
     Then chart.has_data_table is <expected-value>

    Examples: chart.has_data_table assignment
      | having-or-not | value | expected-value |
      | not having    | True  | True           |
      | having        | True  | True           |
      | having        | False | False          |
      | not having    | False | False          |


  Scenario: Default data-table flags are all on
    Given a chart not having a data-table
     When I assign True to chart.has_data_table
     Then chart.data_table.show_horz_border is True
      And chart.data_table.show_vert_border is True
      And chart.data_table.show_outline is True
      And chart.data_table.show_keys is True


  Scenario Outline: Toggle individual data-table flags
    Given a chart having a data-table
     When I assign <value> to chart.data_table.<flag>
     Then chart.data_table.<flag> is <expected-value>

    Examples: show-* flag read/write
      | flag              | value | expected-value |
      | show_horz_border  | False | False          |
      | show_vert_border  | False | False          |
      | show_outline      | False | False          |
      | show_keys         | False | False          |


  Scenario: Access data-table formatting
    Given a chart having a data-table
     Then chart.data_table.format is a ChartFormat object
