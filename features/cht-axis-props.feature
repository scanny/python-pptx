Feature: Get and set axis properties
  In order to customize the formatting of an axis on a chart
  As a developer using python-pptx
  I need a way to get and set axis properties


  @wip
  Scenario Outline: Determine whether an axis has major and/or minor gridlines
    Given an axis <having-or-not> <major-or-minor> gridlines
     Then axis.has_<major-or-minor>_gridlines is <expected-value>

    Examples: having gridlines or not
      | having-or-not | major-or-minor | expected-value |
      | having        | major          | True           |
      | not having    | major          | False          |
      | having        | minor          | True           |
      | not having    | minor          | False          |


  @wip
  Scenario Outline: Change whether an axis has major and/or minor gridlines
    Given an axis <having-or-not> <major-or-minor> gridlines
     When I assign <value> to axis.has_<major-or-minor>_gridlines
     Then axis.has_<major-or-minor>_gridlines is <expected-value>

    Examples: expected results of changing has_<major-or-minor>_gridlines
      | having-or-not | major-or-minor | value | expected-value |
      | having        | major          | False | False          |
      | having        | major          | True  | True           |
      | not having    | minor          | False | False          |
      | not having    | minor          | True  | True           |
