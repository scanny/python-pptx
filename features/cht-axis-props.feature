Feature: Get and set axis properties
  In order to customize the formatting of an axis on a chart
  As a developer using python-pptx
  I need a way to get and set axis properties


  Scenario Outline: Determine whether an axis has major and/or minor gridlines
    Given an axis <having-or-not> <major-or-minor> gridlines
     Then axis.has_<major-or-minor>_gridlines is <expected-value>

    Examples: having gridlines or not
      | having-or-not | major-or-minor | expected-value |
      | having        | major          | True           |
      | not having    | major          | False          |
      | having        | minor          | True           |
      | not having    | minor          | False          |


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


  Scenario Outline: Determine axis major and minor unit
    Given an axis having <major-or-minor> unit of <value>
     Then axis.<major-or-minor>_unit is <expected-value>

    Examples: axis unit values
      | major-or-minor | value | expected-value |
      | major          | 20.0  | 20.0           |
      | major          | Auto  | None           |
      | minor          | 4.2   | 4.2            |
      | minor          | Auto  | None           |


  Scenario Outline: Change axis major or minor unit
    Given an axis having <major-or-minor> unit of <value>
     When I assign <new-value> to axis.<major-or-minor>_unit
     Then axis.<major-or-minor>_unit is <expected-value>

    Examples: expected results of changing <major-or-minor>_unit
      | major-or-minor | value | new-value | expected-value |
      | major          | 20.0  | 5         | 5.0            |
      | major          | 20.0  | None      | None           |
      | major          | Auto  | 5         | 5.0            |
      | major          | Auto  | None      | None           |
      | minor          | 4.2   | 8.4       | 8.4            |
      | minor          | 4.2   | None      | None           |
      | minor          | Auto  | 8.4       | 8.4            |
      | minor          | Auto  | None      | None           |
