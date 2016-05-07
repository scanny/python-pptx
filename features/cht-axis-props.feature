Feature: Axis properties
  In order to customize the formatting of an axis on a chart
  As a developer using python-pptx
  I need read/write properties on Axis


  Scenario Outline: Get Axis.has_[major/minor]_gridlines
    Given an axis <having-or-not> <major-or-minor> gridlines
     Then axis.has_<major-or-minor>_gridlines is <expected-value>

    Examples: gridlines presence cases
      | having-or-not | major-or-minor | expected-value |
      | having        | major          | True           |
      | not having    | major          | False          |
      | having        | minor          | True           |
      | not having    | minor          | False          |


  Scenario Outline: Set Axis.has_[major/minor]_gridlines
    Given an axis <having-or-not> <major-or-minor> gridlines
     When I assign <value> to axis.has_<major-or-minor>_gridlines
     Then axis.has_<major-or-minor>_gridlines is <expected-value>

    Examples: has_major/minor_gridlines assignment cases
      | having-or-not | major-or-minor | value | expected-value |
      | having        | major          | False | False          |
      | having        | major          | True  | True           |
      | not having    | minor          | False | False          |
      | not having    | minor          | True  | True           |


  Scenario Outline: Get Axis.major/minor_unit
    Given an axis having <major-or-minor> unit of <value>
     Then axis.<major-or-minor>_unit is <expected-value>

    Examples: axis unit cases
      | major-or-minor | value | expected-value |
      | major          | 20.0  | 20.0           |
      | major          | Auto  | None           |
      | minor          | 4.2   | 4.2            |
      | minor          | Auto  | None           |


  Scenario Outline: Set Axis.major/minor_unit
    Given an axis having <major-or-minor> unit of <value>
     When I assign <new-value> to axis.<major-or-minor>_unit
     Then axis.<major-or-minor>_unit is <expected-value>

    Examples: major/minor_unit assignment cases
      | major-or-minor | value | new-value | expected-value |
      | major          | 20.0  | 5         | 5.0            |
      | major          | 20.0  | None      | None           |
      | major          | Auto  | 5         | 5.0            |
      | major          | Auto  | None      | None           |
      | minor          | 4.2   | 8.4       | 8.4            |
      | minor          | 4.2   | None      | None           |
      | minor          | Auto  | 8.4       | 8.4            |
      | minor          | Auto  | None      | None           |


  Scenario: Get Axis.major_gridlines
    Given an axis
     Then axis.major_gridlines is a MajorGridlines object


  @wip
  Scenario Outline: Get Axis.format
    Given a <axis-type> axis
     Then axis.format is a ChartFormat object
      And axis.format.fill is a FillFormat object
      And axis.format.line is a LineFormat object

    Examples: axis types
      | axis-type |
      | category  |
      | value     |
