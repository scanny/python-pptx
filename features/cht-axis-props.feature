Feature: Axis properties
  In order to customize the formatting of an axis on a chart
  As a developer using python-pptx
  I need read/write properties on Axis


  Scenario: Get Axis.axis_title
    Given an axis
     Then axis.axis_title is an AxisTitle object


  Scenario Outline: Get Axis.category_type
    Given an axis of type <axis-type>
     Then axis.category_type is XL_CATEGORY_TYPE.<member>

    Examples: axis category type cases
      | axis-type    | member         |
      | CategoryAxis | CATEGORY_SCALE |
      | DateAxis     | TIME_SCALE     |


  Scenario Outline: Get value_axis.crosses
    Given a value axis having category axis crossing of <crossing>
     Then value_axis.crosses is <member>

    Examples: value_axis.crosses cases
      | crossing  | member    |
      | automatic | AUTOMATIC |
      | maximum   | MAXIMUM   |
      | minimum   | MINIMUM   |
      | 2.75      | CUSTOM    |


  Scenario Outline: Set value_axis.crosses
    Given a value axis having category axis crossing of <crossing>
     When I assign <member> to value_axis.crosses
     Then value_axis.crosses is <member>

    Examples: value_axis.crosses assignment cases
      | crossing  | member    |
      | automatic | MAXIMUM   |
      | maximum   | MINIMUM   |
      | minimum   | CUSTOM    |
      | 2.75      | CUSTOM    |


  Scenario Outline: Get value_axis.crosses_at
    Given a value axis having category axis crossing of <crossing>
     Then value_axis.crosses_at is <value>

    Examples: value_axis.crosses_at cases
      | crossing  | value |
      | automatic | None  |
      | maximum   | None  |
      | minimum   | None  |
      | 2.75      | 2.75  |
      | -1.5      | -1.5  |


  Scenario Outline: Set value_axis.crosses_at
    Given a value axis having category axis crossing of <crossing>
     When I assign <value> to value_axis.crosses_at
     Then value_axis.crosses_at is <value>

    Examples: value_axis.crosses_at assignment cases
      | crossing  | value |
      | automatic | 2.75  |
      | 2.75      | -1.5  |
      | -1.5      | None  |
      | automatic | None  |


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


  Scenario Outline: Get Axis.has_title
    Given an axis having <a-or-no> title
     Then axis.has_title is <expected-value>

    Examples: axis title presence cases
      | a-or-no | expected-value |
      | a       | True           |
      | no      | False          |


  Scenario Outline: Set Axis.has_title
    Given an axis having <a-or-no> title
     When I assign <value> to axis.has_title
     Then axis.has_title is <expected-value>

    Examples: axis title assignment cases
      | a-or-no | value | expected-value |
      | a       | True  | True           |
      | a       | False | False          |
      | no      | True  | True           |
      | no      | False | False          |


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


  Scenario Outline: Get Axis.format
    Given a <axis-type> axis
     Then axis.format is a ChartFormat object
      And axis.format.fill is a FillFormat object
      And axis.format.line is a LineFormat object

    Examples: axis types
      | axis-type |
      | category  |
      | value     |
