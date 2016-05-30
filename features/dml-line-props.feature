Feature: Get and change line properties
  In order to format a shape outline and other line elements
  As a developer using python-pptx
  I need a set of read/write line properties on line elements


  Scenario Outline: Get line fill type
    Given an autoshape outline having <outline type>
     Then the line fill type is <line fill type>

    Examples: Line fill types
      | outline type                | line fill type           |
      | an inherited outline format | None                     |
      | no outline                  | MSO_FILL_TYPE.BACKGROUND |
      | a solid outline             | MSO_FILL_TYPE.SOLID      |


  Scenario Outline: Get line color type
    Given a line with <color type> color
     Then the line's color type is <value>

    Examples: Color type settings
      | color type | value       |
      | no         | None        |
      | an RGB     | RGB         |
      | a theme    | theme color |


  Scenario Outline: Get line width
    Given a line of <line width> width
     Then the reported line width is <reported line width>

    Examples: Line widths
      | line width  | reported line width |
      | no explicit | 0                   |
      | 1 pt        | 1 pt                |


  Scenario Outline: Set line fill type
    Given an autoshape outline having <outline type>
     When I set the line fill type to <fill type>
     Then the line fill type is <fill type index>

    Examples: Line fill types
      | outline type                | fill type  | fill type index          |
      | an inherited outline format | solid      | MSO_FILL_TYPE.SOLID      |
      | a solid outline             | background | MSO_FILL_TYPE.BACKGROUND |
      | no outline                  | solid      | MSO_FILL_TYPE.SOLID      |


  Scenario Outline: Set line color
    Given a line with no color
     When I set the line <color type> value
     Then the line's <color type> value matches the new value

    Examples: Color types
      | color type  |
      | RGB         |
      | theme color |


  Scenario Outline: Set line width
    Given a line of <line width> width
     When I set the line width to <new line width>
     Then the reported line width is <reported line width>

    Examples: Line widths
      | line width  | new line width | reported line width |
      | no explicit | None           | 0                   |
      | no explicit | 1 pt           | 1 pt                |
      | 1 pt        | None           | 0                   |
      | 1 pt        | 1 pt           | 1 pt                |
      | 1 pt        | 2.34 pt        | 2.34 pt             |
