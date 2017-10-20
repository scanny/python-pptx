Feature: Get and change line properties
  In order to format a shape outline and other line elements
  As a developer using python-pptx
  I need a set of read/write line properties on line elements


  Scenario: LineFormat.color
    Given a LineFormat object as line
     Then line.color is a ColorFormat object


  Scenario: LineFormat.fill
    Given a LineFormat object as line
     Then line.fill is a FillFormat object


  Scenario Outline: LineFormat.width getter
    Given a LineFormat object as line having <line width> width
     Then line.width is <reported line width>

    Examples: Line widths
      | line width  | reported line width |
      | no explicit | 0                   |
      | 1 pt        | 1 pt                |


  Scenario Outline: LineFormat.width setter
    Given a LineFormat object as line having <line width> width
     When I assign <new line width> to line.width
     Then line.width is <reported line width>

    Examples: Line widths
      | line width  | new line width | reported line width |
      | no explicit | None           | 0                   |
      | no explicit | 1 pt           | 1 pt                |
      | 1 pt        | None           | 0                   |
      | 1 pt        | 1 pt           | 1 pt                |
      | 1 pt        | 2.34 pt        | 2.34 pt             |
