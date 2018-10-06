Feature: Get and change line properties
  In order to format a shape outline and other line elements
  As a developer using python-pptx
  I need a set of read/write line properties on line elements


  Scenario: LineFormat.color
    Given a LineFormat object as line
     Then line.color is a ColorFormat object


  Scenario Outline: LineFormat.dash_style getter
    Given a LineFormat object as line having <current> dash style
     Then line.dash_style is <dash-style>

    Examples: Line dash styles
      | current     | dash-style        |
      | no explicit | None              |
      | solid       | MSO_LINE.SOLID    |
      | dashed      | MSO_LINE.DASH     |
      | dash-dot    | MSO_LINE.DASH_DOT |


  Scenario Outline: LineFormat.head_end getter
    Given a LineFormat object as line having <current> head end
     Then line.head_end is <head-end>
     Examples: Line head ends
      | current     | head-end                |
      | no explicit | None                    |
      | diamond     | MSO_ARROWHEAD.DIAMOND   |
      | triangle    | MSO_ARROWHEAD.TRIANGLE  |
      | oval        | MSO_ARROWHEAD.OVAL      |


  Scenario Outline: LineFormat.tail_end getter
    Given a LineFormat object as line having <current> tail end
     Then line.tail_end is <tail-end>
     Examples: Line tail ends
      | current     | tail-end                |
      | no explicit | None                    |
      | diamond     | MSO_ARROWHEAD.DIAMOND   |
      | triangle    | MSO_ARROWHEAD.TRIANGLE  |
      | oval        | MSO_ARROWHEAD.OVAL      |


  Scenario Outline: LineFormat.width setter
    Given a LineFormat object as line having <current> dash style
     When I assign <dash-style> to line.dash_style
     Then line.dash_style is <dash-style>

    Examples: Line dash style assignment scenarios
      | current     | dash-style        |
      | no explicit | MSO_LINE.DASH     |
      | dashed      | MSO_LINE.SOLID    |
      | solid       | None              |


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
