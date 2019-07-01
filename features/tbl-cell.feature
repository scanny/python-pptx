Feature: Table cell proxy objects
  In order to change the formatting of a cell in a table to my needs
  As a developer using python-pptx
  I need properties and methods on cell objects


  Scenario Outline: Cell objects proxying same a:tc element compare equal
    Given a <role> _Cell object as cell
      And a second proxy instance for that cell as other_cell
     Then cell == other_cell

    Examples: merged cell roles
      | role         |
      | merge-origin |
      | spanned      |
      | unmerged     |


  Scenario: _Cell.fill
    Given a _Cell object as cell
     Then cell.fill is a FillFormat object


  Scenario Outline: Cell role discovery
    Given a <role> _Cell object as cell
     Then cell.is_merge_origin is <is_merge_origin>
      And cell.is_spanned is <is_spanned>

    Examples: merged cell roles
      | role         | is_merge_origin | is_spanned |
      | merge-origin | True            | False      |
      | spanned      | False           | True       |
      | unmerged     | False           | False      |


  Scenario: _Cell.margin_{x} getters
    Given a _Cell object with known margins as cell
     Then cell.margin_left == Inches(0.2)
      And cell.margin_top == Inches(0.3)
      And cell.margin_right == Inches(0.4)
      And cell.margin_bottom == Inches(0.5)


  Scenario Outline: _Cell.margin_{x} setters
    Given a _Cell object as cell
     When I assign cell.margin_<side> = <value>
     Then cell.margin_<side> == <new_value>

    Examples: Cell margin assignment cases
      | side   | value       | new_value    |
      | left   | Inches(0.2) | Inches(0.2)  |
      | top    | None        | Inches(0.05) |
      | right  | None        | Inches(0.1)  |
      | bottom | Inches(0.3) | Inches(0.3)  |


  Scenario: _Cell.merge()
    Given a 3x3 Table object with cells a to i as table
     When I assign origin_cell = table.cell(0, 0)
      And I assign other_cell = table.cell(1, 1)
      And I call origin_cell.merge(other_cell)
     Then origin_cell.is_merge_origin is True
      And other_cell.is_spanned is True
      And origin_cell.text == "a\nb\nd\ne"
      And other_cell.text == ""


  Scenario: Merged cell size
    Given a 2x3 _MergeOriginCell object as cell
     Then cell.span_height == 2
      And cell.span_width == 3


  Scenario: _Cell.split()
    Given a merge-origin _Cell object as cell
     When I call cell.split()
     Then cell.is_merge_origin is False
      And cell.span_height == 1
      And cell.span_width == 1


  Scenario: _Cell.text getter
    Given a _Cell object containing "unladen swallows" as cell
     Then cell.text == "unladen swallows"


  Scenario: _Cell.text setter
    Given a _Cell object as cell
     When I assign cell.text = "test text"
     Then cell.text == "test text"


  Scenario Outline: _Cell.vertical_anchor getter
    Given a _Cell object with <setting> vertical alignment as cell
     Then cell.vertical_anchor == <value>

    Examples: Cell margin assignment cases
      | setting   | value             |
      | inherited | None              |
      | middle    | MSO_ANCHOR.MIDDLE |
      | bottom    | MSO_ANCHOR.BOTTOM |


  Scenario Outline: _Cell.vertical_anchor setter
    Given a _Cell object with <setting> vertical alignment as cell
     When I assign cell.vertical_anchor = <value>
     Then cell.vertical_anchor == <value>

    Examples: Cell margin assignment cases
      | setting   | value             |
      | inherited | MSO_ANCHOR.TOP    |
      | middle    | MSO_ANCHOR.BOTTOM |
      | bottom    | None              |
