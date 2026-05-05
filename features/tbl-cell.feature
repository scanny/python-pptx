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


  Scenario Outline: _Cell.border_{side} is a LineFormat object
    Given a _Cell object as cell
     Then cell.<border> is a LineFormat object

    Examples: Cell border sides
      | border                |
      | border_left           |
      | border_right          |
      | border_top            |
      | border_bottom         |
      | border_diagonal_down  |
      | border_diagonal_up    |


  Scenario: _Cell.border_{side} round-trips a color, width, and dash
    Given a _Cell object as cell
     When I set cell.border_left to a red, 1.5pt, dashed line
     Then cell.border_left.color.rgb is RGBColor(0xFF, 0x00, 0x00)
      And cell.border_left.width == Pt(1.5)
      And cell.border_left.dash_style == MSO_LINE.DASH


  Scenario Outline: Cell role discovery
    Given a <role> _Cell object as cell
     Then cell.is_merge_origin is <is_merge_origin>
      And cell.is_spanned is <is_spanned>

    Examples: merged cell roles
      | role         | is_merge_origin | is_spanned |
      | merge-origin | True            | False      |
      | spanned      | False           | True       |
      | unmerged     | False           | False      |


  Scenario: _Cell.row_idx and _Cell.col_idx
    Given a 3x3 Table object as table
     Then table.cell(0, 0).row_idx == 0
      And table.cell(0, 0).col_idx == 0
      And table.cell(1, 2).row_idx == 1
      And table.cell(1, 2).col_idx == 2
      And table.cell(2, 1).row_idx == 2
      And table.cell(2, 1).col_idx == 1


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


  Scenario: _Cell.merge() in single-column table deletes merged rows (#636)
    Given a freshly-added 5x1 Table object as table
     When I assign origin_cell = table.cell(2, 0)
      And I assign other_cell = table.cell(3, 0)
      And I call origin_cell.merge(other_cell)
     Then len(table.rows) == 4
      And origin_cell.is_merge_origin is False
      And origin_cell.span_height == 1


  Scenario: _Cell.merge() in single-row table deletes merged columns (#636)
    Given a freshly-added 1x5 Table object as table
     When I assign origin_cell = table.cell(0, 2)
      And I assign other_cell = table.cell(0, 3)
      And I call origin_cell.merge(other_cell)
     Then len(table.columns) == 4
      And origin_cell.is_merge_origin is False
      And origin_cell.span_width == 1


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


  Scenario: _Cell.clone_from() copies visuals and content
    Given a fully-styled source _Cell and a plain target _Cell
     When I call target_cell.clone_from(source_cell)
     Then target_cell.text == "styled"
      And target_cell.fill.fore_color.rgb is RGBColor(0xFF, 0x00, 0x00)
      And target_cell.border_left.color.rgb is RGBColor(0x00, 0x00, 0xFF)
      And target_cell.border_left.width == Pt(1.5)
      And target_cell.border_diagonal_down.color.rgb is RGBColor(0x00, 0xFF, 0x00)
      And target_cell.margin_left == Inches(0.2)
      And target_cell.vertical_anchor == MSO_ANCHOR.MIDDLE
      And target_cell.clone_from returned target_cell
      And source_cell was not mutated by the clone
