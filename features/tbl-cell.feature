Feature: Table cell proxy objects
  In order to change the formatting of a cell in a table to my needs
  As a developer using python-pptx
  I need properties and methods on cell objects


  Scenario: _Cell.fill
    Given a _Cell object as cell
     Then cell.fill is a FillFormat object


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


  Scenario: _Cell.text setter
    Given a _Cell object as cell
     When I assign cell.text = "test text"
     Then cell.text_frame.text == "test text"


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
