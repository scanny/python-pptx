Feature: Axis Title properties
  In order to customize the formatting of an axis title
  As a developer using python-pptx
  I need read/write properties on AxisTitle


  Scenario: AxisTitle.format
    Given an axis title
     Then axis_title.format is a ChartFormat object
      And axis_title.format.fill is a FillFormat object
      And axis_title.format.line is a LineFormat object


  Scenario Outline: Get AxisTitle.has_text_frame
    Given an axis title having <a-or-no> text frame
     Then axis_title.has_text_frame is <value>

    Examples: text frame presence cases
      | a-or-no | value |
      | a       | True  |
      | no      | False |


  Scenario Outline: Set AxisTitle.has_text_frame
    Given an axis title having <a-or-no> text frame
     When I assign <new-value> to axis_title.has_text_frame
     Then axis_title.has_text_frame is <value>

    Examples: axis_title.has_text_frame assignment cases
      | a-or-no | new-value | value |
      | no      | True      | True  |
      | a       | False     | False |
      | no      | False     | False |
      | a       | True      | True  |


  Scenario Outline: Get AxisTitle.text_frame
    Given an axis title having <a-or-no> text frame
     Then axis_title.text_frame is a TextFrame object

    Examples: text frame presence cases
      | a-or-no |
      | a       |
      | no      |
