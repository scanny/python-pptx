Feature: Chart Title properties
  In order to customize the formatting of a chart title
  As a developer using python-pptx
  I need read/write properties on ChartTitle


  Scenario: ChartTitle.format
    Given a chart title
     Then chart_title.format is a ChartFormat object
      And chart_title.format.fill is a FillFormat object
      And chart_title.format.line is a LineFormat object


  Scenario Outline: Get ChartTitle.has_text_frame
    Given a chart title having <a-or-no> text frame
     Then chart_title.has_text_frame is <value>

    Examples: text frame presence cases
      | a-or-no | value |
      | a       | True  |
      | no      | False |


  Scenario Outline: Set ChartTitle.has_text_frame
    Given a chart title having <a-or-no> text frame
     When I assign <new-value> to chart_title.has_text_frame
     Then chart_title.has_text_frame is <value>

    Examples: chart_title.has_text_frame assignment cases
      | a-or-no | new-value | value |
      | no      | True      | True  |
      | a       | False     | False |
      | no      | False     | False |
      | a       | True      | True  |


  Scenario Outline: Get ChartTitle.text_frame
    Given a chart title having <a-or-no> text frame
     Then chart_title.text_frame is a TextFrame object

    Examples: text frame presence cases
      | a-or-no |
      | a       |
      | no      |
