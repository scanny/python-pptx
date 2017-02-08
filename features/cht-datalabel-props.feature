Feature: DataLabel properties
  In order to customize the appearance of a data label on a chart
  As a developer using python-pptx
  I need read/write properties on DataLabel


  Scenario Outline: Get DataLabel.font
    Given a data label <having-or-not> custom font
     Then data_label.font is a Font object

    Examples: data_label.font states
      | having-or-not |
      | having a      |
      | having no     |


  Scenario Outline: Get DataLabel.position
    Given a data label positioned <relation-to> its data point
     Then data_label.position is <value>

    Examples: data_label.position cases
      | relation-to                | value  |
      | in unspecified relation to | None   |
      | centered on                | CENTER |


  Scenario Outline: Set DataLabel.position
    Given a data label positioned <relation-to> its data point
     When I assign <value> to data_label.position
     Then data_label.position is <value>

    Examples: data_label.position assignment cases
      | relation-to                | value  |
      | in unspecified relation to | CENTER |
      | centered on                | BELOW  |
      | below                      | None   |


  Scenario Outline: Get DataLabel.has_text_frame
    Given a data label <having-or-not> custom text
     Then data_label.has_text_frame is <value>

    Examples: text frame presence cases
      | having-or-not | value |
      | having        | True  |
      | having no     | False |


  Scenario Outline: Set DataLabel.has_text_frame
    Given a data label <having-or-not> custom text
     When I assign <new-value> to data_label.has_text_frame
     Then data_label.has_text_frame is <value>

    Examples: data_label.has_text_frame assignment cases
      | having-or-not | new-value | value |
      | having no     | True      | True  |
      | having        | False     | False |
      | having no     | False     | False |
      | having        | True      | True  |


  Scenario Outline: Get DataLabel.text_frame
    Given a data label <having-or-not> custom text
     Then data_label.text_frame is a TextFrame object

    Examples: text frame presence cases
      | having-or-not |
      | having        |
      | having no     |
