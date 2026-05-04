Feature: Access and modify data labels properties
  In order to customize the appearance of data labels on a chart
  As a developer using python-pptx
  I need access to data label objects and their properties

# ---DataLabels---

  Scenario Outline: DataLabel.show_category_name getter
    Given a DataLabels object <showing-or-not> category-name as data_labels
     Then data_labels.show_category_name is <value>

    Examples: data_label.show_category_name states
      | showing-or-not | value |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_category_name setter
    Given a DataLabels object <showing-or-not> category-name as data_labels
     When I assign <value> to data_labels.show_category_name
     Then data_labels.show_category_name is <value>

    Examples: data_label.show_category_name states
      | showing-or-not | value |
      | not showing    | True  |
      | showing        | False |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_legend_key getter
    Given a DataLabels object <showing-or-not> legend-key as data_labels
     Then data_labels.show_legend_key is <value>

    Examples: data_label.show_legend_key states
      | showing-or-not | value |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_legend_key setter
    Given a DataLabels object <showing-or-not> legend-key as data_labels
     When I assign <value> to data_labels.show_legend_key
     Then data_labels.show_legend_key is <value>

    Examples: data_label.show_legend_key states
      | showing-or-not | value |
      | not showing    | True  |
      | showing        | False |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_percentage getter
    Given a DataLabels object <showing-or-not> percentage as data_labels
     Then data_labels.show_percentage is <value>

    Examples: data_label.show_percentage states
      | showing-or-not | value |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_percentage setter
    Given a DataLabels object <showing-or-not> percentage as data_labels
     When I assign <value> to data_labels.show_percentage
     Then data_labels.show_percentage is <value>

    Examples: data_label.show_percentage states
      | showing-or-not | value |
      | not showing    | True  |
      | showing        | False |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_series_name getter
    Given a DataLabels object <showing-or-not> series-name as data_labels
     Then data_labels.show_series_name is <value>

    Examples: data_label.show_series_name states
      | showing-or-not | value |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_series_name setter
    Given a DataLabels object <showing-or-not> series-name as data_labels
     When I assign <value> to data_labels.show_series_name
     Then data_labels.show_series_name is <value>

    Examples: data_label.show_series_name states
      | showing-or-not | value |
      | not showing    | True  |
      | showing        | False |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_value getter
    Given a DataLabels object <showing-or-not> value as data_labels
     Then data_labels.show_series_name is <value>

    Examples: data_label.show_value states
      | showing-or-not | value |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabel.show_value setter
    Given a DataLabels object <showing-or-not> value as data_labels
     When I assign <value> to data_labels.show_value
     Then data_labels.show_value is <value>

    Examples: data_label.show_value states
      | showing-or-not | value |
      | not showing    | True  |
      | showing        | False |
      | not showing    | False |
      | showing        | True  |


  Scenario Outline: DataLabels.position getter
    Given a DataLabels object with <position> position as data_labels
     Then data_labels.position is <expected-value>

    Examples: data_labels position values
      | position    | expected-value |
      | inherited   | None           |
      | inside-base | INSIDE_BASE    |


  Scenario Outline: DataLabels.position setter
    Given a DataLabels object with <position> position as data_labels
     When I assign <new-value> to data_labels.position
     Then data_labels.position is <expected-value>

    Examples: expected results of assignment to data_labels.position
      | position    | new-value   | expected-value |
      | inherited   | INSIDE_END  | INSIDE_END     |
      | inside-base | OUTSIDE_END | OUTSIDE_END    |
      | inside-base | None        | None           |

  Scenario: DataLabels.format exposes fill and line for color customization
    Given a DataLabels object with inherited position as data_labels
     Then data_labels.format is a ChartFormat object
      And data_labels.format.fill is a FillFormat object
      And data_labels.format.line is a LineFormat object


  Scenario: DataLabels.text_frame returns a TextFrame wrapping c:dLbls/c:txPr
    Given a DataLabels object with inherited position as data_labels
     Then data_labels.text_frame is a TextFrame object
      And data_labels.text_frame wraps the c:dLbls/c:txPr element


  Scenario: Assigning DataLabels.text_frame.word_wrap writes c:txPr/a:bodyPr/@wrap
    Given a DataLabels object with inherited position as data_labels
     When I assign False to data_labels.text_frame.word_wrap
     Then data_labels.text_frame.word_wrap is False
      And the c:dLbls XML has c:txPr/a:bodyPr with wrap="none"
      And the c:dLbls XML has no c:tx/c:rich subtree


# ---DataLabel---

  Scenario Outline: DataLabel.font
    Given a data label <having-or-not> custom font as data_label
     Then data_label.font is a Font object

    Examples: data_label.font states
      | having-or-not |
      | having a      |
      | having no     |


  Scenario: DataLabel.format
    Given a data label
     Then data_label.format is a ChartFormat object
      And data_label.format.fill is a FillFormat object
      And data_label.format.line is a LineFormat object


  Scenario Outline: DataLabel.has_text_frame getter
    Given a data label <having-or-not> custom text as data_label
     Then data_label.has_text_frame is <value>

    Examples: text frame presence cases
      | having-or-not | value |
      | having        | True  |
      | having no     | False |


  Scenario Outline: DataLabel.has_text_frame setter
    Given a data label <having-or-not> custom text as data_label
     When I assign <new-value> to data_label.has_text_frame
     Then data_label.has_text_frame is <value>

    Examples: data_label.has_text_frame assignment cases
      | having-or-not | new-value | value |
      | having no     | True      | True  |
      | having        | False     | False |
      | having no     | False     | False |
      | having        | True      | True  |


  Scenario Outline: DataLabel.position getter
    Given a data label with <position> position as data_label
     Then data_label.position is <value>

    Examples: data_label.position cases
      | position  | value  |
      | inherited | None   |
      | centered  | CENTER |


  Scenario Outline: DataLabel.position setter
    Given a data label with <position> position as data_label
     When I assign <value> to data_label.position
     Then data_label.position is <value>

    Examples: data_label.position assignment cases
      | position  | value  |
      | inherited | CENTER |
      | centered  | BELOW  |
      | below     | None   |


  Scenario Outline: DataLabel.text_frame
    Given a data label <having-or-not> custom text as data_label
     Then data_label.text_frame is a TextFrame object

    Examples: text frame presence cases
      | having-or-not |
      | having        |
      | having no     |


  Scenario: DataLabel.number_format defaults to 'General' (issues #638 and #803)
    Given a data label
     Then data_label.number_format is 'General'
      And data_label.number_format_is_linked is True


  Scenario: Assigning DataLabel.number_format writes c:dLbl/c:numFmt/@formatCode
    Given a data label
     When I assign '0.00%' to data_label.number_format
     Then data_label.number_format is '0.00%'
      And data_label.number_format_is_linked is False
      And the c:dLbl for the point has c:numFmt with formatCode '0.00%' and sourceLinked '0'


  Scenario: DataLabel.number_format_is_linked setter toggles sourceLinked
    Given a data label
     When I assign '0.0%' to data_label.number_format
      And I assign True to data_label.number_format_is_linked
     Then data_label.number_format_is_linked is True
      And the c:dLbl for the point has c:numFmt with formatCode '0.0%' and sourceLinked '1'
