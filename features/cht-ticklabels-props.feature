Feature: Get and set tick label properties
  In order to customize the appearance of the tick labels on a chart
  As a developer using python-pptx
  I need a way to get and set tick label properties


  Scenario Outline: Access tick label offset for an axis
    Given tick labels having an offset of <xml-value>
     Then tick_labels.offset is <expected-value>

    Examples: expected values of TickLabels.offset
      | xml-value           | expected-value |
      | no explicit setting | 100            |
      | 420                 | 420            |


  Scenario Outline: Change tick label offset
    Given tick labels having an offset of <xml-value>
     When I assign <new-value> to tick_labels.offset
     Then tick_labels.offset is <new-value>

    Examples: expected values of TickLabels.offset
      | xml-value           | new-value |
      | no explicit setting | 100       |
      | no explicit setting | 420       |
      | 420                 | 100       |


  Scenario Outline: Access tick label rotation for an axis
    Given tick labels having a rotation of <xml-value>
     Then tick_labels.rotation is <expected-value>

    Examples: expected values of TickLabels.rotation
      | xml-value           | expected-value |
      | no explicit setting | 0.0            |
      | 45                  | 45.0           |
      | -90                 | 270.0          |


  Scenario Outline: Change tick label rotation
    Given tick labels having a rotation of <xml-value>
     When I assign <new-value> to tick_labels.rotation
     Then tick_labels.rotation is <expected-value>

    Examples: expected values of TickLabels.rotation
      | xml-value           | new-value | expected-value |
      | no explicit setting | 45        | 45.0           |
      | no explicit setting | 45.5      | 45.5           |
      | no explicit setting | -90       | 270.0          |
      | 45                  | 0         | 0.0            |
      | 45                  | 90        | 90.0           |
