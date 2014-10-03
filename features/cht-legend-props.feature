Feature: Get and set legend properties
  In order to customize the appearance of the legend on a chart
  As a developer using python-pptx
  I need a way to get and set legend properties


  Scenario: Access legend font
    Given a legend
     Then legend.font is a Font object


  Scenario Outline: Determine legend horizontal offset
    Given a legend having horizontal offset of <value>
     Then legend.horz_offset is <expected-value>

    Examples: legend horizontal offset values
      | value | expected-value |
      | none  | 0.0            |
      | -0.5  | -0.5           |
      | 0.42  | 0.42           |


  Scenario Outline: Change legend horizontal offset
    Given a legend having horizontal offset of <value>
     When I assign <new-value> to legend.horz_offset
     Then legend.horz_offset is <expected-value>

    Examples: expected results of changing horz_offset
      | value  | new-value | expected-value |
      | none   |   -1.0    |     -1.0       |
      | -0.5   |    0.5    |      0.5       |
      |  0.42  |   -0.5    |     -0.5       |
      | -0.5   |    0      |      0.0       |


  Scenario Outline: Determine whether the legend is beside or overlays the chart
    Given a legend with overlay setting of <setting>
     Then legend.include_in_layout is <expected-value>

    Examples: legend.include_in_layout expected values
      | setting             | expected-value |
      | no explicit setting | True           |
      | True                | True           |
      | False               | False          |


  Scenario Outline: Change whether the legend is beside or overlays the chart
    Given a legend with overlay setting of <setting>
     When I assign <value> to legend.include_in_layout
     Then legend.include_in_layout is <expected-value>

    Examples: expected results of changing legend.include_in_layout
      | setting             | value | expected-value |
      | no explicit setting | True  | True           |
      | no explicit setting | False | False          |
      | True                | False | False          |
      | True                | True  | True           |
      | False               | True  | True           |
      | False               | False | False          |


  Scenario Outline: Determine legend position
    Given a legend positioned <location> the chart
     Then legend.position is <expected-value>

    Examples: legend position values
      | location                      | expected-value |
      | at an unspecified location of | RIGHT          |
      | below                         | BOTTOM         |
      | to the right of               | RIGHT          |


  Scenario Outline: Change legend position
    Given a legend positioned <location> the chart
     When I assign <new-value> to legend.position
     Then legend.position is <expected-value>

    Examples: expected results of changing legend.position
      | location                      | new-value | expected-value |
      | at an unspecified location of | RIGHT     | RIGHT          |
      | below                         | TOP       | TOP            |
      | to the right of               | BOTTOM    | BOTTOM         |
