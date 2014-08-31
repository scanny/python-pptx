Feature: Get and set legend properties
  In order to customize the appearance of the legend on a chart
  As a developer using python-pptx
  I need a way to get and set legend properties


  Scenario Outline: Determine legend horizontal offset
    Given a legend having horizontal offset of <value>
     Then legend.horz_offset is <expected-value>

    Examples: legend horizontal offset values
      | value | expected-value |
      | none  | 0.0            |
      | -0.5  | -0.5           |
      | 0.42  | 0.42           |


  @wip
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
