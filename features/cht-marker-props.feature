Feature: Get and set marker properties
  In order to customize the visual display of a data point marker
  As a developer using python-pptx
  I need a way to get and set marker properties


  Scenario: Get Marker.format
    Given a marker
     Then marker.format is a ChartFormat object
      And marker.format.fill is a FillFormat object
      And marker.format.line is a LineFormat object


  Scenario Outline: Get Marker.size
    Given a marker having size of <size>
     Then marker.size is <value>

    Examples: Marker size value cases
      | size              | value |
      | no explicit value | None  |
      | 24 points         | 24    |
      | 36 points         | 36    |


  Scenario Outline: Set Marker.size
    Given a marker having size of <size>
     When I assign <value> to marker.size
     Then marker.size is <value>

    Examples: Marker size value cases
      | size              | value |
      | no explicit value | 24    |
      | 24 points         | 36    |
      | 36 points         | None  |
      | no explicit value | None  |


  Scenario Outline: Get Marker.style
    Given a marker having style of <shape>
     Then marker.style is <value>

    Examples: Marker shape value cases
      | shape             | value    |
      | no explicit value | None     |
      | circle            | CIRCLE   |
      | triangle          | TRIANGLE |


  Scenario Outline: Set Marker.style
    Given a marker having style of <style>
     When I assign <value> to marker.style
     Then marker.style is <value>

    Examples: Marker style value cases
      | style             | value    |
      | no explicit value | CIRCLE   |
      | circle            | TRIANGLE |
      | triangle          | None     |
      | no explicit value | None     |
