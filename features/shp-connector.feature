Feature: Connector properties and methods
  In order to characterize and adjust Connectors
  As a developer using python-pptx
  I need properties and methods on Connector


  Scenario: Connector.begin_connect()
    Given a connector and a 1 inch square picture at 0, 0
     When I call connector.begin_connect(picture, 3)
     Then connector.begin_x == 914400
      And connector.begin_y == 457200


  Scenario Outline: Get Connector.begin_x/y
    Given a connector having its begin point at (<x>, <y>)
     Then connector.begin_x is an Emu object with value <x>
      And connector.begin_y is an Emu object with value <y>

    Examples: Connector begin point states
      |   x    |   y    |
      | 914400 | 914400 |


  Scenario Outline: Get Connector.end_x/y
    Given a connector having its end point at (<x>, <y>)
     Then connector.end_x is an Emu object with value <x>
      And connector.end_y is an Emu object with value <y>

    Examples: Connector end point states
      |    x    |    y    |
      | 1828800 | 1828800 |


  Scenario Outline: Set Connector.begin_x
    Given a connector having its begin point at (<x>, <y>)
     When I assign <value> to connector.begin_x
     Then connector.begin_x is an Emu object with value <value>

    Examples: Connector begin point assignment results
      |   x    |   y    | value   |
      | 914400 | 914400 | 1828800 |


  Scenario Outline: Set Connector.begin_y
    Given a connector having its begin point at (<x>, <y>)
     When I assign <value> to connector.begin_y
     Then connector.begin_y is an Emu object with value <value>

    Examples: Connector begin point assignment results
      |   x    |   y    | value   |
      | 914400 | 914400 | 1828800 |


  Scenario: Connector.end_connect()
    Given a connector and a 1 inch square picture at 0, 0
     When I call connector.end_connect(picture, 3)
     Then connector.end_x == 914400
      And connector.end_y == 457200


  Scenario Outline: Set Connector.end_x
    Given a connector having its end point at (<x>, <y>)
     When I assign <value> to connector.end_x
     Then connector.end_x is an Emu object with value <value>

    Examples: Connector end point assignment results
      |   x     |   y     | value  |
      | 1828800 | 1828800 | 914400 |


  Scenario Outline: Set Connector.end_y
    Given a connector having its end point at (<x>, <y>)
     When I assign <value> to connector.end_y
     Then connector.end_y is an Emu object with value <value>

    Examples: Connector end point assignment results
      |   x     |   y     | value  |
      | 1828800 | 1828800 | 914400 |


  Scenario: Connector.line
    Given a Connector object as shape
     Then shape.line is a LineFormat object
