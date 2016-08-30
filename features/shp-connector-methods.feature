Feature: Connector methods
  In order to modify a connector
  As a developer using python-pptx
  I need a set of methods on Connector objects


  Scenario: Connector.begin_connect()
    Given a connector and a 1 inch square picture at 0, 0 
     When I call connector.begin_connect(picture, 3)
     Then connector.begin_x == 914400
      And connector.begin_y == 457200


  Scenario: Connector.end_connect()
    Given a connector and a 1 inch square picture at 0, 0 
     When I call connector.end_connect(picture, 3)
     Then connector.end_x == 914400
      And connector.end_y == 457200
