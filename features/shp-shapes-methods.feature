Feature: Shape collection methods
  In order to add a shape to a shape collection
  As a developer using python-pptx
  I need a set of methods on Shapes objects


  Scenario: SlideShapes.add_connector()
    Given a SlideShapes object
     When I call shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4)
     Then connector is a Connector object
      And connector.begin_x == 1
      And connector.begin_y == 2
      And connector.end_x == 3
      And connector.end_y == 4
