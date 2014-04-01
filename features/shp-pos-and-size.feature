Feature: Query and change shape position and size
  In order to manipulate shapes on an existing slide
  As a developer using python-pptx
  I need to get and set the position and size of a shape

  Scenario: get position and size of existing shape
     Given a shape of known position and size
      Then the position and size of the shape matches the known values

  Scenario: change position and size of an existing shape
     Given a shape of known position and size
      When I change the position and size of the shape
      Then the position and size of the shape matches the new values

  Scenario: get position and size of existing picture
     Given a picture of known position and size
      Then the position and size of the picture matches the known values

  Scenario: change position and size of an existing picture
     Given a picture of known position and size
      When I change the position and size of the picture
      Then the position and size of the picture matches the new values

  Scenario: get position and size of a table
     Given a table of known position and size
      Then the position and size of the table match the known values

  Scenario: change position of an existing table
     Given a table of known position and size
      When I change the position of the table
      Then the position of the table matches the new values

  Scenario: get position of a group shape
     Given a group shape
      Then the left and top of the group shape match their known values

  Scenario: change position of a group shape
     Given a group shape
      When I change the position of the group shape
      Then the left and top of the group shape match the new values

  Scenario: get position of a connector
     Given a connector
      Then the left and top of the connector match their known values

  Scenario: change position of a connector
     Given a connector
      When I change the position of the connector
      Then the left and top of the connector match the new values
