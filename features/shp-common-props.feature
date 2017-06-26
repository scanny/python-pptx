Feature: Common shape properties
  In order to interact with shapes of assorted types
  As a developer using python-pptx
  I need a common set of properties available on all shapes


  Scenario Outline: Get shape.shape_id
    Given a <shape-of-type>
     Then shape.shape_id == <value>

    Examples: Shape types
      | shape-of-type | value |
      | shape         |   2   |
      | picture       |   3   |
      | graphic frame |   4   |
      | group shape   |   9   |
      | connector     |  11   |


  Scenario Outline: Get shape.name
     Given a <shape-type>
      Then shape.name is '<name>'

    Examples: Shape types
      | shape-type    | name                |
      | shape         | Rounded Rectangle 1 |
      | picture       | Picture 2           |
      | graphic frame | Table 3             |
      | group shape   | Group 8             |
      | connector     | Elbow Connector 10  |


  Scenario Outline: Set shape.name
    Given a <shape-type>
     When I assign '<value>' to shape.name
     Then shape.name is '<expected-value>'

    Examples: Expected results of changing shape.name
      | shape-type    | value                | expected-value       |
      | shape         | New Shape 42         | New Shape 42         |
      | picture       | New Picture 42       | New Picture 42       |
      | graphic frame | New Graphic Frame 42 | New Graphic Frame 42 |
      | group shape   | New Group Shape 42   | New Group Shape 42   |
      | connector     | New Connector 42     | New Connector 42     |


  Scenario Outline: Get shape.part
     Given a <shape-type> on a slide
      Then shape.part is the SlidePart of the shape

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: Get shape.has_text_frame
     Given a <shape-of-type>
      Then shape.has_text_frame is <value>

    Examples: Shape types
      | shape-of-type | value |
      | shape         | True  |
      | picture       | False |
      | graphic frame | False |
      | group shape   | False |
      | connector     | False |


  Scenario Outline: Get shape.rotation
     Given a rotated <shape-type>
      Then shape.rotation is <value>

    Examples: Shape types
      | shape-type    | value |
      | shape         | 10.0  |
      | picture       | 20.0  |
      | graphic frame | 0.0   |
      | group shape   | 40.0  |
      | connector     | 50.0  |


  Scenario Outline: Set shape.rotation
     Given a <shape-type>
      When I assign <value> to shape.rotation
      Then shape.rotation is <expected-value>

    Examples: Shape types
      | shape-type    | value | expected-value |
      | shape         |  12.3 |      12.3      |
      | picture       | 520.4 |     160.4      |
      | graphic frame |  10.0 |      10.0      |
      | group shape   |  -5.2 |     354.8      |
      | connector     |  50.0 |      50.0      |


  Scenario Outline: Get shape.click_action
     Given <shape-type>
      Then shape.click_action is an ActionSetting object

    Examples: Shape types
      | shape-type      |
      | an autoshape    |
      | a textbox       |
      | a picture       |
      | a connector     |
      | a graphic frame |
      | a group shape   |


  Scenario Outline: Get shape.left and shape.top
    Given a <shape-type>
     Then the left and top of the <shape-type> match their known values

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: Set shape.left and shape.top
    Given a <shape-type>
     When I change the left and top of the <shape-type>
     Then the left and top of the <shape-type> match their new values

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: Get shape.width and shape.height
    Given a <shape-type>
     Then the width and height of the <shape-type> match their known values

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: Set shape.width and shape.height
    Given a <shape-type>
     When I change the width and height of the <shape-type>
     Then the width and height of the <shape-type> match their new values

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |
