Feature: Common shape properties
  In order to interact with shapes of assorted types
  As a developer using python-pptx
  I need a common set of properties available on all shapes


  Scenario Outline: Get shape.click_action
    Given a <shape-type> object as shape
     Then shape.click_action is an ActionSetting object

    Examples: Shape types
      | shape-type   |
      | Shape        |
      | Picture      |
      | GraphicFrame |
      | Connector    |


  Scenario Outline: shape.has_chart for shapes that never contain a chart
    Given a <shape-type> object as shape
     Then shape.has_chart is False

    Examples: Shape types
      | shape-type |
      | Connector  |
      | Shape      |
      | Picture    |
      | GroupShape |


  Scenario Outline: shape.has_table for shapes that never contain a table
    Given a <shape-type> object as shape
     Then shape.has_table is False

    Examples: Shape types
      | shape-type |
      | Connector  |
      | Shape      |
      | Picture    |
      | GroupShape |


  Scenario Outline: Get shape.has_text_frame
    Given a <shape-type> object as shape
     Then shape.has_text_frame is <value>

    Examples: Shape types
      | shape-type   | value |
      | Shape        | True  |
      | Picture      | False |
      | GraphicFrame | False |
      | GroupShape   | False |
      | Connector    | False |


  Scenario Outline: Get shape.left and shape.top
    Given a <shape-type> object as shape
     Then shape.left == <left>
      And shape.top == <top>

    Examples: Shape types
      | shape-type   |  left   |   top   |
      | Shape        | 1339552 |  692696 |
      | Picture      | 2711152 | 1835696 |
      | GraphicFrame | 4082752 | 2978696 |
      | GroupShape   | 5454352 | 4121696 |
      | Connector    | 6825952 | 5264696 |


  Scenario Outline: Set shape.left and shape.top
    Given a <shape-type> object as shape
     When I assign <left> to shape.left
      And I assign <top> to shape.top
     Then shape.left == <left>
      And shape.top == <top>

    Examples: Shape types
      | shape-type   |  left   |   top   |
      | Shape        |  692696 | 1339552 |
      | Picture      | 1835696 | 2711152 |
      | GraphicFrame | 2978696 | 4082752 |
      | GroupShape   | 4121696 | 5454352 |
      | Connector    | 5264696 | 6825952 |


  Scenario Outline: Get shape.name
    Given a <shape-type> object as shape
     Then shape.name == '<name>'

    Examples: Shape types
      | shape-type   | name                |
      | Shape        | Rounded Rectangle 1 |
      | Picture      | Picture 2           |
      | GraphicFrame | Table 3             |
      | GroupShape   | Group 8             |
      | Connector    | Elbow Connector 10  |


  Scenario Outline: Set shape.name
    Given a <shape-type> object as shape
     When I assign '<value>' to shape.name
     Then shape.name == '<expected-value>'

    Examples: Expected results of changing shape.name
      | shape-type   | value                | expected-value       |
      | Shape        | New Shape 42         | New Shape 42         |
      | Picture      | New Picture 42       | New Picture 42       |
      | GraphicFrame | New Graphic Frame 42 | New Graphic Frame 42 |
      | GroupShape   | New Group 42         | New Group 42         |
      | Connector    | New Connector 42     | New Connector 42     |


  Scenario Outline: Get shape.part
    Given a <shape-type> object on a slide as shape
     Then shape.part is a SlidePart object
      And shape.part is slide.part

    Examples: Shape types
      | shape-type   |
      | Shape        |
      | Picture      |
      | GraphicFrame |
      | GroupShape   |
      | Connector    |


  Scenario Outline: Get shape.rotation
    Given a rotated <shape-type> object as shape
     Then shape.rotation == <value>

    Examples: Shape types
      | shape-type   | value |
      | Shape        | 10.0  |
      | Picture      | 20.0  |
      | GraphicFrame |  0.0  |
      | GroupShape   | 40.0  |
      | Connector    | 50.0  |


  Scenario Outline: Set shape.rotation
    Given a <shape-type> object as shape
     When I assign <value> to shape.rotation
     Then shape.rotation == <expected-value>

    Examples: Shape types
      | shape-type   | value | expected-value |
      | Shape        |  12.3 |      12.3      |
      | Picture      | 520.4 |     160.4      |
      | GraphicFrame |  10.0 |      10.0      |
      | GroupShape   |  -5.2 |     354.8      |
      | Connector    |  50.0 |      50.0      |


  Scenario Outline: shape.shadow
    Given a <shape-type> object as shape
     Then shape.shadow is a ShadowFormat object

    Examples: Shape types
      | shape-type   |
      | Shape        |
      | Picture      |
      | GroupShape   |
      | Connector    |


  Scenario: GraphicFrame.shadow (not-implemented)
    Given a GraphicFrame object as shape
     Then shape.shadow raises NotImplementedError


  Scenario Outline: Get shape.shape_id
    Given a <shape-type> object as shape
     Then shape.shape_id == <value>

    Examples: Shape types
      | shape-type   | value |
      | Shape        |   2   |
      | Picture      |   3   |
      | GraphicFrame |   4   |
      | GroupShape   |   9   |
      | Connector    |  11   |


  Scenario Outline: Get shape.width and shape.height
    Given a <shape-type> object as shape
     Then shape.width == <width>
      And shape.height == <height>

    Examples: Shape types
      | shape-type   | width  | height |
      | Shape        | 928192 | 914400 |
      | Picture      | 914400 | 945232 |
      | GraphicFrame | 993304 | 914400 |
      | GroupShape   | 914400 | 914400 |
      | Connector    | 986408 | 828600 |


  Scenario Outline: Set shape.width and shape.height
    Given a <shape-type> object as shape
     When I assign <width> to shape.width
      And I assign <height> to shape.height
     Then shape.width == <width>
      And shape.height == <height>

    Examples: Shape types
      | shape-type   |  width  | height  |
      | Shape        |  692696 | 1339552 |
      | Picture      | 1835696 | 2711152 |
      | GraphicFrame | 2978696 | 4082752 |
      | GroupShape   | 4121696 | 5454352 |
      | Connector    | 5264696 | 6825952 |
