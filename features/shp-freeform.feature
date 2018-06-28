Feature: Build a freeform shape
  In order to form a desired freeform shape
  As a developer using python-pptx
  I need a set of properties and methods on a FreeformBuilder object


  Scenario Outline: FreeformBuilder.add_vertices()
    Given a FreeformBuilder object as builder
      And (builder._start_x, builder._start_y) is (<start_x>, <start_y>)
      And (builder._x_scale, builder._y_scale) is (<x_scale>, <y_scale>)
     When I call builder.add_line_segments([(100, 25), (25, 100)])
      And I assign builder.convert_to_shape() to shape
     Then shape.left == <left>
      And shape.top == <top>
      And shape.width == <width>
      And shape.height == <height>

    Examples: Pen start position and scaling factor
      | start_x | start_y | x_scale | y_scale | left | top | width | height |
      | 0       | 0       | 1.0     | 1.0     | 0    | 0   | 100   | 100    |
      | 25      | 25      | 1.0     | 1.0     | 25   | 25  | 75    | 75     |
      | 0       | 0       | 2.0     | 2.0     | 0    | 0   | 200   | 200    |
      | 25      | 25      | 2.0     | 2.0     | 50   | 50  | 150   | 150    |
      | 25      | 10      | 2.0     | 3.0     | 50   | 30  | 150   | 270    |


  Scenario Outline: FreeformBuilder.convert_to_shape()
    Given a FreeformBuilder object as builder
      And (builder._start_x, builder._start_y) is (<start_x>, <start_y>)
     When I call builder.add_line_segments([(100, 25), (25, 100)])
      And I assign builder.convert_to_shape(<origin_x>, <origin_y>) to shape
     Then shape.left == <x>
      And shape.top == <y>
      And shape.width == <cx>
      And shape.height == <cy>

    Examples: Local origin position
      | start_x | start_y | origin_x | origin_y | x  | y  | cx  | cy  |
      | 0       | 0       | 0        | 0        | 0  | 0  | 100 | 100 |
      | 0       | 0       | 42       | 0        | 42 | 0  | 100 | 100 |
      | 0       | 0       | 0        | 42       | 0  | 42 | 100 | 100 |
      | 0       | 0       | 24       | 24       | 24 | 24 | 100 | 100 |
      | 25      | 0       | 24       | 24       | 49 | 24 | 75  | 100 |
      | 25      | 12      | 24       | 24       | 49 | 36 | 75  | 88  |
      | 10      | 20      | 30       | 40       | 40 | 60 | 90  | 80  |
