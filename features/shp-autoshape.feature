Feature: Shape properties and methods
  In order to identify and adjust autoshapes
  As a developer using python-pptx
  I need properties and methods on Shape


  Scenario: Shape.adjustments setter
     Given a chevron shape
      When I assign 0.15 to shape.adjustments[0]
      Then shape.adjustments[0] is 0.15


  Scenario: Shape.line
     Given a Shape object as shape
      Then shape.line is a LineFormat object


  Scenario: Shape.text getter
     Given a Shape object having text as shape
      Then shape.text == "Fee Fi\vF\xf8\xf8 Fum\nI am a shape\vwith textium"


  Scenario: Shape.text setter
     Given a Shape object having text as shape
      When I assign shape.text = "F\xf8o\vBar\nBaz\x1b"
      Then shape.text == "F\xf8o\vBar\nBaz_x001B_"


  Scenario: Shape.path_geometry for a static preset shape
     Given a 1-inch rectangle shape
      Then shape.path_geometry is a PathGeometry object
      And shape.path_geometry has 1 path
      And the first path traces the shape's bounding-box rectangle


  Scenario: Shape.path_geometry for a dynamic preset shape
     Given a rounded rectangle shape
      Then shape.path_geometry is None
  Scenario: Shape.auto_shape_type for a shape with prst="line" (issue #749)
     Given an auto-shape with prst="line" as shape
      Then shape.auto_shape_type is MSO_SHAPE.LINE


  Scenario: Shape.auto_shape_type for a shape with an unknown prst (issue #749)
     Given an auto-shape with an unknown prst as shape
      Then shape.auto_shape_type is None
