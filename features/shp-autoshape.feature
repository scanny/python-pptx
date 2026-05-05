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


  Scenario: BaseShape.theme_style_refs on a newly-added auto-shape (issue #447)
     Given a 1-inch rectangle shape
      Then shape.theme_style_refs is (1, 3, 2, "minor")


  Scenario: BaseShape.theme_style_refs writes a new p:style (issue #447)
     Given a 1-inch rectangle shape
      When I assign shape.theme_style_refs = (2, 4, 1, "major")
      Then shape.theme_style_refs is (2, 4, 1, "major")


  Scenario: BaseShape.theme_style_refs can be cleared (issue #447)
     Given a 1-inch rectangle shape
      When I assign shape.theme_style_refs = None
      Then shape.theme_style_refs is None


  Scenario: BaseShape.theme_style reads per-ref schemeClr colors (CLO-7)
     Given a 1-inch rectangle shape
      Then shape.theme_style is (1, "accent1", 3, "accent1", 2, "accent1", "minor", "lt1")


  Scenario: BaseShape.theme_style writes a new p:style with distinct colors (CLO-7)
     Given a 1-inch rectangle shape
      When I assign shape.theme_style = (2, "accent2", 4, "bg1", 1, "tx2", "major", "accent3")
      Then shape.theme_style is (2, "accent2", 4, "bg1", 1, "tx2", "major", "accent3")


  Scenario: BaseShape.theme_style can be cleared (CLO-7)
     Given a 1-inch rectangle shape
      When I assign shape.theme_style = None
      Then shape.theme_style is None
       And shape.theme_style_refs is None


  Scenario: theme_style_refs preserves existing schemeClr colors on re-assignment (CLO-7)
     Given a 1-inch rectangle shape
      When I assign shape.theme_style = (1, "accent5", 3, "accent5", 2, "accent5", "minor", "tx1")
      And I assign shape.theme_style_refs = (4, 5, 6, "major")
      Then shape.theme_style is (4, "accent5", 5, "accent5", 6, "accent5", "major", "tx1")
