Feature: BaseShape.clone_onto
  In order to copy a shape (with all its referenced parts) between slides
  As a developer using python-pptx
  I need `BaseShape.clone_onto(shape_tree, left=None, top=None)` which deep-
  copies the shape onto the target tree, re-embeds every referenced part
  into the target package, and returns the appropriately-typed proxy


  Scenario: Clone an auto-shape between two slides in the same presentation
    Given a slide with an auto-shape
      And a second slide in the same presentation
     When I clone the auto-shape onto the second slide
     Then the clone is a Shape with a fresh shape id
      And the clone has the same text as the source
      And the clone is topmost in the second slide's z-order


  Scenario: Clone a picture re-embeds its image part on the target package
    Given a slide with a picture
      And a second slide in the same presentation
     When I clone the picture onto the second slide
     Then the clone is a Picture
      And the clone's image bytes match the source


  Scenario: Clone a picture into a different presentation
    Given a source presentation with a picture on slide 1
      And a target presentation with one empty slide
     When I clone the picture onto the target presentation's slide
     Then the clone is a Picture
      And the clone's image part lives in the target presentation's package


  Scenario: Clone a connector
    Given a slide with a connector
      And a second slide in the same presentation
     When I clone the connector onto the second slide
     Then the clone is a Connector with the same begin and end points


  Scenario: Clone a table graphic frame
    Given a slide with a table
      And a second slide in the same presentation
     When I clone the table onto the second slide
     Then the clone is a GraphicFrame that has a table


  Scenario: Clone an empty group shape
    Given a slide with a group shape
      And a second slide in the same presentation
     When I clone the group onto the second slide
     Then the clone is a GroupShape with a fresh shape id


  Scenario: Apply left/top override when cloning
    Given a slide with an auto-shape
      And a second slide in the same presentation
     When I clone the auto-shape onto the second slide at (3in, 4in)
     Then the clone's position is (3in, 4in)


  Scenario: Clone a group with three autoshapes preserves child positions
    Given a slide with a group of three autoshapes
      And a second slide in the same presentation
     When I clone the group onto the second slide
     Then the clone is a GroupShape containing three autoshapes
      And the clone's children have the same relative positions

  Scenario: Clone a nested group preserves every descendant
    Given a slide with a nested group shape
      And a second slide in the same presentation
     When I clone the nested group onto the second slide
     Then every descendant shape survives on the clone
      And every descendant id is unique and fresh

  Scenario: Clone a group containing a picture re-embeds the image
    Given a slide with a group containing a picture
      And a second slide in the same presentation
     When I clone the group onto the second slide
     Then the cloned group contains a picture with matching image bytes

  Scenario: Clone a group containing a chart re-embeds the chart part
    Given a slide with a group containing a chart
      And a second slide in the same presentation
     When I clone the group onto the second slide
     Then the cloned group contains a chart resolvable on the target part

  Scenario: Reposition cloned group preserves child coordinate system
    Given a slide with a group of two autoshapes
      And a second slide in the same presentation
     When I clone the group onto the second slide at (5in, 5in)
     Then the cloned group is at (5in, 5in)
      And the cloned group's child coordinate system is unchanged
