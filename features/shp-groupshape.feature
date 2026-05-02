Feature: GroupShape properties and methods
  In order to operate on a GroupShape shape
  As a developer using python-pptx
  I need properties and methods on GroupShape


  Scenario: Get GroupShape.click_action raises TypeError
    Given a GroupShape object as shape
     Then accessing shape.click_action raises TypeError


  Scenario: GroupShape.shape_type
    Given a GroupShape object as shape
     Then shape.shape_type == MSO_SHAPE_TYPE.GROUP


  Scenario: GroupShape.shapes
    Given a GroupShape object as group_shape
     Then group_shape.shapes is a GroupShapes object


  Scenario: GroupShape position and size, 0 shapes
    Given an empty GroupShape object as shape
     Then shape.left == 0
      And shape.top == 0
      And shape.width == 0
      And shape.height == 0


  Scenario: GroupShape position and size, 1 shapes
    Given an empty GroupShape object as shape
     When I add a 100 x 200 shape at (300, 400)
     Then shape.left == 300
      And shape.top == 400
      And shape.width == 100
      And shape.height == 200


  Scenario: GroupShape position and size, 2 shapes
    Given an empty GroupShape object as shape
     When I add a 100 x 200 shape at (300, 400)
      And I add a 150 x 250 shape at (450, 650)
     Then shape.left == 300
      And shape.top == 400
      And shape.width == 300
      And shape.height == 500


  Scenario: Child shape of a resized group reports effective slide geometry
    Given a group shape whose child space is scaled 2:1 onto the slide
     Then the child shape's raw shape.left matches its group-local xfrm
      And the child shape's effective_left is the slide-relative coordinate
      And the child shape's effective_width is half its raw width


  Scenario: GroupShape.ungroup() hoists children to the slide at slide coords
    Given a top-level group of two autoshapes at known slide positions
     When I call group.ungroup()
     Then the group is removed from the shape tree
      And each freed shape is a top-level sibling on the slide
      And each freed shape's effective slide rectangle is unchanged


  Scenario: GroupShape.ungroup() resolves nested-group cascade correctly
    Given a nested group whose outer maps child space 2:1 onto the slide
     When I call ungroup() on the inner group
     Then the freed child renders at the same slide rectangle as before
