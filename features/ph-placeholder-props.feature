Feature: Get and change common slide placeholder properties
  In order to identify and characterize placeholder shapes
  As a developer using python-pptx
  I need a common set of properties available on any slide placeholder


  @wip
  Scenario Outline: Get placeholder shape type
     Given an unpopulated <type> placeholder shape
      Then shape.shape_type is MSO_SHAPE_TYPE.PLACEHOLDER

    Examples: Unpopulated placeholder types
      | type      |
      | title     |
      | content   |
      | text      |
      | chart     |
      | table     |
      | smart art |
      | media     |
      | clip art  |
      | picture   |
