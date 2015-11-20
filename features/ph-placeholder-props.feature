Feature: Get and change common slide placeholder properties
  In order to identify and characterize placeholder shapes
  As a developer using python-pptx
  I need a common set of properties available on any slide placeholder


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


  Scenario Outline: Get placeholder position and size
     Given an unpopulated <type> placeholder shape
      Then the placeholder's position and size are inherited from its layout

    Examples: Unpopulated placeholder types
      | type      |
      | content   |
      | text      |
      | chart     |
      | table     |
      | smart art |
      | media     |
      | clip art  |
      | picture   |


  Scenario Outline: Get placeholder format
     Given an unpopulated <type> placeholder shape
      Then shape.placeholder_format is its _PlaceholderFormat object

    Examples: Unpopulated placeholder types
      | type      |
      | content   |
      | text      |
      | chart     |
      | table     |
      | smart art |
      | media     |
      | clip art  |
      | picture   |
