Feature: Query and change shape position and size
  In order to manipulate shapes on an existing slide
  As a developer using python-pptx
  I need to get and set the position and size of a shape

  Scenario Outline: get the position of a shape
    Given a <shape-type>
     Then the left and top of the <shape-type> match their known values

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: change the position of a shape
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


  Scenario Outline: get the size of a shape
    Given a <shape-type>
     Then the width and height of the <shape-type> match their known values

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: change the size of a shape
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
