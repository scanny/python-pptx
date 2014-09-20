Feature: Query and change common shape properties
  In order to identify and characterize shapes
  As a developer using python-pptx
  I need a set of common properties available for all shapes

  Scenario Outline: get shape id
     Given a <shape-type>
      Then I can get the id of the <shape-type>

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: get shape name
     Given a <shape-type>
      Then I can get the name of the <shape-type>

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: get slide on which shape appears
     Given a <shape-type> on a slide
      Then I can access the slide from the shape

    Examples: Shape types
      | shape-type    |
      | shape         |
      | picture       |
      | graphic frame |
      | group shape   |
      | connector     |


  Scenario Outline: determine whether a shape can contain text
     Given a <shape-type>
      Then I can determine the shape <has_text_frame status>

    Examples: Shape types
      | shape-type    | has_text_frame status |
      | shape         | has a text frame      |
      | picture       | has no text frame     |
      | graphic frame | has no text frame     |
      | group shape   | has no text frame     |
      | connector     | has no text frame     |
