Feature: Query and change common shape properties
  In order to identify and characterize shapes
  As a developer using python-pptx
  I need a set of common properties available for all shapes

  Scenario Outline: get slide on which shape appears
     Given a <shape type> on a slide
      Then I can access the slide from the shape

    Examples: Shape types
      | shape type  |
      | shape       |
      | picture     |
      | table       |
      | group shape |
      | connector   |


  Scenario Outline: determine whether a shape can contain text
     Given a <shape type>
      Then I can determine the shape <has_textframe status>

    Examples: Shape types
      | shape type  | has_textframe status |
      | shape       | has a text frame     |
      | picture     | has no text frame    |
      | table       | has no text frame    |
      | group shape | has no text frame    |
      | connector   | has no text frame    |
