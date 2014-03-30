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
