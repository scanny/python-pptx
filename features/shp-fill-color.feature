Feature: Change shape fill color
  In order to fine-tune the visual experience produced by a slide
  As a developer using python-pptx
  I need to specify the precise fill color of an autoshape

  Scenario: set shape fill to transparent
     Given an autoshape
      When I set the fill type to background
      Then the fill type of the shape is background

  Scenario: set shape solid fill RGB color
     Given an autoshape
      When I set the fill type to solid
       And I set the foreground color to an RGB value
      Then the foreground color of the shape is the RGB value I set

  Scenario: set shape solid fill theme color
     Given an autoshape
      When I set the fill type to solid
       And I set the foreground color to a theme color
      Then the foreground color of the shape is the theme color I set

  Scenario: set shape solid fill brightness
     Given an autoshape
      When I set the fill type to solid
       And I set the foreground color to an RGB value
       And I set the foreground color brightness to 0.5
      Then the foreground color brightness of the shape is 0.5
