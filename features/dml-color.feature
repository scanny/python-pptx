Feature: Control color
  In order to fine-tune the color of filled areas and lines
  As a developer using python-pptx
  I need properties and methods on ColorFormat


  Scenario: ColorFormat.rgb
    Given a ColorFormat object as color
     When I assign RGBColor(12, 34, 56) to color.rgb
     Then color.rgb is RGBColor(12, 34, 56)


  Scenario: ColorFormat.theme_color
    Given a ColorFormat object as color
     When I assign MSO_THEME_COLOR.ACCENT_6 to color.theme_color
     Then color.theme_color is MSO_THEME_COLOR.ACCENT_6


  Scenario: ColorFormat.brightness
    Given a ColorFormat object as color
     When I assign 0.42 to color.brightness
     Then color.brightness is 0.42
