Feature: slide background
  In order to manipulate the background of a slide or master
  As a developer using python-pptx
  I need properties and methods on the _Background object


  Scenario Outline: _Background.fill
    Given a _Background object having <type> background as background
     Then background.fill is a FillFormat object

    Examples: _Background.fill cases
      | type              |
      | no                |
      | a fill            |
      | a style reference |


  Scenario: _Background.bg_element returns p:bg when present
    Given a _Background object having a fill background as background
     Then background.bg_element is the p:bg element


  Scenario: _Background.bg_element returns None when bg is inherited
    Given a _Background object having no background as background
     Then background.bg_element is None


  Scenario: Slide.copy_background_from copies an explicit background
    Given a Slide object with no explicit background
      And a source Slide object with an explicit fill background
     When I call slide.copy_background_from(source)
     Then the destination slide has an explicit p:bg element
      And the destination p:bg subtree matches the source


  Scenario: Slide.copy_background_from restores inheritance from an inherited source
    Given a Slide object with an explicit fill background
      And a source Slide object with no explicit background
     When I call slide.copy_background_from(source)
     Then the destination slide has no explicit p:bg element
