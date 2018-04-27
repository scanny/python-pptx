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
