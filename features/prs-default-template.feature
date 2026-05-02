Feature: Default presentation template is provided
  In order to get started on a presentation with a minimum of preparation
  As a developer using python-pptx
  I would like the option to start from a default template

  Scenario: Create a minimal presentation from the default template
     Given an initialized pptx environment
      When I construct a Presentation instance with no path argument
      Then I receive a presentation based on the default template


  Scenario Outline: Select a widescreen or standard template via pptx_format
    Given an initialized pptx environment
     When I construct a Presentation with pptx_format=<preset>
     Then the new presentation has slide width <cx> and height <cy>
      And len(prs.slide_masters) is 1
      And len(prs.slide_layouts) is 11

    Examples: Aspect-ratio presets
      | preset       |    cx    |   cy    |
      | "4x3"        | 9144000  | 6858000 |
      | "standard"   | 9144000  | 6858000 |
      | "16x9"       | 12192000 | 6858000 |
      | "widescreen" | 12192000 | 6858000 |
