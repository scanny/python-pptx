Feature: Get and change line properties
  In order to format the outline of an auto shape
  As a developer using python-pptx
  I need access to the line properties of a shape

  Scenario: Access line format of a shape
     Given an autoshape
      Then I can access the line format of the shape

  @wip
  Scenario Outline: Get line fill type
     Given an autoshape having <outline type>
      Then the line fill type is <line fill type>

    Examples: Line fill types
      | outline type                | line fill type           |
      | an inherited outline format | None                     |
      | no outline                  | MSO_FILL_TYPE.BACKGROUND |
      | a solid outline             | MSO_FILL_TYPE.SOLID      |
