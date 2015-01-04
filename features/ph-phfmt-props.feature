Feature: Get placeholder format properties
  In order to identify and characterize placeholder shapes
  As a developer using python-pptx
  I need a set of placeholder format properties


  Scenario Outline: Get placeholder idx
     Given a known <type> placeholder shape
      Then placeholder_format.idx is <value>

    Examples: Unpopulated placeholder types
      | type      | value |
      | title     |   0   |
      | content   |  10   |


  Scenario Outline: Get placeholder type
     Given a known <type> placeholder shape
      Then placeholder_format.type is <value>

    Examples: Unpopulated placeholder types
      | type      | value                     |
      | title     | PP_PLACEHOLDER.TITLE      |
      | content   | PP_PLACEHOLDER.OBJECT     |
      | text      | PP_PLACEHOLDER.BODY       |
