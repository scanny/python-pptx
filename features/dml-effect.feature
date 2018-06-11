Feature: ShadowFormat
  In order to adjust a shadow effect
  As a developer using python-pptx
  I need properties and methods on ShadowFormat


  Scenario Outline: ShadowFormat.inherit getter
    Given a ShadowFormat object that <inherits-or-not> as shadow
     Then shadow.inherit is <value>

    Examples: shadow inheritance cases
      | inherits-or-not  | value |
      | inherits         | True  |
      | does not inherit | False |


  Scenario Outline: ShadowFormat.inherit setter
    Given a ShadowFormat object that <inherits-or-not> as shadow
     When I assign <new-value> to shadow.inherit
     Then shadow.inherit is <value>

    Examples: shadow inheritance assignment cases
      | inherits-or-not  | new-value | value |
      | inherits         | False     | False |
      | does not inherit | True      | True  |
      | inherits         | None      | False |
      | inherits         | True      | True  |
      | does not inherit | None      | False |
      | does not inherit | False     | False |
