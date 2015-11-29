Feature: Access axis title
  In order to customize the title for an axis
  As a developer using python-pptx
  I need a way to access the axis title

  Scenario Outline: Determine presence of title
    Given an axis <having-or-not> a title
     Then axis.has_title is <expected-value>

    Examples: axis.has_title states
      | having-or-not | expected-value |
      | having        | True           |
      | not having    | False          |


  Scenario Outline: Adding a title
    Given an axis <having-or-not> a title
     When I assign <value> to axis.has_title
     Then axis.has_title is <expected-value>

    Examples: axis.has_title states
      | having-or-not | value | expected-value |
      | having        | False | False          |
      | not having    | True  | True           |
      | having        | True  | True           |
      | not having    | False | False          |


  Scenario: Access title object of axis
    Given an axis having a title
     Then axis.title is a title object
