Feature: Access a category
  In order to operate on an individual category
  As a developer using python-pptx
  I need sequence operations on the Categories object


  Scenario: Categories.__len__()
    Given a Categories object containing 3 categories
     Then len(categories) is 3


  Scenario: Categories.__getitem__()
    Given a Categories object containing 3 categories
     Then categories[2] is a Category object


  Scenario: Categories.__iter__()
    Given a Categories object containing 3 categories
     Then iterating categories produces 3 Category objects
      And list(categories) == ['Foo', '', 'Baz']


  Scenario Outline: Categories.depth
    Given a Categories object having <depth> category levels
     Then categories.depth is <depth>

    Examples: hierarchical category depths
      | depth |
      |   0   |
      |   1   |
      |   2   |
      |   3   |
