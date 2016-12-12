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


  Scenario Outline: Categories.levels
    Given a Categories object having <depth> category levels
     Then categories.levels contains <len> CategoryLevel objects

    Examples: hierarchical category depths
      | depth | len |
      |   0   |  0  |
      |   1   |  0  |
      |   2   |  2  |
      |   3   |  3  |


  Scenario Outline: Categories.flattened_labels
    Given a Categories object having <leafs> categories and <levels> levels
     Then categories.flattened_labels is a tuple of <leafs> tuples
      And each label tuple contains <levels> labels

    Examples: hierarchical category combinations
      | leafs | levels |
      |   3   |    1   |
      |   8   |    3   |
      |   0   |    0   |
      |   4   |    2   |


  Scenario: CategoryLevel.__len__()
    Given a CategoryLevel object containing 4 categories
     Then len(category_level) is 4


  Scenario: CategoryLevel.__getitem__()
    Given a CategoryLevel object containing 4 categories
     Then category_level[2] is a Category object


  Scenario: CategoryLevel.__iter__()
    Given a CategoryLevel object containing 4 categories
     Then iterating category_level produces 4 Category objects
