Feature: Category properties
  In order to characterize a chart category
  As a developer using python-pptx
  I need properties on the Category object


  Scenario Outline: Category.idx
    Given a Category object having idx value <idx>
     Then category.idx is <idx>

    Examples: idx cases
      | idx |
      |  1  |
      |  2  |


  Scenario Outline: Category.label
    Given a Category object having <label>
     Then category.label is <value>

    Examples: label cases
      | label       | value |
      | label 'Foo' | 'Foo' |
      | no label    | ''    |
