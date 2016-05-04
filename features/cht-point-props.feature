Feature: Get and set point properties
  In order to customize the formatting of a data point
  As a developer using python-pptx
  I need a way to get and set point properties


  @wip
  Scenario: Get point data label
    Given a point
     Then point.data_label is a DataLabel object
