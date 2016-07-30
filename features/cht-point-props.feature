Feature: Get and set point properties
  In order to customize the formatting of a data point
  As a developer using python-pptx
  I need a way to get and set point properties


  Scenario: Get Point.data_label
    Given a point
     Then point.data_label is a DataLabel object


  Scenario: Get Point.format
    Given a point
     Then point.format is a ChartFormat object
      And point.format.fill is a FillFormat object
      And point.format.line is a LineFormat object


  Scenario: Get Point.marker
    Given a point
     Then point.marker is a Marker object
