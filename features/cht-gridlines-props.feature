Feature: Gridlines properties
  To customize the formatting of chart gridlines
  As a developer using python-pptx
  I need read/write properties on major and minor gridlines objects


  Scenario: Get gridlines.format
    Given a major gridlines
     Then gridlines.format is a ChartFormat object
      And gridlines.format.fill is a FillFormat object
      And gridlines.format.line is a LineFormat object
