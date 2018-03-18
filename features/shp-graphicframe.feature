Feature: GraphicFrame properties and methods
  In order to operate on a GraphicFrame shape
  As a developer using python-pptx
  I need properties and methods on GraphicFrame


  Scenario Outline: GraphicFrame.has_chart
    Given a GraphicFrame object containing a <graphical-object> as shape
     Then shape.has_chart is <value>

    Examples: Shape types
      | graphical-object | value |
      | chart            | True  |
      | table            | False |


  Scenario: GraphicFrame.chart
    Given a GraphicFrame object containing a chart as shape
     Then shape.chart is a Chart object
