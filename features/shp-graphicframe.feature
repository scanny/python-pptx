Feature: GraphicFrame properties and methods
  In order to operate on a GraphicFrame shape
  As a developer using python-pptx
  I need properties and methods on GraphicFrame


  Scenario: GraphicFrame.chart
    Given a GraphicFrame object containing a chart as shape
     Then shape.chart is a Chart object


  Scenario Outline: GraphicFrame.has_chart
    Given a GraphicFrame object containing a <graphical-object> as shape
     Then shape.has_chart is <value>

    Examples: Shape types
      | graphical-object | value |
      | chart            | True  |
      | table            | False |


  Scenario: GraphicFrame.ole_format
    Given a GraphicFrame object containing an OLE object as shape
     Then shape.ole_format is an _OleFormat object


  Scenario Outline: GraphicFrame.shape_type
    Given a GraphicFrame object containing <graphical-object> as shape
     Then shape.shape_type == <expected-value>

    Examples: Shape types
      | graphical-object | expected-value                     |
      | an OLE object    | MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT |
      | a chart          | MSO_SHAPE_TYPE.CHART               |
      | a table          | MSO_SHAPE_TYPE.TABLE               |


  Scenario: _OleFormat.blob
    Given an _OleFormat object for an embedded XLSX as ole_format
     Then len(ole_format.blob) == 8287


  Scenario: _OleFormat.prog_id
    Given an _OleFormat object for an embedded XLSX as ole_format
     Then ole_format.prog_id == "Excel.Sheet.12"


  Scenario: _OleFormat.show_as_icon
    Given an _OleFormat object for an OLE object as ole_format
     Then ole_format.show_as_icon is True
