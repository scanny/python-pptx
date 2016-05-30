Feature: Access chart object
  In order to operate on a chart in a presentation
  As a developer using python-pptx
  I need a way to get a chart from its graphic frame container


  Scenario Outline: Identify a shape containing a chart
    Given a <shape type>
     Then the shape <has?> a chart

    Examples: Shape types
      | shape type                       | has?          |
      | shape                            | does not have |
      | graphic frame containing a chart | has           |
      | picture                          | does not have |
      | graphic frame containing a table | does not have |
      | group shape                      | does not have |
      | connector                        | does not have |


  Scenario: Access chart object from graphic frame shape
    Given a graphic frame containing a chart
     When I get the chart from its graphic frame
     Then the chart is a Chart object
