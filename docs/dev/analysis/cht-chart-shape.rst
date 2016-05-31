
Chart Shape
===========

A chart is not actually a shape. It is a graphical object held inside
a graphics frame. The graphics frame is a shape and the chart must be
retrieved from it.


Protocol
--------

::

    >>> shape = shapes.add_textbox(0, 0, 0, 0)
    >>> shape.has_chart
    False
    >>> shape.chart
    ...
    AttributeError: 'Shape' object has no attribute 'chart'
    >>> shape = shapes.add_chart(style, type, x, y, cx, cy)
    >>> type(shape)
    <class 'pptx.shapes.graphfrm.GraphicFrame'>
    >>> shape.has_chart
    True
    >>> shape.chart
    <pptx.parts.chart.ChartPart object at 0x108c0e290>


Acceptance tests
----------------

::

    Feature: Access chart object
      In order to operate on a chart in a presentation
      As a developer using python-pptx
      I need a way to find and access a chart

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
         Then the chart is a ChartPart object
