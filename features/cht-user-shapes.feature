Feature: Chart user-shapes (issue #351)
  In order to inspect annotation shapes drawn on top of a chart
  As a developer using python-pptx
  I need read access to the ``c:userShapes`` relationship part


  Scenario: Chart.user_shapes is None when no chart-drawing rel exists
    Given a newly-authored bar chart
     Then chart.user_shapes is None
      And chart.has_user_shapes is False


  Scenario: Chart.user_shapes returns the related ChartDrawingPart
    Given a chart with two user-shape anchors attached
     Then chart.user_shapes is a ChartDrawingPart object
      And chart.user_shapes.anchor_count is 2
      And chart.has_user_shapes is True


  Scenario: ChartDrawingPart.iter_anchor_elements yields each anchor
    Given a chart with two user-shape anchors attached
     Then iterating chart.user_shapes yields 2 anchor elements
      And the anchor tags are cdr:relSizeAnchor and cdr:absSizeAnchor
