Feature: Combo chart (multi-plot chart) support
  In order to build a chart that overlays two chart types
  As a developer using python-pptx
  I need a way to add an additional plot to an existing chart


  Scenario: Add a line plot on top of a column chart (issue #338)
    Given a column chart and new line-plot chart data
     When I call chart.add_plot(XL_CHART_TYPE.LINE, line_data)
     Then chart.plots has length 2
      And the second plot is a LinePlot
      And the line plot's axId elements match the column plot's
      And the line plot's series use inline numLit and strLit data
