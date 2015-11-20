Feature: Insert a chart into a placeholder
  In order to add a chart to a slide at a pre-defined location and size
  As a developer using python-pptx
  I need a way to insert a chart into a placeholder


  Scenario: Insert a chart into a chart placeholder
     Given an unpopulated chart placeholder shape
      When I call placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data)
      Then the return value is a PlaceholderGraphicFrame object
       And the placeholder contains the chart
       And the chart is a pie chart
