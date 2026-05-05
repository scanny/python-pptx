Feature: Cross-slide chart copy
  In order to duplicate a chart onto another slide or presentation without
  re-authoring its XML and without sharing its embedded workbook part
  As a developer using python-pptx
  I need a Chart.clone_to() method that targets a shape tree and copies the
  chart part, all its relationships, and a distinct copy of the embedded
  Excel workbook


  Scenario: Chart.clone_to copies a chart onto a different slide in the same presentation
    Given a chart with an embedded workbook on slide 1
     When I call chart.clone_to(slide_2.shapes, x, y, cx, cy)
     Then slide 2 has a chart shape at (x, y) sized (cx, cy)
      And the cloned chart has a distinct chart part
      And the cloned chart has a distinct embedded xlsx part
      And the cloned chart's workbook bytes equal the source workbook bytes
      And the round-tripped presentation still has two charts

  Scenario: Chart.clone_to copies a chart into a different presentation
    Given a chart in presentation A
      And an empty presentation B with one slide
     When I call chart.clone_to(slide_B.shapes, x, y, cx, cy)
     Then presentation B's slide has a chart
      And that chart's chart part belongs to presentation B's package
      And saving presentation B alone round-trips the chart

  Scenario: SlideShapes.add_chart_from adds a full-fidelity clone of a source chart
    Given a chart with an embedded workbook on slide 1
     When I call slide_2.shapes.add_chart_from(source_chart, x, y, cx, cy)
     Then slide 2 has a chart shape at (x, y) sized (cx, cy)
      And the cloned chart has a distinct chart part
      And the cloned chart has a distinct embedded xlsx part

  Scenario: SlideShapes.add_chart_from defaults width and height when omitted
    Given a chart with an embedded workbook on slide 1
     When I call slide_2.shapes.add_chart_from(source_chart, x, y) with no size
     Then slide 2 has a chart shape at (x, y) sized 5 by 3 inches
      And the cloned chart has a distinct chart part
