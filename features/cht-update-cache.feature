Feature: Refresh cached chart values from the embedded workbook
  In order to reflect edits made to a chart's embedded Excel data in the
  chart that PowerPoint shows, without having to click "Refresh Data" in
  PowerPoint
  As a developer using python-pptx
  I need a way to rewrite the chart's cached values from the embedded xlsx


  Scenario: Chart.update_cached_values() rewrites numCache and strCache
    Given a chart whose embedded workbook has been updated externally
     When I call chart.update_cached_values()
     Then the cached series values match the embedded workbook
      And the cached category labels match the embedded workbook
      And the cached series name matches the embedded workbook
