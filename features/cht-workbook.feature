Feature: Read/write access to a chart's embedded Excel workbook
  In order to read, duplicate, and surgically edit the data file that
  backs a chart without having to re-author its XML from scratch
  As a developer using python-pptx
  I need bytes-level access to the embedded workbook, a cross-part clone
  helper, and a targeted cell-write that keeps the cached chart values in
  sync with the workbook


  Scenario: Chart.workbook returns the bytes of the embedded xlsx
    Given a chart with an embedded workbook
     Then chart.workbook is the bytes of the embedded xlsx
      And chart.workbook reads the .xlsx zip magic

  Scenario: Chart.workbook accepts a bytes replacement
    Given a chart with an embedded workbook
     When I assign new bytes to chart.workbook
     Then chart.workbook returns the new bytes

  Scenario: update_embedded_xlsx_cell rewrites both the workbook and the cache
    Given a chart with an embedded workbook
     When I call update_embedded_xlsx_cell(chart, "Sheet1", "B2", 42.0)
     Then the workbook's Sheet1!B2 cell reads 42.0
      And the first series' first value reads 42.0

  Scenario: clone_embedded_xlsx duplicates the workbook into a second chart
    Given two charts in the same presentation
     When I call clone_embedded_xlsx(source_chart.part, target_chart.part)
     Then target_chart.workbook equals source_chart.workbook
      And the target chart's xlsx part is not the source's xlsx part
