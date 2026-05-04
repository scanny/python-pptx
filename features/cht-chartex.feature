Feature: Office 2016+ extended (chartex) chart passthrough
  In order to preserve Office 2016+ extended chart types (funnel, treemap, sunburst,
  waterfall, histogram, box-and-whisker, map) on round-trip
  As a developer using python-pptx
  I need chartex-containing graphic frames to appear in slide.shapes and survive save


  Scenario: chartex graphic frames are surfaced in slide.shapes
    Given a presentation containing a chartex chart
     Then slide.shapes contains the chartex graphic-frame shapes


  Scenario Outline: chartex graphic frame has_chart and has_chartex
    Given the <which> chartex graphic-frame shape from a presentation
     Then shape.has_chart is True
      And shape.has_chartex is True
      And shape.chart_type is XL_CHART_TYPE.UNSUPPORTED_CHARTEX
      And shape.shape_type == MSO_SHAPE_TYPE.CHART
      And shape.chartex_type == "waterfall"

    Examples: chartex variants
      | which                    |
      | AlternateContent-wrapped |
      | direct                   |


  Scenario: chartex round-trip preserves the original XML
    Given a presentation containing a chartex chart
     When I save the presentation and reload it
     Then the reloaded slide still exposes the chartex graphic-frame shapes
      And the saved package still contains the chartex part and its AlternateContent wrapper
