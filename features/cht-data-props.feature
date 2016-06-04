Feature: chart_data properties
  In order to compose data for a chart
  As a developer using python-pptx
  I need chart data transfer objects


  @wip
  Scenario: CategoryChartData.categories
    Given a CategoryChartData object
     Then chart_data.categories is a Categories object


  @wip
  Scenario: CategoryChartData.add_category()
    Given a CategoryChartData object
     Then chart_data.add_category(name) is a Category object
      And chart_data.categories[-1] is the category
