Feature: chart_data properties
  In order to compose data for a chart
  As a developer using python-pptx
  I need chart data transfer objects


  Scenario: CategoryChartData.categories getter
    Given a CategoryChartData object
     Then chart_data.categories is a Categories object


  @wip
  Scenario: CategoryChartData.categories setter
    Given a CategoryChartData object
     When I assign ['a', 'b', 'c'] to chart_data.categories
     Then [c.name for c in chart_data.categories] is ['a', 'b', 'c']


  Scenario: CategoryChartData.add_category()
    Given a CategoryChartData object
     Then chart_data.add_category(name) is a Category object
      And chart_data.categories[-1] is the category
