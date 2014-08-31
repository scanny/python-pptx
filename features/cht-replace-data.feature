Feature: Replace chart data
  In order to update the data for a chart while retaining its formatting
  As a developer using python-pptx
  I need a way to replace the data of a chart

  Scenario Outline: Replace chart data
    Given a chart of size and type <spec>
     When I replace its data with <cats> categories and <sers> series
     Then the chart has <cats> categories
      And the chart has <sers> series
      And each series has <cats> values
      And each series has a new name
      And the chart has new chart data

    Examples: Replacement chart data
      | spec                 | cats | sers |
      | 2x2 Clustered Bar    |   3  |   3  |
      | 2x2 100% Stacked Bar |   3  |   3  |
      | 2x2 Clustered Column |   3  |   3  |
      | 4x3 Line             |   3  |   2  |
      | 3x1 Pie              |   5  |   1  |
