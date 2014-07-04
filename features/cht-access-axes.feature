Feature: Access chart axes
  In order to customize axis formatting on a chart
  As a developer using python-pptx
  I need a way to access the category and value axes of a chart


  Scenario: Access category axis
    Given a bar chart
     Then I can access the chart category axis


  Scenario: Access value axis
    Given a bar chart
     Then I can access the chart value axis
