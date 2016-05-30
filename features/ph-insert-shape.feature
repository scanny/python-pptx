Feature: Insert a shape into a placeholder
  In order to add a shape to a slide at a pre-defined location and size
  As a developer using python-pptx
  I need a way to insert a shape into a placeholder


  Scenario Outline: Insert an image into a picture placeholder
     Given an unpopulated picture placeholder shape
      When I call placeholder.insert_picture('<filename>')
      Then the return value is a PlaceholderPicture object
       And the placeholder contains the image
       And the <sides> crop is <value>

    Examples: Images inserted into a picture placeholder
      | filename           | sides          | value   |
      | monty-truth.png    | top and bottom | 0.23715 |
      | python-powered.png | left and right | 0.23333 |


  Scenario: Insert a table into a table placeholder
     Given an unpopulated table placeholder shape
      When I call placeholder.insert_table(rows=2, cols=3)
      Then the return value is a PlaceholderGraphicFrame object
       And the placeholder contains the table
       And the table has 2 rows and 3 columns


  Scenario: Insert a chart into a chart placeholder
     Given an unpopulated chart placeholder shape
      When I call placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data)
      Then the return value is a PlaceholderGraphicFrame object
       And the placeholder contains the chart
       And the chart is a pie chart
