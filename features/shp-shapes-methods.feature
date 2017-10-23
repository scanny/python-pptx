Feature: Shape collection methods
  In order to add a shape to a shape collection
  As a developer using python-pptx
  I need a set of methods on Shapes objects


  Scenario: SlideShapes.add_connector()
    Given a SlideShapes object as shapes
     When I call shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4)
     Then connector is a Connector object
      And connector.begin_x == 1
      And connector.begin_y == 2
      And connector.end_x == 3
      And connector.end_y == 4


  Scenario: SlideShapes.add_shape()
    Given a blank slide
     When I add an auto shape to the slide's shape collection
      And I save the presentation
     Then the auto shape appears in the slide


  Scenario Outline: SlideShapes.add_chart() (category chart)
    Given a blank slide
     When I add a <type> chart with <cats> categories and <sers> series
     Then chart.chart_type is <chart-type>
      And len(plot.categories) is <cats>
      And len(chart.series) is <sers>
      And len(series.values) is <cats> for each series
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | type                      | chart-type               | cats | sers |
      | Area                      | AREA                     |   3  |   3  |
      | Stacked Area              | AREA_STACKED             |   3  |   3  |
      | 100% Stacked Area         | AREA_STACKED_100         |   3  |   3  |
      | Clustered Bar             | BAR_CLUSTERED            |   3  |   3  |
      | Stacked Bar               | BAR_STACKED              |   3  |   3  |
      | 100% Stacked Bar          | BAR_STACKED_100          |   3  |   3  |
      | Clustered Column          | COLUMN_CLUSTERED         |   3  |   3  |
      | Stacked Column            | COLUMN_STACKED           |   3  |   3  |
      | 100% Stacked Column       | COLUMN_STACKED_100       |   3  |   3  |
      | Doughnut                  | DOUGHNUT                 |   5  |   1  |
      | Exploded Doughnut         | DOUGHNUT_EXPLODED        |   5  |   1  |
      | Line                      | LINE                     |   3  |   2  |
      | Line with Markers         | LINE_MARKERS             |   3  |   2  |
      | Line Markers Stacked      | LINE_MARKERS_STACKED     |   3  |   2  |
      | 100% Line Markers Stacked | LINE_MARKERS_STACKED_100 |   3  |   2  |
      | Line Stacked              | LINE_STACKED             |   3  |   2  |
      | 100% Line Stacked         | LINE_STACKED_100         |   3  |   2  |
      | Pie                       | PIE                      |   5  |   1  |
      | Exploded Pie              | PIE_EXPLODED             |   5  |   1  |
      | Radar                     | RADAR                    |   5  |   2  |
      | Filled Radar              | RADAR_FILLED             |   5  |   2  |
      | Radar with markers        | RADAR_MARKERS            |   5  |   2  |


  Scenario Outline: SlideShapes.add_chart() (category chart with date axis)
    Given a SlideShapes object as shapes
      And a CategoryChartData object having date categories
     When I call shapes.add_chart(<chart-type>, chart_data)
     Then chart.category_axis is a DateAxis object

    Examples: Chart specs
      | chart-type       |
      | AREA             |
      | BAR_CLUSTERED    |
      | COLUMN_CLUSTERED |
      | LINE             |


  Scenario: SlideShapes.add_chart() (multi-level category chart)
    Given a blank slide
     When I add a Clustered bar chart with multi-level categories
     Then chart.chart_type is BAR_CLUSTERED
      And len(plot.categories) is 4
      And the chart has an Excel data worksheet


  Scenario Outline: SlideShapes.add_chart() (XY chart)
    Given a blank slide
     When I add an <chart-type> chart having 2 series of 3 points each
     Then chart.chart_type is <chart-type>
      And len(chart.series) is 2
      And len(series.values) is 3 for each series
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | chart-type                   |
      | XY_SCATTER                   |
      | XY_SCATTER_LINES             |
      | XY_SCATTER_LINES_NO_MARKERS  |
      | XY_SCATTER_SMOOTH            |
      | XY_SCATTER_SMOOTH_NO_MARKERS |


  Scenario Outline: SlideShapes.add_chart() (bubble chart)
    Given a blank slide
     When I add a <chart-type> chart having 2 series of 3 points each
     Then chart.chart_type is <chart-type>
      And len(chart.series) is 2
      And len(series.values) is 3 for each series
      And the chart has an Excel data worksheet

    Examples: Chart specs
      | chart-type            |
      | BUBBLE                |
      | BUBBLE_THREE_D_EFFECT |


  Scenario Outline: SlideShapes.add_picture() (using filename)
    Given a blank slide
     When I add the image <filename> using shapes.add_picture()
      And I save the presentation
     Then a <ext> image part appears in the pptx file
      And the picture appears in the slide

    Examples: image files of supported types
      | filename         | ext  |
      | sonic.gif        | gif  |
      | python-icon.jpeg | jpg  |
      | monty-truth.png  | png  |
      | 72-dpi.tiff      | tiff |
      | CVS_LOGO.WMF     | wmf  |
      | pic.emf          | wmf  |
      | python.bmp       | bmp  |


  Scenario Outline: SlideShapes.add_picture() (using file-like object)
    Given a blank slide
     When I add the stream image <filename> using shapes.add_picture()
      And I save the presentation
     Then a <ext> image part appears in the pptx file
      And the picture appears in the slide

    Examples: image files of supported types
      | filename         | ext  |
      | sonic.gif        | gif  |
      | python-icon.jpeg | jpg  |
      | monty-truth.png  | png  |
      | 72-dpi.tiff      | tiff |
      | CVS_LOGO.WMF     | wmf  |
      | pic.emf          | wmf  |
      | python.bmp       | bmp  |


  Scenario Outline: SlideShapes.add_movie()
    Given a SlideShapes object containing <a-or-no> movies
     When I call shapes.add_movie(file, x, y, cx, cy, poster_frame)
      And I save the presentation
     Then movie is a Movie object
      And movie.left, movie.top == x, y
      And movie.width, movie.height == cx, cy
      And movie.poster_frame is the same image as poster_frame

    Examples: add_movie() preconditions
      | a-or-no     |
      | one or more |
      | no          |


  Scenario: SlideShapes.add_table()
    Given a blank slide
     When I add a table to the slide's shape collection
      And I save the presentation
     Then the table appears in the slide


  Scenario: SlideShapes.add_textbox()
    Given a blank slide
     When I add a text box to the slide's shape collection
      And I save the presentation
     Then the text box appears in the slide


  @wip
  Scenario: SlideShapes.build_freeform() (no parameters)
    Given a SlideShapes object as shapes
     When I assign shapes.build_freeform() to builder
     Then builder is a FreeformBuilder object
      And (builder._start_x, builder._start_y) is (0, 0)
      And (builder._x_scale, builder._y_scale) is (1.0, 1.0)


  @wip
  Scenario: SlideShapes.build_freeform() (with pen start position)
    Given a SlideShapes object as shapes
    When I assign shapes.build_freeform(start_x=25, start_y=125) to builder
     Then builder is a FreeformBuilder object
      And (builder._start_x, builder._start_y) is (25, 125)


  @wip
  Scenario: SlideShapes.build_freeform() (square scaling)
    Given a SlideShapes object as shapes
     When I assign shapes.build_freeform(scale=100.0) to builder
     Then builder is a FreeformBuilder object
      And (builder._x_scale, builder._y_scale) is (100.0, 100.0)


  @wip
  Scenario: SlideShapes.build_freeform() (rectangular scaling)
    Given a SlideShapes object as shapes
     When I assign shapes.build_freeform(scale=(200.0, 100.0)) to builder
     Then builder is a FreeformBuilder object
      And (builder._x_scale, builder._y_scale) is (200.0, 100.0)
