Feature: Access a shape on a slide
  In order to interact with a shape
  As a developer using python-pptx
  I need ways to access a shape on a slide


  Scenario: _BaseShapes.turbo_add_enabled default
    Given a _BaseShapes object as shapes
     Then shapes.turbo_add_enabled is False


  Scenario: _BaseShapes.turbo_add_enabled turned on
    Given a _BaseShapes object as shapes
     When I assign True to shapes.turbo_add_enabled
      And I add 100 shapes
     Then len(shapes) == 100


  Scenario: GroupShapes is a sequence
    Given a GroupShapes object of length 3 as shapes
     Then len(shapes) == 3
      And shapes[1] is a Shape object
      And iterating shapes produces 3 objects that subclass BaseShape
      And shapes.index(shape) for each shape matches its sequence position


  Scenario: GroupShapes.add_chart()
    Given a GroupShapes object as shapes
     When I assign shapes.add_chart() to shape
     Then shape is a GraphicFrame object
      And shapes[-1] == shape


  Scenario: GroupShapes.add_connector()
    Given a GroupShapes object as shapes
     When I assign shapes.add_connector() to shape
     Then shape is a Connector object
      And shapes[-1] == shape


  Scenario: GroupShapes.add_group_shape()
    Given a GroupShapes object as shapes
     When I assign shapes.add_group_shape() to shape
     Then shape is a GroupShape object
      And shapes[-1] == shape


  Scenario: GroupShapes.add_picture()
    Given a GroupShapes object as shapes
     When I assign shapes.add_picture() to shape
     Then shape is a Picture object
      And shapes[-1] == shape


  Scenario: GroupShapes.add_shape()
    Given a GroupShapes object as shapes
     When I assign shapes.add_shape() to shape
     Then shape is a Shape object
      And shapes[-1] == shape


  Scenario: GroupShapes.add_textbox()
    Given a GroupShapes object as shapes
     When I assign shapes.add_textbox() to shape
     Then shape is a Shape object
      And shapes[-1] == shape


  Scenario: GroupShapes.build_freeform()
    Given a GroupShapes object as shapes
     When I assign shapes.build_freeform() to builder
     Then builder is a FreeformBuilder object


  Scenario: LayoutPlaceholders is a sequence
    Given a LayoutPlaceholders object of length 2 as shapes
     Then len(shapes) == 2
      And shapes[1] is a LayoutPlaceholder object
      And iterating shapes produces 2 objects of type LayoutPlaceholder
      And shapes.get(idx=10) is the body placeholder


  Scenario: LayoutShapes is a sequence
    Given a LayoutShapes object of length 3 as shapes
     Then len(shapes) == 3
      And shapes[1] is a LayoutPlaceholder object
      And iterating shapes produces 3 objects that subclass BaseShape


  Scenario: MasterPlaceholders is a sequence
    Given a MasterPlaceholders object of length 2 as shapes
     Then len(shapes) == 2
      And shapes[1] is a MasterPlaceholder object
      And iterating shapes produces 2 objects of type MasterPlaceholder
      And shapes.get(PP_PLACEHOLDER.BODY) is the body placeholder


  Scenario: MasterShapes is a sequence
    Given a MasterShapes object of length 2 as shapes
     Then len(shapes) == 2
      And shapes[1] is a Picture object
      And iterating shapes produces 2 objects that subclass BaseShape


  Scenario: SlideShapes is a sequence
    Given a SlideShapes object of length 6 shapes as shapes
     Then len(shapes) == 6
      And shapes[4] is a GraphicFrame object
      And iterating shapes produces 6 objects that subclass BaseShape
      And shapes.index(shape) for each shape matches its sequence position


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


  Scenario: SlideShapes.add_connector()
    Given a SlideShapes object as shapes
     When I call shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4)
     Then connector is a Connector object
      And connector.begin_x == 1
      And connector.begin_y == 2
      And connector.end_x == 3
      And connector.end_y == 4


  Scenario: SlideShapes.add_group_shape()
    Given a SlideShapes object as shapes
     When I assign shapes.add_group_shape() to shape
     Then shape is a GroupShape object
      And shapes[-1] == shape


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


  Scenario Outline: SlideShapes.add_ole_object()
    Given a SlideShapes object as shapes
      And a <prog-id> file as ole_object_file
     When I assign shapes.add_ole_object(ole_object_file) to shape
      And I assign shape.ole_format to ole_format
     Then shapes[-1] == shape
      And shape is a GraphicFrame object
      And shape.ole_format is an _OleFormat object
      And shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT
      And ole_format.blob matches ole_object_file byte-for-byte
      And ole_format.prog_id == <progId>
      And ole_format.show_as_icon is True

    Examples: add_ole_object() variations
      | prog-id | progId               |
      | DOCX    | "Word.Document.12"   |
      | PPTX    | "PowerPoint.Show.12" |
      | XLSX    | "Excel.Sheet.12"     |


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


  Scenario: SlideShapes.add_shape()
    Given a SlideShapes object as shapes
     When I assign shapes.add_shape() to shape
     Then shape is a Shape object
      And shapes[-1] == shape


  Scenario: SlideShapes.add_table()
    Given a blank slide
     When I add a table to the slide's shape collection
      And I save the presentation
     Then the table appears in the slide


  Scenario: SlideShapes.add_textbox()
    Given a SlideShapes object as shapes
     When I assign shapes.add_textbox() to shape
     Then shape is a Shape object
      And shapes[-1] == shape


  Scenario: SlideShapes.build_freeform() (no parameters)
    Given a SlideShapes object as shapes
     When I assign shapes.build_freeform() to builder
     Then builder is a FreeformBuilder object
      And (builder._start_x, builder._start_y) is (0, 0)
      And (builder._x_scale, builder._y_scale) is (1.0, 1.0)


  Scenario: SlideShapes.build_freeform() (with pen start position)
    Given a SlideShapes object as shapes
    When I assign shapes.build_freeform(start_x=25, start_y=125) to builder
     Then builder is a FreeformBuilder object
      And (builder._start_x, builder._start_y) is (25, 125)


  Scenario: SlideShapes.build_freeform() (square scaling)
    Given a SlideShapes object as shapes
     When I assign shapes.build_freeform(scale=100.0) to builder
     Then builder is a FreeformBuilder object
      And (builder._x_scale, builder._y_scale) is (100.0, 100.0)


  Scenario: SlideShapes.build_freeform() (rectangular scaling)
    Given a SlideShapes object as shapes
     When I assign shapes.build_freeform(scale=(200.0, 100.0)) to builder
     Then builder is a FreeformBuilder object
      And (builder._x_scale, builder._y_scale) is (200.0, 100.0)


  Scenario: SlideShapes.title
    Given a SlideShapes object as shapes
     Then shapes.title is the title placeholder


  Scenario: SlidePlaceholders is a sequence
    Given a SlidePlaceholders object of length 2 as shapes
     Then len(shapes) == 2
      And shapes[10] is a SlidePlaceholder object
      And iterating shapes produces 2 objects of type SlidePlaceholder


  Scenario Outline: Access unpopulated placeholder shape
    Given a slide with an unpopulated <type> placeholder
     Then slide.shapes[0] is a <cls> proxy object for that placeholder

    Examples: Unpopulated placeholder shapes
      | type     | cls                |
      | picture  | PicturePlaceholder |
      | clip art | PicturePlaceholder |
      | table    | TablePlaceholder   |
      | chart    | ChartPlaceholder   |


  Scenario Outline: Access populated placeholder shape
    Given a slide with a <type> placeholder populated with <content>
     Then slide.shapes[0] is a <cls> proxy object for that placeholder

    Examples: Populated placeholder shapes
      | type     | content  | cls                     |
      | picture  | an image | PlaceholderPicture      |
      | clip art | an image | PlaceholderPicture      |
      | table    | a table  | PlaceholderGraphicFrame |
      | chart    | a chart  | PlaceholderGraphicFrame |


  Scenario Outline: ShapeFactory contructs appropriate proxy object
    Given a SlideShapes object having a <type> shape at offset <idx>
     Then shapes[<idx>] is a <cls> object

    Examples: Shape object types
      | type      | idx | cls        |
      | connector |  0  | Connector  |
      | picture   |  1  | Picture    |
      | rectangle |  2  | Shape      |
      | group     |  3  | GroupShape |
