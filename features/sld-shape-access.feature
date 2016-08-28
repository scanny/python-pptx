Feature: Access a shape on a slide
  In order to interact with a shape
  As a developer using python-pptx
  I need ways to access a shape on a slide


  Scenario: SlideShapes.__len__()
    Given a SlideShapes object containing 6 shapes
     Then len(shapes) is 6


  Scenario: SlidePlaceholders.__len__()
    Given a SlidePlaceholders object containing 2 placeholders
     Then len(placeholders) is 2


  Scenario: LayoutShapes.__len__()
    Given a LayoutShapes object containing 3 shapes
     Then len(layout_shapes) is 3


  Scenario: LayoutPlaceholders.__len__()
    Given a LayoutPlaceholders object containing 2 placeholders
     Then len(layout_placeholders) is 2


  Scenario: MasterShapes.__len__()
    Given a MasterShapes object containing 2 shapes
     Then len(master_shapes) is 2


  Scenario: MasterPlaceholders.__len__()
    Given a MasterPlaceholders object containing 2 placeholders
     Then len(master_placeholders) is 2


  Scenario: SlideShapes.__getitem__()
    Given a SlideShapes object containing 6 shapes
     Then shapes[4] is a GraphicFrame object


  Scenario: SlidePlaceholders.__getitem__()
    Given a SlidePlaceholders object containing 2 placeholders
     Then placeholders[10] is a SlidePlaceholder object


  Scenario: LayoutShapes.__getitem__()
    Given a LayoutShapes object containing 3 shapes
     Then layout_shapes[1] is a LayoutPlaceholder object


  Scenario: LayoutPlaceholders.__getitem__
    Given a LayoutPlaceholders object containing 2 placeholders
     Then layout_placeholders[1] is a LayoutPlaceholder object


  Scenario: MasterShapes.__getitem__()
    Given a MasterShapes object containing 2 shapes
     Then master_shapes[1] is a Picture object


  Scenario: MasterPlaceholders.__getitem__
    Given a MasterPlaceholders object containing 2 placeholders
     Then master_placeholders[1] is a MasterPlaceholder object


  Scenario: SlideShapes.__iter__()
    Given a SlideShapes object containing 6 shapes
     Then iterating shapes produces 6 BaseShape objects


  Scenario: SlidePlaceholders.__iter__()
    Given a SlidePlaceholders object containing 2 placeholders
     Then iterating placeholders produces 2 SlidePlaceholder objects


  Scenario: LayoutShapes.__iter__()
    Given a LayoutShapes object containing 3 shapes
     Then iterating layout_shapes produces 3 BaseShape objects


  Scenario: LayoutPlaceholders.__iter__()
    Given a LayoutPlaceholders object containing 2 placeholders
     Then iterating layout_placeholders produces 2 LayoutPlaceholder objects


  Scenario: MasterShapes.__iter__()
    Given a MasterShapes object containing 2 shapes
     Then iterating master_shapes produces 2 BaseShape objects


  Scenario: MasterPlaceholders.__iter__()
    Given a MasterPlaceholders object containing 2 placeholders
     Then iterating master_placeholders produces 2 MasterPlaceholder objects


  Scenario: Shapes.index(shape)
    Given a SlideShapes object containing 6 shapes
     Then the index of each shape matches its position in the sequence


  Scenario: Shapes.title
    Given a SlideShapes object containing 6 shapes
     Then shapes.title is the title placeholder


  Scenario: LayoutPlaceholders.get()
    Given a LayoutPlaceholders object containing 2 placeholders
     Then layout_placeholders.get(idx=10) is the body placeholder


  Scenario: MasterPlaceholders.get()
    Given a MasterPlaceholders object containing 2 placeholders
     Then master_placeholders.get(PP_PLACEHOLDER.BODY) is the body placeholder


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


  Scenario Outline: Access a shape
    Given a SlideShapes object having a <type> shape at offset <idx>
     Then shapes[<idx>] is a <cls> object

    Examples: Shape object types
      | type      | idx | cls       |
      | connector |  0  | Connector |
      | picture   |  1  | Picture   |
      | rectangle |  2  | Shape     |
