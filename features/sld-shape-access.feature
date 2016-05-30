Feature: Access shapes on a slide
  In order to interact with a shape
  As a developer using python-pptx
  I need ways to access a shape on a slide


  Scenario: Slide.shapes
     Given a slide having six shapes
      Then I can access the shape collection of the slide
       And the length of the shape collection is 6


  Scenario: SlideLayout.shapes
     Given a slide layout having three shapes
      Then I can access the shape collection of the slide layout
       And the length of the layout shape collection counts all its shapes


  Scenario: SlideMaster.shapes
     Given a slide master having two shapes
      Then I can access the shape collection of the slide master
       And the length of the master shape collection is 2


  Scenario: Shapes.__iter__ and Shapes.__getitem__
     Given a slide shape collection
      Then I can iterate over the shapes
       And I can access a shape by index
       And each slide shape is of the appropriate type


  Scenario: LayoutShapes.__iter__ and LayoutShapes.__getitem__
     Given a layout shape collection
      Then I can iterate over the layout shapes
       And I can access a layout shape by index
       And each shape is of the appropriate type


  Scenario: MasterShapes.__iter__ and __getitem__
     Given a master shape collection containing two shapes
      Then I can iterate over the master shapes
       And I can access a master shape by index


  Scenario: Shapes.index(shape)
     Given a slide shape collection
      Then the index of each shape matches its position in the sequence


  Scenario: Shapes.title
     Given a slide shape collection
      Then I can access the title placeholder


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
