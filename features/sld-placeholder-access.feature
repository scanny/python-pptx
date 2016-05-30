Feature: Access slide placeholder
  In order to interact with a placeholder shape
  As a developer using python-pptx
  I need ways to access a placeholder on a slide


  Scenario: Slide.placeholders
     Given a slide having two placeholders
      Then I can access the placeholder collection of the slide
       And the length of the slide placeholder collection is 2


  Scenario: SlideLayout.placeholders
     Given a slide layout having two placeholders
      Then I can access the placeholder collection of the slide layout
       And the length of the layout placeholder collection is 2


  Scenario: SlideMaster.placeholders
     Given a slide master having two placeholders
      Then I can access the placeholder collection of the slide master
       And the length of the master placeholder collection is 2


  Scenario: SlidePlaceholders.__iter__ and __getitem__
     Given a slide placeholder collection
      Then I can iterate over the slide placeholders
       And I can access a slide placeholder by index


  Scenario: LayoutPlaceholders.__iter__ and __getitem__
     Given a layout placeholder collection
      Then I can iterate over the layout placeholders
       And I can access a layout placeholder by index
       And I can access a layout placeholder by idx value


  Scenario: MasterPlaceholders.__iter__ and __getitem__
     Given a master placeholder collection
      Then I can iterate over the master placeholders
       And I can access a master placeholder by index
       And I can access a master placeholder by type
