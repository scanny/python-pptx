Feature: Slide master properties
  In order to interact with a slide master
  As a developer using python-pptx
  I need properties and methods on SlideMaster


  Scenario: SlideMaster.shapes
     Given a slide master
      Then slide_master.shapes is a MasterShapes object


  Scenario: SlideMaster.placeholders
     Given a slide master
      Then slide_master.placeholders is a MasterPlaceholders object


  Scenario: SlideMaster.slide_layouts
     Given a slide master
      Then slide_master.slide_layouts is a SlideLayouts object
