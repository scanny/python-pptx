Feature: Slide layout properties
  In order to interact with a slide layout
  As a developer using python-pptx
  I need properties and methods on SlideLayout


  Scenario: SlideLayout.shapes
     Given a slide layout
      Then slide_layout.shapes is a LayoutShapes object


  Scenario: SlideLayout.placeholders
     Given a slide layout
      Then slide_layout.placeholders is a LayoutPlaceholders object


  Scenario: SlideLayout.slide_master
     Given a slide layout
      Then slide_layout.slide_master is a SlideMaster object
