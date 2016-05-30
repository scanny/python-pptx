Feature: Access slide layouts from a slide master
  In order to operate on the slide layouts of a presentation
  As a developer using python-pptx
  I need access to the slide layouts of a slide master


  Scenario: Access slide layout collection
     Given a slide master having two slide layouts
      Then I can access the slide layouts of the slide master
       And len(slide_layouts) is 2


  Scenario: Access slide layout in slide layout collection
     Given a slide layout collection containing two layouts
      Then I can iterate slide_layouts
       And I can access a slide layout by index
