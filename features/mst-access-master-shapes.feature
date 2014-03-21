Feature: Access shapes on a slide master
  In order to operate on the shapes of a slide master
  As a developer using python-pptx
  I need access to the shape collection of a slide master

  Scenario: Access shape collection of a slide master
     Given a slide master having two shapes
      Then I can access the shape collection of the slide master
       And the length of the master shape collection is 2

  Scenario: Access shape in master shape collection
     Given a master shape collection containing two shapes
      Then I can iterate over the master shapes
       And I can access a master shape by index
